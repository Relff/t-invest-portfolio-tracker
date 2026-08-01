/**
 * ldv.js — Льгота на долгосрочное владение (ЛДВ)
 *
 * При владении ценной бумагой 3 года и дольше прибыль от продажи
 * освобождается от НДФЛ в пределах лимита (3 млн ₽ за каждый полный год
 * владения, максимум 9 млн за 3 года — актуально на 2026 год, если лимит
 * или срок изменятся, поправь LDV_YEARS_REQUIRED ниже).
 *
 * Показывает по каждой акции: сколько штук уже прошли порог в 3 года
 * (и попадают под льготу при продаже), сколько ещё не прошли, и когда
 * ближайший транш станет льготным.
 *
 * Зависимости: readConfig_(), readPositions_(), rub_(), DST, C — dashboard.js
 *              getAccounts_(), fetchOperations_() — history.js
 *              matchesName_(), buildFifoLots_() — avgprice.js
 *              readAdvancedParams_() — advparams.js
 *              renderSection_() — sections.js
 */

const LDV_PROP           = 'LDV_DATA';
const LDV_SECTION_TITLE  = '▌ ЛЬГОТА НА ДОЛГОСРОЧНОЕ ВЛАДЕНИЕ (ЛДВ)';
const LDV_YEARS_REQUIRED = 3;

function calculateLdvEligibility() {
  let config;
  try { config = readConfig_(); }
  catch (e) { SpreadsheetApp.getUi().alert('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  let shares    = positions.filter(function(p) { return p.category === 'Акции'; });
  if (!shares.length) { SpreadsheetApp.getUi().alert('Нет акций в портфеле.'); return; }

  let params   = readAdvancedParams_();
  if (params.fifoYears < LDV_YEARS_REQUIRED) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Глубина истории FIFO (Config → Продвинутые параметры) — ' + params.fifoYears +
      ' лет, а для ЛДВ нужно минимум ' + LDV_YEARS_REQUIRED + '. Увеличь значение, иначе расчёт будет неполным.'
    );
  }

  let fromDate = new Date(new Date().getTime() - Math.max(params.fifoYears, LDV_YEARS_REQUIRED + 1) * 365 * 24 * 3600 * 1000);
  let toDate   = new Date();

  let allOps = [];
  getAccounts_().forEach(function(acc) {
    fetchOperations_(acc.id, acc.name, fromDate, toDate).forEach(function(op) {
      let isTrade = op.type === 'OPERATION_TYPE_BUY'  || op.type === 'OPERATION_TYPE_BUY_CARD' ||
                    op.type === 'OPERATION_TYPE_SELL' || op.type === 'OPERATION_TYPE_SELL_CARD';
      if (isTrade && op.quantity && op.price) allOps.push(op);
    });
  });

  let now = new Date();
  let results = shares.map(function(p) {
    let myOps = allOps.filter(function(op) { return matchesName_(op.instrumentName, p.name); })
                       .sort(function(a, b) { return a.date - b.date; });
    let lots = buildFifoLots_(myOps);

    let eligibleQty = 0, notYetQty = 0, nextEligibleDate = null;
    lots.forEach(function(lot) {
      let lotDate = lot.date instanceof Date ? lot.date : new Date(lot.date);
      let ageYears = (now - lotDate) / (365 * 24 * 3600 * 1000);
      if (ageYears >= LDV_YEARS_REQUIRED) {
        eligibleQty += lot.qty;
      } else {
        notYetQty += lot.qty;
        let eligibleOn = new Date(lotDate.getTime() + LDV_YEARS_REQUIRED * 365 * 24 * 3600 * 1000);
        if (!nextEligibleDate || eligibleOn < nextEligibleDate) nextEligibleDate = eligibleOn;
      }
    });

    return {
      name: p.name,
      eligibleQty: eligibleQty,
      notYetQty: notYetQty,
      nextEligibleDate: nextEligibleDate ? nextEligibleDate.getTime() : null,
    };
  }).filter(function(r) { return r.eligibleQty > 0 || r.notYetQty > 0; });

  PropertiesService.getScriptProperties().setProperty(LDV_PROP, JSON.stringify(results));
  redrawLdvSection_();
}

function redrawLdvSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(LDV_PROP);
  if (!raw) return;
  let results = JSON.parse(raw);
  let tz = Session.getScriptTimeZone();

  renderSection_(sh, LDV_SECTION_TITLE, function(sh, r, COLS) {
    sh.getRange(r, 1).setValue('Акция');
    sh.getRange(r, 2).setValue('Есть льгота (шт.)');
    sh.getRange(r, 3, 1, 2).merge().setValue('Пока нет льготы (шт.) / когда появится');
    sh.getRange(r, 1, 1, COLS).setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold');
    r++;

    results.forEach(function(res, idx) {
      let bg = idx % 2 === 0 ? C.EVEN : C.ODD;
      sh.getRange(r, 1).setValue(res.name).setBackground(bg);
      sh.getRange(r, 2).setValue(res.eligibleQty)
        .setFontColor(res.eligibleQty > 0 ? C.OK : '#9e9e9e').setFontWeight('bold').setBackground(bg);

      if (res.notYetQty > 0) {
        let dateStr = res.nextEligibleDate
          ? Utilities.formatDate(new Date(res.nextEligibleDate), tz, 'dd.MM.yyyy')
          : '—';
        sh.getRange(r, 3, 1, 2).merge()
          .setValue(res.notYetQty + ' шт. — ближайший транш станет льготным ' + dateStr)
          .setFontColor(C.WARN).setBackground(bg);
      } else {
        sh.getRange(r, 3, 1, 2).merge()
          .setValue('✅ Вся позиция уже под льготой')
          .setFontColor(C.OK).setBackground(bg);
      }
      r++;
    });

    r++;
    sh.getRange(r, 1, 1, COLS).merge()
      .setValue('ℹ️ Продажа лотов «пока нет льготы» облагается НДФЛ на прибыль как обычно. Проверяй перед продажей, какие именно лоты списываются по FIFO.')
      .setFontColor('#9e9e9e').setFontStyle('italic').setFontSize(9);
  });
}
