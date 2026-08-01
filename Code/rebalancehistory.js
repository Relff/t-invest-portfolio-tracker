/**
 * rebalancehistory.js — История ребалансировок: рекомендовано vs куплено по факту
 *
 * Каждый прогон calculateRebalance() (dashboard.js) тихо логирует рекомендации
 * по акциям в скрытый технический лист (без побочных эффектов, просто дописывает
 * строку). Пункт меню «История ребалансировок: сравнить с фактом» берёт
 * последнюю такую пачку и сверяет её с реальными покупками, которые прошли
 * с момента расчёта до сейчас — так постепенно копится история того,
 * насколько ты в реальности следуешь собственному плану ребалансировки.
 *
 * Сравнение — только по акциям (для золота/денежного рынка/замещаек
 * сопоставление по названию категории с конкретным инструментом ненадёжно).
 *
 * Зависимости: getAccounts_(), fetchOperations_() — history.js
 *              matchesName_() — avgprice.js
 *              rub_(), mergedCell_(), hdrRow_(), C — dashboard.js
 */

const REBAL_LOG_SHEET     = '_Лог_Ребаланс';
const REBAL_HISTORY_SHEET = 'История ребалансировок';

/**
 * Тихо логирует одну пачку рекомендаций по акциям — вызывается из
 * calculateRebalance() при каждом пересчёте с ненулевой суммой пополнения.
 * items: [{ name, amount }]
 */
function logRebalanceRecommendation_(batchDate, items) {
  if (!items.length) return;
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(REBAL_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(REBAL_LOG_SHEET);
    sh.appendRow(['Дата пачки', 'Название', 'Рекомендовано, ₽']);
    sh.hideSheet();
  }
  let rows = items.map(function(it) { return [batchDate, it.name, Math.round(it.amount)]; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}

function compareRebalanceExecution() {
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActive();
  let logSh = ss.getSheetByName(REBAL_LOG_SHEET);
  if (!logSh || logSh.getLastRow() < 2) {
    ui.alert('Нет ни одной сохранённой рекомендации. Сначала хотя бы раз запусти «Пересчитать калькулятор пополнения» с заполненной суммой.');
    return;
  }

  let data = logSh.getRange(2, 1, logSh.getLastRow() - 1, 3).getValues();

  // Последняя пачка — максимальная дата среди залогированных
  let lastBatchTime = data.reduce(function(max, row) {
    let t = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();
    return Math.max(max, t);
  }, 0);
  let lastBatchDate = new Date(lastBatchTime);

  let batchItems = data
    .filter(function(row) {
      let t = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();
      return t === lastBatchTime;
    })
    .map(function(row) { return { name: row[1], recommended: Number(row[2]) || 0 }; });

  if (!batchItems.length) {
    ui.alert('Не удалось найти последнюю пачку рекомендаций.');
    return;
  }

  // Реальные покупки с момента последней рекомендации до сейчас
  let now = new Date();
  let allBuyOps = [];
  getAccounts_().forEach(function(acc) {
    fetchOperations_(acc.id, acc.name, lastBatchDate, now).forEach(function(op) {
      let isBuy = op.type === 'OPERATION_TYPE_BUY' || op.type === 'OPERATION_TYPE_BUY_CARD';
      if (isBuy) allBuyOps.push(op);
    });
  });

  let results = batchItems.map(function(item) {
    let boughtOps  = allBuyOps.filter(function(op) { return matchesName_(op.instrumentName, item.name); });
    let boughtSum  = boughtOps.reduce(function(s, op) { return s + (op.amount || 0); }, 0);
    let pct        = item.recommended > 0 ? boughtSum / item.recommended : 0;
    return { name: item.name, recommended: item.recommended, bought: boughtSum, pct: pct };
  });

  writeRebalanceHistorySection_(lastBatchDate, results);
}

function writeRebalanceHistorySection_(batchDate, results) {
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(REBAL_HISTORY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(REBAL_HISTORY_SHEET);
    mergedCell_(sh, 1, 1, 1, 6,
      '📋  ИСТОРИЯ РЕБАЛАНСИРОВОК — рекомендовано vs куплено по факту',
      { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });
  }

  let tz = Session.getScriptTimeZone();
  let r  = sh.getLastRow() + 2;

  mergedCell_(sh, r, 1, 1, 6,
    '▌ Пачка от ' + Utilities.formatDate(batchDate, tz, 'dd.MM.yyyy HH:mm'),
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;
  hdrRow_(sh, r, ['Название', 'Рекомендовано, ₽', 'Куплено факт, ₽', '% выполнения', 'Статус', ''], 6);
  r++;

  let totalRec = 0, totalBought = 0;
  results.forEach(function(res, idx) {
    let bg     = idx % 2 === 0 ? C.EVEN : C.ODD;
    let status = res.pct >= 0.9 ? '✅ Выполнено' : res.pct >= 0.5 ? '⚠️ Частично' : '❌ Не куплено';
    let clr    = res.pct >= 0.9 ? C.OK : res.pct >= 0.5 ? C.WARN : C.CRIT;

    sh.getRange(r, 1).setValue(res.name).setBackground(bg);
    sh.getRange(r, 2).setValue(res.recommended).setNumberFormat('#,##0 [$₽-ru-RU]').setBackground(bg);
    sh.getRange(r, 3).setValue(res.bought).setNumberFormat('#,##0 [$₽-ru-RU]').setBackground(bg);
    sh.getRange(r, 4).setValue(res.pct).setNumberFormat('0%').setBackground(bg)
      .setFontColor(clr).setFontWeight('bold');
    sh.getRange(r, 5, 1, 2).merge().setValue(status).setBackground(bg)
      .setFontColor(clr).setFontWeight('bold');

    totalRec    += res.recommended;
    totalBought += res.bought;
    r++;
  });

  let totalPct = totalRec > 0 ? totalBought / totalRec : 0;
  sh.getRange(r, 1, 1, 6).merge()
    .setValue('ИТОГО ПО ПАЧКЕ:  ' + rub_(totalBought) + ' из ' + rub_(totalRec) +
              '  (' + (totalPct * 100).toFixed(0) + '%)')
    .setBackground(C.DARK).setFontColor('#ffd54f').setFontWeight('bold').setHorizontalAlignment('center');

  [250, 150, 150, 120, 150, 50].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  ui.alert('✅ Сравнение добавлено в лист «' + REBAL_HISTORY_SHEET + '» — новой секцией снизу, старые не перезаписываются.');
}
