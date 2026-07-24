/**
 * calendar.js — Календарь выплат
 *
 * Показывает ожидаемые купоны и дивиденды по месяцам.
 *
 * Листы:
 *   «Календарь выплат» — месячная сводка + список ближайших выплат
 *
 * Установка:
 *   1. Добавить файл calendar в Apps Script
 *   2. В onOpen() добавить:
 *      .addItem('📅  Обновить Календарь выплат', 'updateCalendarSheet')
 *
 * Зависимости (из других файлов проекта):
 *   tiFetch_(), moneyToNumber_()         — tinvest.gs
 *   C, DST, rub_(), mergedCell_(), hdrRow_() — dashboard.gs
 */

// Название листа берётся из DST.CALENDAR (dashboard.gs)

// ════════════════════════════════════════════════════════════════════
// ГЛАВНАЯ ФУНКЦИЯ
// ════════════════════════════════════════════════════════════════════

function updateCalendarSheet() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.CALENDAR);
  if (!sh) sh = ss.insertSheet(DST.CALENDAR);
  sh.clearContents();
  sh.clearFormats();

  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var nowStr = Utilities.formatDate(now, tz, 'dd.MM.yyyy HH:mm');

  // Читаем позиции
  var bondPositions  = readSheetPositions_('_Дан_Облигации');
  var sharePositions = readSheetPositions_('_Дан_Акции');
  var figiMap        = buildFigiMap_();
  var divMap         = readDividendsFromConfig_();

  // Собираем расписание выплат
  var couponPayments   = fetchCouponCalendar_(bondPositions, figiMap);
  var dividendPayments = fetchDividendCalendar_(sharePositions, divMap, figiMap);
  var allPayments      = couponPayments.concat(dividendPayments);

  // Строим месячные итоги
  var monthlyGrid = buildMonthlyGrid_(allPayments, now);

  // Рендерим лист
  var r    = 1;
  var COLS = 7;

  // ── Шапка ─────────────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS,
    '📅  КАЛЕНДАРЬ ВЫПЛАТ — купоны и дивиденды',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });
  r++;

  mergedCell_(sh, r, 1, 1, COLS,
    'Обновлено: ' + nowStr + '   ·   Данные за следующие 12 месяцев',
    { bg: '#263238', fg: '#b0bec5', align: 'center' });
  r += 2;

  // ── Месячная сводка ───────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '▌ МЕСЯЧНАЯ СВОДКА',
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;

  hdrRow_(sh, r,
    ['Месяц', 'Купоны, ₽', 'Дивиденды, ₽', 'Итого, ₽', 'Прогресс', '', ''],
    COLS);
  r++;
let monthlyDataStartRow = r; // запоминаем начало данных для графика

  // Находим максимум для прогресс-баров
  var maxMonthly = 0;
  Object.keys(monthlyGrid).forEach(function(m) {
    var tot = (monthlyGrid[m].coupons || 0) + (monthlyGrid[m].dividends || 0);
    if (tot > maxMonthly) maxMonthly = tot;
  });

  var MONTH_NAMES = [
    'Январь','Февраль','Март','Апрель','Май','Июнь',
    'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'
  ];

  var yearCoupons   = 0;
  var yearDividends = 0;

  for (var mi = 0; mi < 12; mi++) {
    var monthIdx  = ((now.getMonth() + mi) % 12) + 1; // 1-12
    var yearOffset = (now.getMonth() + mi) >= 12 ? 1 : 0;
    var dispYear  = now.getFullYear() + yearOffset;
    var grid      = monthlyGrid[monthIdx + '_' + dispYear] || { coupons: 0, dividends: 0 };
    var couponAmt = grid.coupons   || 0;
    var divAmt    = grid.dividends || 0;
    var total     = couponAmt + divAmt;

    yearCoupons   += couponAmt;
    yearDividends += divAmt;

    // Прогресс-бар
    var barLen  = maxMonthly > 0 ? Math.round((total / maxMonthly) * 20) : 0;
    var barStr  = '█'.repeat(barLen) + '░'.repeat(20 - barLen);
    var barColor = total > 5000 ? '#1b5e20' : total > 2000 ? '#f57f17' : '#9e9e9e';

    // Подсветка текущего месяца
    var isCurrentMonth = (monthIdx === now.getMonth() + 1 && dispYear === now.getFullYear());
    var bg = isCurrentMonth ? '#e8f5e9' : (mi % 2 === 0 ? C.EVEN : C.ODD);

    sh.getRange(r, 1).setValue(MONTH_NAMES[monthIdx - 1] + ' ' + dispYear)
      .setBackground(bg).setFontWeight(isCurrentMonth ? 'bold' : 'normal');
    sh.getRange(r, 2).setValue(Math.round(couponAmt)).setNumberFormat('#,##0 [$₽-ru-RU]').setBackground(bg);
    sh.getRange(r, 3).setValue(Math.round(divAmt)).setNumberFormat('#,##0 [$₽-ru-RU]').setBackground(bg);
    sh.getRange(r, 4).setValue(Math.round(total)).setNumberFormat('#,##0 [$₽-ru-RU]').setBackground(bg)
      .setFontWeight('bold');
    sh.getRange(r, 5, 1, 3).merge()
      .setValue(barStr + '  ' + rub_(total))
      .setBackground(bg).setFontColor(barColor)
      .setFontFamily('Courier New').setFontSize(9);
    r++;
  }

buildIncomeChart_(sh, monthlyDataStartRow); 

  // Итоговая строка
  var yearTotal = yearCoupons + yearDividends;
  sh.getRange(r, 1, 1, COLS).merge()
    .setValue('ИТОГО ЗА 12 МЕСЯЦЕВ:   купоны ' + rub_(yearCoupons) +
              '  +  дивиденды ' + rub_(yearDividends) +
              '  =  ' + rub_(yearTotal))
    .setBackground(C.DARK).setFontColor('#ffd54f')
    .setFontWeight('bold').setHorizontalAlignment('center');
  r += 2;

  // ── Все будущие выплаты (12 месяцев, без дублирования) ──────────
  // 🟡 Строки с выплатами в ближайшие 7 дней подсвечены жёлтым
  var allFuture = allPayments
    .filter(function(p) { return p.date && p.date >= now; })
    .sort(function(a, b) { return a.date - b.date; });

  mergedCell_(sh, r, 1, 1, COLS,
    '▌ ВСЕ ВЫПЛАТЫ — следующие 12 месяцев (' + allFuture.length + ' выплат)' +
    '   |   🟢 купон (точно)   🟡 дивиденд (оценка)   |   жирная дата = ближайшие 7 дней',
    { bg: '#37474f', fg: '#ffffff', bold: true });
  r++;

  let paymentsHeaderRow = r;
  hdrRow_(sh, r,
    ['Дата', 'Инструмент', 'Тип', '₽ / шт', 'Кол-во', 'Итого, ₽', 'Месяц'],
    COLS);
  r++;

  allFuture.forEach(function(p, idx) {
    let monthStr  = MONTH_NAMES[p.date.getMonth()] + ' ' + p.date.getFullYear();
    let isCoupon  = p.type.indexOf('Купон') === 0;
    let bg        = isCoupon
      ? (idx % 2 === 0 ? '#e8f5e9' : '#f1f8f2')
      : (idx % 2 === 0 ? '#fff8e1' : '#fffdf0');
    let daysLeft  = Math.ceil((p.date - now) / (24 * 3600 * 1000));
    let urgent    = daysLeft <= 7;

    sh.getRange(r, 1).setValue(p.date).setNumberFormat('dd.mm.yyyy').setBackground(bg);
    sh.getRange(r, 2, 1, COLS - 1).setValues([[
      p.name, p.type,
      Math.round(p.perUnit * 100) / 100,
      p.qty,
      Math.round(p.total),
      monthStr
    ]]).setBackground(bg);

    sh.getRange(r, 4).setNumberFormat('#,##0.00 [$₽-ru-RU]');
    sh.getRange(r, 6).setNumberFormat('#,##0 [$₽-ru-RU]');
    if (urgent) sh.getRange(r, 1).setFontColor('#e65100').setFontWeight('bold');
    r++;
  });

  // ── Ширина колонок ────────────────────────────────────────────────
  [100, 230, 90, 130, 80, 140, 130].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(6);

  // Добавляем фильтр на таблицу выплат (чтобы можно было фильтровать по месяцу/инструменту)
  try {
    var existingFilter = sh.getFilter();
    if (existingFilter) existingFilter.remove();
    // Фильтр на всю таблицу (от строки заголовка до конца данных)
    sh.getRange(paymentsHeaderRow, 1, allFuture.length + 1, COLS).createFilter();
  } catch(e) {
    console.warn('Календарь: не удалось создать фильтр: ' + e.message);
  }

  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════════════
// КУПОННЫЙ КАЛЕНДАРЬ — точные даты из T-Invest API
// ════════════════════════════════════════════════════════════════════

/**
 * Для каждой облигации в портфеле получаем расписание купонов
 * через GetBondCoupons. Возвращаем массив платежей с точными датами.
 */
function fetchCouponCalendar_(bondPositions, figiMap) {
  var now  = new Date();
  var in1y = new Date(now.getTime() + 366 * 24 * 3600 * 1000);
  var payments = [];

  bondPositions.forEach(function(p) {
    var figi = p.figi || figiMap[p.name] || figiMap[p.ticker] || '';
    if (!figi) return;

    try {
      var resp = tiFetch_(
        '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetBondCoupons',
        { figi: figi, from: now.toISOString(), to: in1y.toISOString() }
      );

      var events = resp.events || [];
      events.forEach(function(c) {
        // Дата купона
        var dateRaw = c.couponDate;
        if (!dateRaw) return;

        var payDate;
        if (typeof dateRaw === 'string') {
          payDate = new Date(dateRaw);
        } else if (dateRaw.seconds) {
          payDate = new Date(Number(dateRaw.seconds) * 1000);
        } else {
          return;
        }

        if (isNaN(payDate.getTime())) return;
        if (payDate < now) return; // пропускаем прошедшие

        var perUnit = moneyToNumber_(c.payOneBond || c.couponAmount || null);
        if (perUnit <= 0) return;

        payments.push({
          name:    p.name,
          ticker:  p.ticker,
          type:    'Купон',
          date:    payDate,
          perUnit: perUnit,
          qty:     p.qty,
          total:   perUnit * p.qty,
        });
      });

      Utilities.sleep(50); // throttle API calls
    } catch(e) {
      // Пропускаем облигацию если ошибка API
      console.warn('Календарь: ошибка купонов для ' + p.name + ': ' + e.message);
    }
  });

  return payments;
}


// ════════════════════════════════════════════════════════════════════
// ДИВИДЕНДНЫЙ КАЛЕНДАРЬ — исторические даты + Config суммы
// ════════════════════════════════════════════════════════════════════

/**
 * Для акций пытаемся определить типичный месяц выплаты дивидендов
 * через GetDividends (исторические данные).
 * Сумму берём из Config (Блок 4).
 */
function fetchDividendCalendar_(sharePositions, divMap, figiMap) {
  let payments = [];
  let now      = new Date();

  sharePositions.forEach(function(p) {
    let divData = findDividend_(p.name, p.ticker, divMap);
    let figi = p.figi || figiMap[p.name] || figiMap[p.ticker] || '';

    // Если в Config пусто — считаем автоматически из истории (та же логика, что в income.js)
    if ((!divData || divData.amount <= 0) && figi) {
      let hist = fetchAnnualDividendFromHistory_(figi);
      if (hist.perUnit > 0) divData = { amount: hist.perUnit, note: hist.note };
    }

    if (!divData || divData.amount <= 0) return; // реально нет дивидендов

    let payMonth = null;
    let payYear  = now.getFullYear();

    if (figi) {
      try {
        let from2y = new Date(now.getTime() - 2 * 365 * 24 * 3600 * 1000);
        let resp   = tiFetch_(
          '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetDividends',
          { figi: figi, from: from2y.toISOString(), to: now.toISOString() }
        );
        let divs = resp.dividends || [];
        if (divs.length > 0) {
          let lastDiv = divs[divs.length - 1];
          let dateRaw = lastDiv.paymentDate || lastDiv.recordDate;
          if (dateRaw) {
            let histDate;
            if (typeof dateRaw === 'string') {
              histDate = new Date(dateRaw);
            } else if (dateRaw.seconds) {
              histDate = new Date(Number(dateRaw.seconds) * 1000);
            }
            if (histDate && !isNaN(histDate.getTime())) {
              payMonth = histDate.getMonth();
              let estimDate = new Date(payYear, payMonth, 15);
              if (estimDate < now) payYear = payYear + 1;
            }
          }
        }
        Utilities.sleep(50);
      } catch(e) {
        console.warn('Календарь: дивидендная история для ' + p.name + ': ' + e.message);
      }
    }

    if (payMonth === null) return;

    let estimatedDate = new Date(payYear, payMonth, 15);
    payments.push({
      name:    p.name,
      ticker:  p.ticker,
      type:    'Дивиденд ~',
      date:    estimatedDate,
      perUnit: divData.amount,
      qty:     p.qty,
      total:   divData.amount * p.qty,
    });
  });

  return payments;
}


// ════════════════════════════════════════════════════════════════════
// ПОСТРОЕНИЕ МЕСЯЧНОЙ СЕТКИ
// ════════════════════════════════════════════════════════════════════

/**
 * Группирует платежи по ключу "месяц_год".
 * Возвращает объект: { "1_2026": {coupons: X, dividends: Y}, ... }
 */
function buildMonthlyGrid_(payments, fromDate) {
  var grid = {};

  payments.forEach(function(p) {
    if (!p.date || isNaN(p.date.getTime())) return;
    if (p.date < fromDate) return;

    var month  = p.date.getMonth() + 1; // 1-12
    var year   = p.date.getFullYear();
    var key    = month + '_' + year;

    if (!grid[key]) {
      grid[key] = { coupons: 0, dividends: 0 };
    }

    if (p.type === 'Купон') {
      grid[key].coupons += p.total;
    } else {
      grid[key].dividends += p.total;
    }
  });

  return grid;
}

// ════════════════════════════════════════════════════════════════════
// ГРАФИК ДОХОДА ПО МЕСЯЦАМ
// ════════════════════════════════════════════════════════════════════

function buildIncomeChart_(sh, dataStartRow) {
  sh.getCharts().forEach(function(c) { sh.removeChart(c); });

  let range = sh.getRange(dataStartRow - 1, 1, 13, 3);

  let chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setOption('isStacked', true)
    .setOption('title', 'Пассивный доход по месяцам')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1a237e' })
    .setOption('legend', { position: 'top', textStyle: { fontSize: 12 } })
    .setOption('series', {
      0: { labelInLegend: 'Купоны, ₽' },
      1: { labelInLegend: 'Дивиденды, ₽' }
    })
    .setOption('colors', ['#43a047', '#fbc02d'])
    .setOption('backgroundColor', '#ffffff')
    .setOption('vAxis', { title: '₽' })
    .setOption('hAxis', { textStyle: { fontSize: 9 } })
    .setOption('width', 620)
    .setOption('height', 340)
    .setPosition(6, 9, 0, 0)
    .build();

  sh.insertChart(chart);
}
