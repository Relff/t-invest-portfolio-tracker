/**
 * correlation.js — Корреляция между акциями портфеля
 *
 * Считает попарную корреляцию Пирсона по дневным доходностям за последние
 * CORR_DAYS_BACK дней и рисует цветную матрицу на отдельном листе.
 * Помогает увидеть, что портфель диверсифицирован по названиям, но не
 * обязательно по поведению — например, два разных банка часто двигаются
 * почти синхронно.
 *
 * ⚠️ Единственная функция в проекте, которая делает по одному GetCandles-
 * запросу на КАЖДУЮ акцию в портфеле — при 9-10 акциях это ощутимо дольше,
 * чем остальная аналитика. Специально не привязана ни к одному триггеру,
 * запускается только вручную из меню.
 *
 * Зависимости: readConfig_(), readPositions_() — dashboard.js
 *              buildFigiMap_() — income.js
 *              tiFetch_(), qToNumber_() — tinvest.js
 *              rub_(), mergedCell_(), C — dashboard.js
 */

const CORR_PROP          = 'CORRELATION_DATA';
const CORR_SHEET         = 'Корреляция акций';
const CORR_DAYS_BACK     = 180;

function calculateStockCorrelation() {
  let ui = SpreadsheetApp.getUi();
  let config;
  try { config = readConfig_(); }
  catch (e) { ui.alert('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  let shares    = positions.filter(function(p) { return p.category === 'Акции'; });
  if (shares.length < 2) {
    ui.alert('Нужно минимум 2 акции в портфеле, чтобы считать корреляцию.');
    return;
  }

  let figiMap  = buildFigiMap_();
  let toDate   = new Date();
  let fromDate = new Date(toDate.getTime() - CORR_DAYS_BACK * 24 * 3600 * 1000);

  // ── Тянем дневные цены по каждой акции ──────────────────────────
  let priceSeries = {}; // name -> { 'yyyy-mm-dd': closePrice }
  shares.forEach(function(p) {
    let figi = p.figi || figiMap[p.name] || figiMap[p.ticker] || '';
    if (!figi) return;
    try {
      let resp = tiFetch_('/tinkoff.public.invest.api.contract.v1.MarketDataService/GetCandles', {
        instrumentId: figi, from: fromDate.toISOString(), to: toDate.toISOString(),
        interval: 'CANDLE_INTERVAL_DAY',
      });
      let map = {};
      (resp.candles || []).forEach(function(c) {
        map[c.time.substring(0, 10)] = qToNumber_(c.close);
      });
      if (Object.keys(map).length > 5) priceSeries[p.name] = map;
      Utilities.sleep(60); // throttle — как в benchmark.js/calendar.js
    } catch (e) {
      console.warn('Корреляция: ошибка свечей для ' + p.name + ': ' + e.message);
    }
  });

  let names = Object.keys(priceSeries);
  if (names.length < 2) {
    ui.alert('Не удалось получить достаточно ценовой истории хотя бы по двум акциям.');
    return;
  }

  // ── Общие торговые дни для всех бумаг сразу (пересечение дат) ──
  let commonDates = null;
  names.forEach(function(n) {
    let dates = Object.keys(priceSeries[n]);
    commonDates = commonDates ? commonDates.filter(function(d) { return dates.indexOf(d) >= 0; }) : dates;
  });
  commonDates.sort();

  if (commonDates.length < 10) {
    ui.alert('Слишком мало общих торговых дней (' + commonDates.length + ') для надёжного расчёта.');
    return;
  }

  // ── Дневные доходности по общим датам ───────────────────────────
  let returns = {};
  names.forEach(function(n) {
    let series = commonDates.map(function(d) { return priceSeries[n][d]; });
    let ret = [];
    for (let i = 1; i < series.length; i++) {
      ret.push(series[i - 1] > 0 ? series[i] / series[i - 1] - 1 : 0);
    }
    returns[n] = ret;
  });

  // ── Попарная корреляция Пирсона ─────────────────────────────────
  let matrix = {};
  let pairs  = []; // для поиска мин/макс пары
  names.forEach(function(a) {
    matrix[a] = {};
    names.forEach(function(b) {
      let r = a === b ? 1 : pearsonCorr_(returns[a], returns[b]);
      matrix[a][b] = Math.round(r * 100) / 100;
      if (a < b) pairs.push({ a: a, b: b, r: r });
    });
  });

  pairs.sort(function(x, y) { return y.r - x.r; });

  PropertiesService.getScriptProperties().setProperty(CORR_PROP, JSON.stringify({
    names: names, matrix: matrix, days: commonDates.length,
    topPair: pairs[0] || null, bottomPair: pairs[pairs.length - 1] || null,
  }));

  writeCorrelationSheet_();
  ui.alert('✅ Готово. Лист «' + CORR_SHEET + '» обновлён (' + names.length + ' акций, ' + commonDates.length + ' общих торговых дней).');
}

function pearsonCorr_(x, y) {
  let n = x.length;
  if (n === 0) return 0;
  let meanX = x.reduce(function(s, v) { return s + v; }, 0) / n;
  let meanY = y.reduce(function(s, v) { return s + v; }, 0) / n;
  let num = 0, denX = 0, denY = 0;
  for (let i = 0; i < n; i++) {
    let dx = x[i] - meanX, dy = y[i] - meanY;
    num += dx * dy; denX += dx * dx; denY += dy * dy;
  }
  let den = Math.sqrt(denX * denY);
  return den > 0 ? num / den : 0;
}

function corrColor_(v) {
  if (v >= 0.7)  return { bg: '#b71c1c', fg: '#ffffff' }; // сильно синхронны — низкая диверсификация
  if (v >= 0.4)  return { bg: '#e65100', fg: '#ffffff' };
  if (v >= 0.15) return { bg: '#f9a825', fg: '#000000' };
  if (v >= -0.15) return { bg: '#43a047', fg: '#ffffff' }; // почти независимы — хорошо
  return { bg: '#1565c0', fg: '#ffffff' }; // отрицательная — отличный диверсификатор
}

function writeCorrelationSheet_() {
  let raw = PropertiesService.getScriptProperties().getProperty(CORR_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);

  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CORR_SHEET);
  if (!sh) sh = ss.insertSheet(CORR_SHEET);
  sh.clearContents();
  sh.clearFormats();

  let n = d.names.length;
  let tz = Session.getScriptTimeZone();
  let now = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');

  mergedCell_(sh, 1, 1, 1, n + 1,
    '🔗  КОРРЕЛЯЦИЯ МЕЖДУ АКЦИЯМИ — дневные доходности за ' + d.days + ' торговых дней',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 13, align: 'center' });
  mergedCell_(sh, 2, 1, 1, n + 1,
    'Обновлено: ' + now + '   ·   🔴 сильно синхронны   🟠/🟡 умеренно   🟢 почти независимы   🔵 отрицательная (лучшая диверсификация)',
    { bg: '#263238', fg: '#b0bec5', align: 'center' });

  // ── Заголовки матрицы ────────────────────────────────────────────
  let headerRow = 4;
  sh.getRange(headerRow, 1).setValue('').setBackground(C.DARK);
  d.names.forEach(function(name, i) {
    let short = name.length > 14 ? name.substring(0, 14) + '…' : name;
    sh.getRange(headerRow, i + 2).setValue(short)
      .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold')
      .setFontSize(9).setHorizontalAlignment('center');
    sh.getRange(headerRow + 1 + i, 1).setValue(short)
      .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
  });

  // ── Тело матрицы ──────────────────────────────────────────────────
  d.names.forEach(function(a, ri) {
    d.names.forEach(function(b, ci) {
      let v = d.matrix[a][b];
      let cell = sh.getRange(headerRow + 1 + ri, ci + 2);
      if (a === b) {
        cell.setValue('—').setBackground('#37474f').setFontColor('#90a4ae').setHorizontalAlignment('center');
      } else {
        let color = corrColor_(v);
        cell.setValue(v).setNumberFormat('0.00')
          .setBackground(color.bg).setFontColor(color.fg)
          .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
      }
    });
  });

  let footerRow = headerRow + 1 + n + 1;
  if (d.topPair) {
    sh.getRange(footerRow, 1, 1, n + 1).merge()
      .setValue('🔴 Сильнее всего синхронны: ' + d.topPair.a + ' ↔ ' + d.topPair.b +
                '  (r = ' + d.topPair.r.toFixed(2) + ') — вместе почти не диверсифицируют')
      .setFontColor(C.CRIT).setFontStyle('italic');
    footerRow++;
  }
  if (d.bottomPair) {
    sh.getRange(footerRow, 1, 1, n + 1).merge()
      .setValue('🔵 Лучшая пара для диверсификации: ' + d.bottomPair.a + ' ↔ ' + d.bottomPair.b +
                '  (r = ' + d.bottomPair.r.toFixed(2) + ')')
      .setFontColor('#1565c0').setFontStyle('italic');
    footerRow++;
  }

  sh.setColumnWidth(1, 140);
  for (let c = 2; c <= n + 1; c++) sh.setColumnWidth(c, 80);
  sh.setFrozenRows(headerRow);
  sh.setFrozenColumns(1);
}
