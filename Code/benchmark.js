/**
 * benchmark.js — Сравнение портфеля с IMOEX + единая точка входа «Аналитика доходности»
 */

const IMOEX_UID_PROPERTY = 'IMOEX_UID';
const BENCH_YEARS_BACK   = 5;

const ANALYTICS_SECTION_TITLE = '▌ АНАЛИТИКА ДОХОДНОСТИ';
const ANALYTICS_XIRR_PROP     = 'ANALYTICS_XIRR_VALUE';
const ANALYTICS_BENCH_PROP    = 'ANALYTICS_BENCH_VALUE';

function getImoexUid_() {
  let cached = PropertiesService.getScriptProperties().getProperty(IMOEX_UID_PROPERTY);
  if (cached) return cached;
  let resp = tiFetch_('/tinkoff.public.invest.api.contract.v1.InstrumentsService/Indicatives', {});
  let list = resp.instruments || [];
  let imoex = list.find(function(i) { return i.ticker === 'IMOEX'; });
  if (!imoex) throw new Error('Не удалось найти IMOEX в списке индикативов T-Invest API');
  PropertiesService.getScriptProperties().setProperty(IMOEX_UID_PROPERTY, imoex.uid);
  return imoex.uid;
}

function getImoexPriceMap_(uid, fromDate, toDate) {
  let priceMap = {};
  let chunkStart = new Date(fromDate);
  while (chunkStart < toDate) {
    let chunkEnd = new Date(Math.min(chunkStart.getTime() + 365 * 24 * 3600 * 1000, toDate.getTime()));
    let resp = tiFetch_('/tinkoff.public.invest.api.contract.v1.MarketDataService/GetCandles', {
      instrumentId: uid, from: chunkStart.toISOString(), to: chunkEnd.toISOString(), interval: 'CANDLE_INTERVAL_DAY',
    });
    (resp.candles || []).forEach(function(c) {
      priceMap[c.time.substring(0, 10)] = qToNumber_(c.close);
    });
    chunkStart = new Date(chunkEnd.getTime() + 24 * 3600 * 1000);
    pause_(60);
  }
  return priceMap;
}

function findNearestPrice_(priceMap, date) {
  for (let i = 0; i <= 5; i++) {
    let d = new Date(date.getTime() - i * 24 * 3600 * 1000);
    let key = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
    if (priceMap[key]) return priceMap[key];
  }
  return null;
}

function calculateBenchmark() {
  let uid;
  try { uid = getImoexUid_(); }
  catch (e) { SpreadsheetApp.getUi().alert('Ошибка получения IMOEX: ' + e.message); return; }

  let fromDate = new Date(new Date().getTime() - BENCH_YEARS_BACK * 365 * 24 * 3600 * 1000);
  let toDate   = new Date();

  let deposits = [];
  getAccounts_().forEach(function(acc) {
    fetchOperations_(acc.id, acc.name, fromDate, toDate).forEach(function(op) {
      if (op.type === 'OPERATION_TYPE_INPUT') deposits.push({ date: op.date, amount: op.amount });
    });
  });
  if (!deposits.length) {
    SpreadsheetApp.getUi().alert('Нет пополнений за последние ' + BENCH_YEARS_BACK + ' лет — сравнивать не с чем.');
    return;
  }

  let priceMap = getImoexPriceMap_(uid, fromDate, toDate);
  let currentPrice = findNearestPrice_(priceMap, toDate);
  if (!currentPrice) {
    SpreadsheetApp.getUi().alert('Не удалось получить текущую цену IMOEX через T-Invest API.');
    return;
  }

  let totalUnits = 0, skipped = 0;
  deposits.forEach(function(d) {
    let priceThen = findNearestPrice_(priceMap, d.date);
    if (priceThen && priceThen > 0) totalUnits += d.amount / priceThen; else skipped++;
  });

  let hypotheticalValue = Math.round(totalUnits * currentPrice);
  let config    = readConfig_();
  let positions = readPositions_(config);
  let actualValue = Math.round(positions.reduce(function(s, p) { return s + p.valueRub; }, 0));

  writeBenchmarkToDashboard_(actualValue, hypotheticalValue, skipped);
}

function writeBenchmarkToDashboard_(actualValue, hypotheticalValue, skipped) {
  PropertiesService.getScriptProperties().setProperty(ANALYTICS_BENCH_PROP, JSON.stringify({
    actual: actualValue, hypothetical: hypotheticalValue, skipped: skipped,
  }));
}

// ── Единая точка входа: считает и XIRR, и бенчмарк, один раз перерисовывает ──
function calculateAnalytics() {
  withLock_('Рассчитать доходность', function() {
    calculateXIRR();
    calculateBenchmark();
    redrawAnalyticsSection_();
  });
}

function redrawAnalyticsSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;

  let props    = PropertiesService.getScriptProperties();
  let xirrRaw  = props.getProperty(ANALYTICS_XIRR_PROP);
  let benchRaw = props.getProperty(ANALYTICS_BENCH_PROP);
  if (!xirrRaw && !benchRaw) return;

  renderSection_(sh, ANALYTICS_SECTION_TITLE, function(sh, r, COLS) {
    if (xirrRaw) {
      let xirr  = JSON.parse(xirrRaw);
      let pct   = xirr.value;
      let bg    = pct === null ? '#9e9e9e' : pct >= 0 ? '#1b5e20' : '#b71c1c';
      let arrow = pct === null ? '' : pct >= 0 ? '▲' : '▼';
      let label = pct === null ? 'н/д' : arrow + ' ' + Math.abs(pct).toFixed(1) + '% годовых';
      r = renderTile_(sh, r, COLS, '📈 Доходность портфеля (XIRR)', label, bg);
    }

    if (benchRaw) {
      let b     = JSON.parse(benchRaw);
      let diff  = b.actual - b.hypothetical;
      let bg    = diff >= 0 ? '#1b5e20' : '#b71c1c';
      let arrow = diff >= 0 ? '▲' : '▼';
      r = renderTile_(sh, r, COLS, '📊 vs IMOEX', arrow + ' ' + rub_(Math.abs(diff)), bg);

      [['Ваш портфель', rub_(b.actual)], ['Если бы покупали IMOEX', rub_(b.hypothetical)]]
        .forEach(function(row) {
          sh.getRange(r, 1, 1, 3).merge().setValue(row[0]);
          sh.getRange(r, 4, 1, COLS - 3).merge().setValue(row[1])
            .setFontWeight('bold').setHorizontalAlignment('right');
          sh.getRange(r, 1, 1, COLS).setBackground(C.EVEN);
          r++;
        });

      if (b.skipped > 0) {
        sh.getRange(r, 1, 1, COLS).merge()
          .setValue('⚠️ ' + b.skipped + ' пополнений не удалось сопоставить с ценой')
          .setFontColor(C.WARN).setFontStyle('italic').setFontSize(9);
        r++;
      }
    }
    return r;
  });
}