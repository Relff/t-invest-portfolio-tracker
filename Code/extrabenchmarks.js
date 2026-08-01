/**
 * extrabenchmarks.js — Дополнительные бенчмарки: RGBI (индекс гособлигаций)
 * и золото (GLDRUB_TOM), в дополнение к сравнению с IMOEX (benchmark.js).
 *
 * Методология идентична benchmark.js: те же самые реальные пополнения
 * гипотетически инвестируются в этот бенчмарк вместо реального портфеля.
 * Это НЕ «доля облигаций против RGBI» — портфель разнородный, и honest
 * привязка конкретного пополнения к конкретной категории не отслеживается.
 * Это просто ещё одна точка сравнения «весь портфель против альтернативы»,
 * помимо уже существующего сравнения с IMOEX.
 *
 * Зависимости: getAccounts_(), fetchOperations_() — history.js
 *              readConfig_(), readPositions_() — dashboard.js
 *              tiFetch_() — tinvest.js
 *              getImoexPriceMap_(), findNearestPrice_() — benchmark.js
 *                (несмотря на название, обе функции универсальны — работают
 *                 по любому instrumentId, не только по IMOEX)
 *              renderSection_(), renderTile_() — sections.js
 */

const EXTRA_BENCH_PROP          = 'EXTRA_BENCH_DATA';
const EXTRA_BENCH_SECTION_TITLE = '▌ ДОПОЛНИТЕЛЬНЫЕ БЕНЧМАРКИ';
const EXTRA_BENCH_YEARS_BACK    = 5;

const EXTRA_BENCHMARKS_ = [
  { key: 'rgbi', label: 'RGBI (гособлигации)', kind: 'indicative', ticker: 'RGBI' },
  { key: 'gold', label: 'Золото (GLDRUB_TOM)',  kind: 'instrument', ticker: 'GLDRUB_TOM' },
];

function getIndicativeUid_(ticker) {
  let propKey = 'INDIC_UID_' + ticker;
  let cached  = PropertiesService.getScriptProperties().getProperty(propKey);
  if (cached) return cached;
  let resp = tiFetch_('/tinkoff.public.invest.api.contract.v1.InstrumentsService/Indicatives', {});
  let list = resp.instruments || [];
  let found = list.find(function(i) { return i.ticker === ticker; });
  if (!found) return null;
  PropertiesService.getScriptProperties().setProperty(propKey, found.uid);
  return found.uid;
}

function findInstrumentUid_(ticker) {
  let propKey = 'INSTR_UID_' + ticker;
  let cached  = PropertiesService.getScriptProperties().getProperty(propKey);
  if (cached) return cached;
  try {
    let resp = tiFetch_('/tinkoff.public.invest.api.contract.v1.InstrumentsService/FindInstrument', { query: ticker });
    let list = resp.instruments || [];
    let exact = list.find(function(i) { return i.ticker === ticker; }) || list[0];
    if (!exact) return null;
    PropertiesService.getScriptProperties().setProperty(propKey, exact.uid);
    return exact.uid;
  } catch (e) {
    return null;
  }
}

function calculateExtraBenchmarks() {
  let fromDate = new Date(new Date().getTime() - EXTRA_BENCH_YEARS_BACK * 365 * 24 * 3600 * 1000);
  let toDate   = new Date();

  let deposits = [];
  getAccounts_().forEach(function(acc) {
    fetchOperations_(acc.id, acc.name, fromDate, toDate).forEach(function(op) {
      if (op.type === 'OPERATION_TYPE_INPUT') deposits.push({ date: op.date, amount: op.amount });
    });
  });
  if (!deposits.length) {
    SpreadsheetApp.getUi().alert('Нет пополнений за последние ' + EXTRA_BENCH_YEARS_BACK + ' лет — сравнивать не с чем.');
    return;
  }

  let config      = readConfig_();
  let positions   = readPositions_(config);
  let actualValue = Math.round(positions.reduce(function(s, p) { return s + p.valueRub; }, 0));

  let results = [];
  EXTRA_BENCHMARKS_.forEach(function(bm) {
    let uid = bm.kind === 'indicative' ? getIndicativeUid_(bm.ticker) : findInstrumentUid_(bm.ticker);
    if (!uid) {
      results.push({ key: bm.key, label: bm.label, error: 'Инструмент не найден в T-Invest API' });
      return;
    }
    try {
      let priceMap      = getImoexPriceMap_(uid, fromDate, toDate);
      let currentPrice  = findNearestPrice_(priceMap, toDate);
      if (!currentPrice) {
        results.push({ key: bm.key, label: bm.label, error: 'Нет свежей цены за нужный период' });
        return;
      }
      let totalUnits = 0, skipped = 0;
      deposits.forEach(function(d) {
        let priceThen = findNearestPrice_(priceMap, d.date);
        if (priceThen && priceThen > 0) totalUnits += d.amount / priceThen; else skipped++;
      });
      results.push({
        key: bm.key, label: bm.label,
        hypothetical: Math.round(totalUnits * currentPrice),
        skipped: skipped,
      });
    } catch (e) {
      results.push({ key: bm.key, label: bm.label, error: e.message.substring(0, 60) });
    }
  });

  PropertiesService.getScriptProperties().setProperty(EXTRA_BENCH_PROP, JSON.stringify({
    actual: actualValue, results: results,
  }));
  redrawExtraBenchmarksSection_();
}

function redrawExtraBenchmarksSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(EXTRA_BENCH_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);

  renderSection_(sh, EXTRA_BENCH_SECTION_TITLE, function(sh, r, COLS) {
    d.results.forEach(function(res) {
      if (res.error) {
        sh.getRange(r, 1, 1, COLS).merge()
          .setValue('⚠️ ' + res.label + ' — ' + res.error)
          .setFontColor(C.WARN).setFontStyle('italic').setFontSize(9);
        r++;
        return;
      }

      let diff  = d.actual - res.hypothetical;
      let arrow = diff >= 0 ? '▲' : '▼';
      let bg    = diff >= 0 ? '#1b5e20' : '#b71c1c';
      r = renderTile_(sh, r, COLS, '📊 vs ' + res.label, arrow + ' ' + rub_(Math.abs(diff)), bg);

      [['Ваш портфель', rub_(d.actual)], ['Если бы покупали ' + res.label, rub_(res.hypothetical)]]
        .forEach(function(row) {
          sh.getRange(r, 1, 1, 3).merge().setValue(row[0]);
          sh.getRange(r, 4, 1, COLS - 3).merge().setValue(row[1])
            .setFontWeight('bold').setHorizontalAlignment('right');
          sh.getRange(r, 1, 1, COLS).setBackground(C.EVEN);
          r++;
        });

      if (res.skipped > 0) {
        sh.getRange(r, 1, 1, COLS).merge()
          .setValue('⚠️ ' + res.skipped + ' пополнений не удалось сопоставить с ценой')
          .setFontColor(C.WARN).setFontStyle('italic').setFontSize(9);
        r++;
      }
    });
    return r;
  });
}
