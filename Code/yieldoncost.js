/**
 * yieldoncost.js — Yield on Cost (доход к себестоимости, не к текущей цене)
 *
 * Зависит от уже посчитанной себестоимости из avgprice.js — сначала
 * запусти «Средняя цена и P/L», потом это.
 *
 * Зависимости: readConfig_(), readPositions_(), rub_(), DST, C — dashboard.js
 *              readDividendsFromConfig_(), buildFigiMap_(), findDividend_(),
 *              fetchDividendProjection_() — income.js
 *              AVGPRICE_PROP — avgprice.js
 *              renderSection_() — sections.js
 */

const YOC_PROP          = 'YIELD_ON_COST_DATA';
const YOC_SECTION_TITLE = '▌ YIELD ON COST';

function calculateYieldOnCost() {
  let avgRaw = PropertiesService.getScriptProperties().getProperty(AVGPRICE_PROP);
  if (!avgRaw) {
    SpreadsheetApp.getUi().alert('Сначала запусти «Средняя цена и P/L» — Yield on Cost использует те же данные о себестоимости.');
    return;
  }
  let avgData = JSON.parse(avgRaw);

  let config    = readConfig_();
  let positions = readPositions_(config);
  let shares    = positions.filter(function(p) { return p.category === 'Акции'; });
  let divMap    = readDividendsFromConfig_();
  let figiMap   = buildFigiMap_();

  let totalCost = 0, totalIncome = 0;
  let rows = [];

  avgData.forEach(function(a) {
    if (!(a.avgPrice > 0)) return;
    let pos = shares.find(function(p) { return p.name === a.name; });
    if (!pos) return;

    let div = findDividend_(pos.name, pos.ticker, divMap);
    if (div.amount <= 0) {
      let figi = pos.figi || figiMap[pos.name] || figiMap[pos.ticker] || '';
      if (figi) {
        let hist = fetchDividendProjection_(figi);
        if (hist.perUnit > 0) div = { amount: hist.perUnit };
      }
    }

    let costBasis = a.avgPrice * a.qty;
    let income    = div.amount * a.qty;
    totalCost   += costBasis;
    totalIncome += income;

    rows.push({ name: a.name, yoc: costBasis > 0 ? income / costBasis : 0 });
  });

  let overallYoc = totalCost > 0 ? totalIncome / totalCost : 0;

  PropertiesService.getScriptProperties().setProperty(YOC_PROP, JSON.stringify({
    overall: overallYoc, rows: rows,
  }));
  redrawYocSection_();
}

function redrawYocSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(YOC_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);

  renderSection_(sh, YOC_SECTION_TITLE, function(sh, r, COLS, colStart) {
    sh.getRange(r, colStart, 1, 3).merge().setValue('Yield on Cost (весь пакет акций)').setFontWeight('bold');
    sh.getRange(r, colStart + 3, 1, COLS - 3).merge().setValue((d.overall * 100).toFixed(1) + '%')
      .setFontColor(C.OK).setFontWeight('bold').setHorizontalAlignment('right');
    sh.getRange(r, colStart, 1, COLS).setBackground(C.EVEN);
    r++;

    d.rows.sort(function(a, b) { return b.yoc - a.yoc; }).forEach(function(row, idx) {
      let bg = idx % 2 === 0 ? C.EVEN : C.ODD;
      sh.getRange(r, colStart, 1, 3).merge().setValue(row.name).setBackground(bg);
      sh.getRange(r, colStart + 3, 1, COLS - 3).merge().setValue((row.yoc * 100).toFixed(1) + '%')
        .setFontWeight('bold').setHorizontalAlignment('right').setBackground(bg);
      r++;
    });
  }, 'right');
}