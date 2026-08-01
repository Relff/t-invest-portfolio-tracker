const AVGPRICE_PROP          = 'AVGPRICE_DATA';
const AVGPRICE_SECTION_TITLE = '▌ СРЕДНЯЯ ЦЕНА И P/L (АКЦИИ)';

function calculateAveragePriceAndPL() {
  withLock_('Средняя цена и P/L', calculateAveragePriceAndPL_impl_);
}

function calculateAveragePriceAndPL_impl_() {
  let config;
  try { config = readConfig_(); }
  catch (e) { SpreadsheetApp.getUi().alert('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  let shares = positions.filter(function(p) { return p.category === 'Акции'; });
  if (!shares.length) { SpreadsheetApp.getUi().alert('Нет акций в портфеле.'); return; }

  let params   = readAdvancedParams_();
  let fromDate = new Date(new Date().getTime() - params.fifoYears * 365 * 24 * 3600 * 1000);
  let toDate   = new Date();

  let allOps = [];
  getAccounts_().forEach(function(acc) {
    fetchOperations_(acc.id, acc.name, fromDate, toDate).forEach(function(op) {
      let isTrade = op.type === 'OPERATION_TYPE_BUY'  || op.type === 'OPERATION_TYPE_BUY_CARD' ||
                    op.type === 'OPERATION_TYPE_SELL' || op.type === 'OPERATION_TYPE_SELL_CARD';
      if (isTrade && op.quantity && op.price) allOps.push(op);
    });
  });

  let results = shares.map(function(p) {
    let myOps = allOps.filter(function(op) { return matchesName_(op.instrumentName, p.name); })
                       .sort(function(a, b) { return a.date - b.date; });
    let lots = buildFifoLots_(myOps);

    let fifoQty   = lots.reduce(function(s, l) { return s + l.qty; }, 0);
    let costBasis = lots.reduce(function(s, l) { return s + l.qty * l.price; }, 0);
    let avgPrice  = fifoQty > 0 ? costBasis / fifoQty : 0;

    return {
      name: p.name, qty: p.qty, avgPrice: avgPrice, curPrice: p.price,
      pl: avgPrice > 0 ? (p.price - avgPrice) * p.qty : null,
      plPct: avgPrice > 0 ? (p.price - avgPrice) / avgPrice : null,
      mismatch: Math.abs(fifoQty - p.qty) > 0.001,
      fifoQty: fifoQty, opsFound: myOps.length,
    };
  });

  PropertiesService.getScriptProperties().setProperty(AVGPRICE_PROP, JSON.stringify(results));
  redrawAvgPriceSection_();
}

/**
 * Строит список открытых FIFO-лотов (с датой покупки каждого) по истории
 * сделок одной бумаги. Общая логика для avgprice.js и ldv.js — держите
 * их в синхронизации, если меняете расчёт FIFO.
 */
function buildFifoLots_(myOps) {
  let lots = [];
  myOps.forEach(function(op) {
    if (op.type.indexOf('BUY') >= 0) {
      lots.push({ qty: op.quantity, price: op.price, date: op.date });
    } else {
      let toSell = op.quantity;
      while (toSell > 0 && lots.length) {
        let lot = lots[0];
        let used = Math.min(lot.qty, toSell);
        lot.qty -= used; toSell -= used;
        if (lot.qty <= 0) lots.shift();
      }
    }
  });
  return lots;
}

function matchesName_(opName, posName) {
  if (!opName || !posName) return false;
  let a = opName.toLowerCase().trim(), b = posName.toLowerCase().trim();
  return a === b || a.indexOf(b) >= 0 || b.indexOf(a) >= 0;
}

function redrawAvgPriceSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(AVGPRICE_PROP);
  if (!raw) return;
  let results = JSON.parse(raw);
  let params = readAdvancedParams_();

  renderSection_(sh, AVGPRICE_SECTION_TITLE, function(sh, r, COLS) {
    sh.getRange(r, 1).setValue('Акция');
    sh.getRange(r, 2).setValue('Ср. цена');
    sh.getRange(r, 3).setValue('Тек. цена');
    sh.getRange(r, 4, 1, 2).merge().setValue('P/L, ₽ / %');
    sh.getRange(r, 6).setValue('Примечание');
    sh.getRange(r, 1, 1, COLS).setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold');
    r++;

    results.forEach(function(res, idx) {
      let bg = idx % 2 === 0 ? C.EVEN : C.ODD;
      sh.getRange(r, 1).setValue(res.name).setBackground(bg);

      if (res.avgPrice > 0) {
        sh.getRange(r, 2).setValue(res.avgPrice).setNumberFormat('#,##0.00 [$₽-ru-RU]').setBackground(bg);
        sh.getRange(r, 3).setValue(res.curPrice).setNumberFormat('#,##0.00 [$₽-ru-RU]').setBackground(bg);
        let clr = res.pl >= 0 ? C.OK : C.CRIT;
        sh.getRange(r, 4, 1, 2).merge()
          .setValue(rub_(res.pl) + '  (' + (res.plPct * 100 >= 0 ? '+' : '') + (res.plPct * 100).toFixed(1) + '%)')
          .setFontColor(clr).setFontWeight('bold').setBackground(bg);
      } else {
        sh.getRange(r, 2, 1, 4).merge().setValue('Нет истории сделок за ' + params.fifoYears + ' лет')
          .setFontColor(C.SKIP).setFontStyle('italic').setBackground(bg);
      }

      let note = res.mismatch
        ? '⚠️ FIFO-кол-во (' + res.fifoQty + ') ≠ факт (' + res.qty + ')'
        : (res.opsFound + ' сделок в истории');
      sh.getRange(r, 6).setValue(note)
        .setFontColor(res.mismatch ? C.WARN : '#9e9e9e')
        .setFontStyle('italic').setFontSize(9).setBackground(bg);
      r++;
    });
  });
}