/**
 * xirr.js — Расчёт XIRR (доходность с учётом дат и сумм пополнений)
 *
 * Зависимости: getAccounts_(), fetchOperations_() — history.js
 *              readConfig_(), readPositions_() — dashboard.js
 *              redrawAnalyticsSection_(), ANALYTICS_XIRR_PROP — benchmark.js
 */

const XIRR_HELPER_SHEET = '_XIRR_calc';
const XIRR_YEARS_BACK   = 5;
/* Produced by Relferium — t-invest-portfolio-tracker */

/**
 * Чистый расчёт XIRR без побочных эффектов (UI/Properties) — переиспользуется
 * как в calculateXIRR() (аналитика на Дашборде), так и в snapshot.js
 * (накопление истории XIRR по датам).
 *
 * Возвращает % годовых (число) или null, если данных недостаточно.
 */
function computeXirrValue_() {
  let ss = SpreadsheetApp.getActive();

  let fromDate = new Date(new Date().getTime() - XIRR_YEARS_BACK * 365 * 24 * 3600 * 1000);
  let toDate   = new Date();

  let flows = [];

  getAccounts_().forEach(function(acc) {
    let ops = fetchOperations_(acc.id, acc.name, fromDate, toDate);
    ops.forEach(function(op) {
      if (op.type === 'OPERATION_TYPE_INPUT')  flows.push({ date: op.date, amount: -op.amount });
      if (op.type === 'OPERATION_TYPE_OUTPUT') flows.push({ date: op.date, amount:  op.amount });
    });
  });

  if (flows.length < 1) return null;

  let config    = readConfig_();
  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);
  flows.push({ date: toDate, amount: totalRub });

  flows.sort(function(a, b) { return a.date - b.date; });

  let sh = ss.getSheetByName(XIRR_HELPER_SHEET);
  if (!sh) sh = ss.insertSheet(XIRR_HELPER_SHEET);
  sh.clearContents();

  sh.getRange(1, 1, flows.length, 1).setValues(flows.map(function(f) { return [f.amount]; }));
  sh.getRange(1, 2, flows.length, 1).setValues(flows.map(function(f) { return [f.date]; }));
  sh.getRange(1, 2, flows.length, 1).setNumberFormat('dd.mm.yyyy');

  let formulaCell = sh.getRange(1, 4);
  formulaCell.setFormula('=XIRR(A1:A' + flows.length + '; B1:B' + flows.length + ')');
  SpreadsheetApp.flush();

  let xirrValue = formulaCell.getValue();
  sh.hideSheet();

  return typeof xirrValue === 'number' ? xirrValue * 100 : null;
}

function calculateXIRR() {
  let pct = computeXirrValue_();
  if (pct === null) {
    SpreadsheetApp.getUi().alert('Недостаточно данных: нет ни одного пополнения за последние ' + XIRR_YEARS_BACK + ' лет, либо XIRR не удалось посчитать.');
  }
  writeXirrToDashboard_(pct);
}

function writeXirrToDashboard_(xirrValue) {
  let pct = typeof xirrValue === 'number' ? xirrValue : null;
  PropertiesService.getScriptProperties().setProperty(ANALYTICS_XIRR_PROP, JSON.stringify({ value: pct }));
}