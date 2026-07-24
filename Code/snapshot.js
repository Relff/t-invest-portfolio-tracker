/**
 * snapshot.js — Ежедневный снимок баланса портфеля (с разбивкой по категориям)
 *
 * Пишет одну строку в день на скрытый лист "История баланса":
 * Дата | Общая стоимость | Пополнено с начала | Доход с начала | Акции | Облигации | Золото | Денежный рынок
 *
 * Совместимость: старые строки (записанные до этого расширения) имеют
 * только первые 4 колонки — новые колонки там останутся пустыми,
 * это ожидаемо и не является ошибкой.
 *
 * Зависимости: readConfig_(), readPositions_() — dashboard.js
 *              getAccounts_(), fetchOperations_() — history.js
 */

const SNAPSHOT_SHEET = 'История баланса';
const SNAPSHOT_CATEGORIES = ['Акции', 'Облигации', 'Золото', 'Денежный рынок'];

function recordDailySnapshot() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SNAPSHOT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SNAPSHOT_SHEET);
    sh.appendRow([
      'Дата', 'Общая стоимость, ₽', 'Пополнено с начала, ₽', 'Доход с начала, ₽',
      'Акции, ₽', 'Облигации, ₽', 'Золото, ₽', 'Денежный рынок, ₽'
    ]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  } else {
    // Если лист существует со старой версии — доращиваем заголовок новыми колонками
    let header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (header.length < 8) {
      sh.getRange(1, 5, 1, 4).setValues([[
        'Акции, ₽', 'Облигации, ₽', 'Золото, ₽', 'Денежный рынок, ₽'
      ]]);
    }
  }

  let tz  = Session.getScriptTimeZone();
  let today = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy');

  let lastRow = sh.getLastRow();
  if (lastRow > 1) {
    let lastCell = sh.getRange(lastRow, 1).getValue();
    let lastDateStr = lastCell instanceof Date
      ? Utilities.formatDate(lastCell, tz, 'dd.MM.yyyy')
      : String(lastCell);
    if (lastDateStr === today) return;
  }

  let config    = readConfig_();
  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);

  // Разбивка по категориям на сегодняшний момент
  let catSums = SNAPSHOT_CATEGORIES.map(function(cat) {
    return Math.round(
      positions.filter(function(p) { return p.category === cat; })
                .reduce(function(s, p) { return s + p.valueRub; }, 0)
    );
  });

  let fromDate = lastRow > 1
    ? (function() {
        let raw = sh.getRange(lastRow, 1).getValue();
        return raw instanceof Date ? raw : new Date(raw.split('.').reverse().join('-'));
      })()
    : new Date(new Date().getTime() - 365 * 24 * 3600 * 1000);
  let toDate = new Date();

  let prevContrib = lastRow > 1 ? sh.getRange(lastRow, 3).getValue() : 0;
  let prevIncome  = lastRow > 1 ? sh.getRange(lastRow, 4).getValue() : 0;

  let deltaContrib = 0;
  let deltaIncome  = 0;

  getAccounts_().forEach(function(acc) {
    let ops = fetchOperations_(acc.id, acc.name, fromDate, toDate);
    ops.forEach(function(op) {
      if (op.type === 'OPERATION_TYPE_INPUT')  deltaContrib += op.amount;
      if (op.type === 'OPERATION_TYPE_OUTPUT') deltaContrib -= op.amount;
      if (op.type === 'OPERATION_TYPE_COUPON' || op.type === 'OPERATION_TYPE_DIVIDEND')
        deltaIncome += op.amount;
    });
  });

  sh.appendRow([
    today,
    Math.round(totalRub),
    Math.round(prevContrib + deltaContrib),
    Math.round(prevIncome + deltaIncome),
    catSums[0], catSums[1], catSums[2], catSums[3],
  ]);
}