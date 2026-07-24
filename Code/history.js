/**
 * history.js — История операций (с автопродолжением при таймауте)
 *
 * Данные копятся построчно в скрытом листе _history_raw_ops, пока не
 * обработаны все счета. Если скрипт близко к 6-минутному лимиту —
 * сохраняет прогресс, ставит одноразовый триггер на продолжение через
 * минуту и завершается. Финальная красивая таблица собирается только
 * после того, как все данные собраны целиком.
 *
 * Зависимости: tiFetch_(), moneyToNumber_() — tinvest.js
 *              C, DST, rub_(), mergedCell_(), hdrRow_() — dashboard.js
 *              withLock_() — lock.js
 */

const HISTORY_MONTHS = 12;
const HISTORY_RAW_SHEET       = '_history_raw_ops';
const HISTORY_PROGRESS_PROP   = 'HISTORY_FETCH_PROGRESS';
const HISTORY_RUNNING_PROP    = 'HISTORY_FETCH_RUNNING';
const HISTORY_TIME_LIMIT_MS   = 5 * 60 * 1000; // запас в минуту от лимита в 6

const OP_TYPES = {
  'OPERATION_TYPE_BUY':                  'Покупка',
  'OPERATION_TYPE_SELL':                 'Продажа',
  'OPERATION_TYPE_INPUT':                'Пополнение',
  'OPERATION_TYPE_OUTPUT':               'Вывод',
  'OPERATION_TYPE_COUPON':               'Купон',
  'OPERATION_TYPE_DIVIDEND':             'Дивиденд',
  'OPERATION_TYPE_BROKER_FEE':           'Комиссия',
  'OPERATION_TYPE_BUY_CARD':             'Покупка',
  'OPERATION_TYPE_SELL_CARD':            'Продажа',
  'OPERATION_TYPE_DIVIDEND_TAX':         'Налог (дивиденд)',
  'OPERATION_TYPE_INPUT_SECURITIES':     'Ввод бумаг',
  'OPERATION_TYPE_OUTPUT_SECURITIES':    'Вывод бумаг',
  'OPERATION_TYPE_BOND_REPAYMENT':       'Погашение облигации',
  'OPERATION_TYPE_BOND_REPAYMENT_TAX':   'Налог (погашение)',
};

const OP_SHOW = [
  'OPERATION_TYPE_BUY', 'OPERATION_TYPE_SELL',
  'OPERATION_TYPE_BUY_CARD', 'OPERATION_TYPE_SELL_CARD',
  'OPERATION_TYPE_INPUT', 'OPERATION_TYPE_OUTPUT',
  'OPERATION_TYPE_COUPON', 'OPERATION_TYPE_DIVIDEND',
  'OPERATION_TYPE_BROKER_FEE',
  'OPERATION_TYPE_BOND_REPAYMENT',
];

const OP_COLORS = {
  'Покупка':              '#e3f2fd',
  'Продажа':              '#e8f5e9',
  'Пополнение':           '#f3e5f5',
  'Вывод':                '#fff3e0',
  'Купон':                '#e8f5e9',
  'Дивиденд':             '#e8f5e9',
  'Комиссия':             '#fce4ec',
  'Погашение облигации':  '#e1f5fe',
};


// ════════════════════════════════════════════════════════════════════
// ТОЧКА ВХОДА С МЕНЮ
// ════════════════════════════════════════════════════════════════════

function updateHistorySheet() {
  let running = PropertiesService.getScriptProperties().getProperty(HISTORY_RUNNING_PROP);
  if (running === 'true') {
    SpreadsheetApp.getUi().alert(
      '⏳ Загрузка истории уже идёт в фоне (продолжится автоматически через триггер).\n' +
      'Просто подожди — таблица обновится сама, когда всё будет собрано.'
    );
    return;
  }
  withLock_('Обновить историю операций', function() {
    clearHistoryTriggers_();
    PropertiesService.getScriptProperties().deleteProperty(HISTORY_PROGRESS_PROP);
    PropertiesService.getScriptProperties().setProperty(HISTORY_RUNNING_PROP, 'true');
    historyFetchStep_();
  });
}

// Вызывается автоматически триггером при продолжении после паузы
function continueHistoryFetch_() {
  withLock_('Продолжение загрузки истории', historyFetchStep_);
}


// ════════════════════════════════════════════════════════════════════
// ОСНОВНОЙ ШАГ ЗАГРУЗКИ (с проверкой времени и возможной паузой)
// ════════════════════════════════════════════════════════════════════

function historyFetchStep_() {
  let startTime = new Date().getTime();
  let ss = SpreadsheetApp.getActive();
  let rawSh = getOrCreateRawHistorySheet_(ss);

  let accounts = getAccounts_();
  if (!accounts.length) {
    finishHistoryRun_(false, 'Не удалось получить список счетов.');
    return;
  }

  let toDate   = new Date();
  let fromDate = new Date(toDate.getTime() - HISTORY_MONTHS * 30 * 24 * 3600 * 1000);

  let progressRaw = PropertiesService.getScriptProperties().getProperty(HISTORY_PROGRESS_PROP);
  let progress = progressRaw ? JSON.parse(progressRaw) : { accountIdx: 0, cursor: '' };

  for (let ai = progress.accountIdx; ai < accounts.length; ai++) {
    let acc = accounts[ai];
    let cursor = ai === progress.accountIdx ? progress.cursor : '';

    while (true) {
      let page = fetchOperationsPage_(acc.id, acc.name, fromDate, toDate, cursor);
      if (page.ops.length) appendRawOps_(rawSh, page.ops);

      if (new Date().getTime() - startTime > HISTORY_TIME_LIMIT_MS) {
        let nextCursor = page.hasNext ? page.nextCursor : '';
        let nextAi = page.hasNext ? ai : ai + 1;
        PropertiesService.getScriptProperties().setProperty(HISTORY_PROGRESS_PROP,
          JSON.stringify({ accountIdx: nextAi, cursor: page.hasNext ? nextCursor : '' }));
        scheduleHistoryContinuation_();
        return; // пауза — продолжим по триггеру
      }

      if (page.hasNext && page.nextCursor) {
        cursor = page.nextCursor;
      } else {
        break; // счёт закончен, переходим к следующему
      }
    }
  }

  // Все счета обработаны — собираем финальную таблицу
  finishHistoryRun_(true, null);
}


// ════════════════════════════════════════════════════════════════════
// ОДНА СТРАНИЦА ЗАПРОСА (без внутреннего цикла — для резюмируемости)
// ════════════════════════════════════════════════════════════════════

function fetchOperationsPage_(accountId, accountName, fromDate, toDate, cursor) {
  let body = {
    accountId: accountId,
    from:      fromDate.toISOString(),
    to:        toDate.toISOString(),
    limit:     1000,
    operationTypes: OP_SHOW,
  };
  if (cursor) body.cursor = cursor;

  let resp;
  try {
    resp = tiFetch_(
      '/tinkoff.public.invest.api.contract.v1.OperationsService/GetOperationsByCursor',
      body
    );
  } catch (e) {
    console.warn('История: ошибка для счёта ' + accountId + ': ' + e.message);
    return { ops: [], hasNext: false, nextCursor: '' };
  }

  let items = resp.items || [];
  let ops = [];
  items.forEach(function(item) {
    let op = parseOperation_(item, accountId, accountName);
    if (op) ops.push(op);
  });

  return { ops: ops, hasNext: !!(resp.hasNext && resp.nextCursor), nextCursor: resp.nextCursor || '' };
}

function parseOperation_(item, accountId, accountName) {
  let type = item.type || item.operationType || '';
  if (!type) return null;

  let dateRaw = item.date;
  let date;
  if (typeof dateRaw === 'string') date = new Date(dateRaw);
  else if (dateRaw && dateRaw.seconds) date = new Date(Number(dateRaw.seconds) * 1000);
  else return null;
  if (isNaN(date.getTime())) return null;

  let amount = 0;
  if (item.payment)    amount = moneyToNumber_(item.payment);
  else if (item.price) amount = moneyToNumber_(item.price);

  let commission = null;
  if (item.commission && moneyToNumber_(item.commission) !== 0) {
    commission = Math.abs(moneyToNumber_(item.commission));
  }

  let price = null;
  if (item.price && item.quantity) price = moneyToNumber_(item.price);

  let quantity = null;
  if (item.quantity && Number(item.quantity) > 0) quantity = Number(item.quantity);

  let instrName = item.name || item.instrumentName || '';
  if (!instrName && item.figi) instrName = item.figi;

  return {
    date: date, accountId: accountId, accountName: accountName, type: type,
    instrumentName: instrName, quantity: quantity,
    price: price ? Math.round(price * 100) / 100 : null,
    amount: Math.round(Math.abs(amount) * 100) / 100,
    commission: commission ? Math.round(commission * 100) / 100 : null,
    note: item.description || '',
  };
}


// ════════════════════════════════════════════════════════════════════
// СКРЫТЫЙ ЛИСТ-НАКОПИТЕЛЬ (переживает паузы между запусками)
// ════════════════════════════════════════════════════════════════════

function getOrCreateRawHistorySheet_(ss) {
  let sh = ss.getSheetByName(HISTORY_RAW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(HISTORY_RAW_SHEET);
    sh.hideSheet();
  }
  return sh;
}

function appendRawOps_(sh, ops) {
  let rows = ops.map(function(op) {
    return [op.date, op.accountName, op.type, op.instrumentName,
            op.quantity, op.price, op.amount, op.commission, op.note];
  });
  if (!rows.length) return;
  let startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, 9).setValues(rows);
}

function readRawOps_(sh) {
  let lastRow = sh.getLastRow();
  if (lastRow < 1) return [];
  let data = sh.getRange(1, 1, lastRow, 9).getValues();
  return data.map(function(row) {
    return {
      date: row[0] instanceof Date ? row[0] : new Date(row[0]),
      accountName: row[1], type: row[2], instrumentName: row[3],
      quantity: row[4] || null, price: row[5] || null,
      amount: row[6], commission: row[7] || null, note: row[8] || '',
    };
  });
}


// ════════════════════════════════════════════════════════════════════
// ТРИГГЕР ПРОДОЛЖЕНИЯ
// ════════════════════════════════════════════════════════════════════

function scheduleHistoryContinuation_() {
  clearHistoryTriggers_();
  ScriptApp.newTrigger('continueHistoryFetch_')
    .timeBased().after(60 * 1000).create();
}

function clearHistoryTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'continueHistoryFetch_') {
      ScriptApp.deleteTrigger(t);
    }
  });
}


// ════════════════════════════════════════════════════════════════════
// ЗАВЕРШЕНИЕ: СБОРКА ФИНАЛЬНОЙ ТАБЛИЦЫ
// ════════════════════════════════════════════════════════════════════

function finishHistoryRun_(success, errorMsg) {
  let props = PropertiesService.getScriptProperties();
  clearHistoryTriggers_();
  props.deleteProperty(HISTORY_PROGRESS_PROP);
  props.deleteProperty(HISTORY_RUNNING_PROP);

  let ss = SpreadsheetApp.getActive();

  if (!success) {
    let sh = ss.getSheetByName(DST.HISTORY);
    if (!sh) sh = ss.insertSheet(DST.HISTORY);
    sh.getRange(1, 1).setValue('⚠️ ' + errorMsg);
    return;
  }

  let rawSh = ss.getSheetByName(HISTORY_RAW_SHEET);
  let allOps = rawSh ? readRawOps_(rawSh) : [];
  allOps.sort(function(a, b) { return b.date - a.date; });

  renderHistorySheet_(allOps);

  // Очищаем накопитель — данные уже перенесены в финальную таблицу
  if (rawSh) rawSh.clearContents();
}

function renderHistorySheet_(allOps) {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.HISTORY);
  if (!sh) sh = ss.insertSheet(DST.HISTORY);
  sh.clearContents();
  sh.clearFormats();

  let tz  = Session.getScriptTimeZone();
  let now = new Date();
  let nowStr = Utilities.formatDate(now, tz, 'dd.MM.yyyy HH:mm');
  let accounts = getAccounts_();

  let COLS = 9;
  let r = 1;

  mergedCell_(sh, r, 1, 1, COLS,
    '📋  ИСТОРИЯ ОПЕРАЦИЙ — последние ' + HISTORY_MONTHS + ' месяцев',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });
  r++;

  mergedCell_(sh, r, 1, 1, COLS,
    'Обновлено: ' + nowStr +
    '   ·   Счетов: ' + accounts.length +
    '   ·   Операций: ' + allOps.length +
    '   ·   Период: ' + HISTORY_MONTHS + ' мес. (менять в HISTORY_MONTHS)',
    { bg: '#263238', fg: '#b0bec5', align: 'center' });
  r++;

  let summary = buildSummary_(allOps);
  r = renderSummary_(sh, r, summary, COLS);
  r++;

  mergedCell_(sh, r, 1, 1, COLS, '▌ ВСЕ ОПЕРАЦИИ', { bg: C.MID, fg: '#ffffff', bold: true });
  r++;

  let tableStartRow = r;
  hdrRow_(sh, r,
    ['Дата', 'Счёт', 'Тип', 'Инструмент', 'Кол-во', 'Цена, ₽', 'Сумма, ₽', 'Комиссия, ₽', 'Примечание'],
    COLS);
  r++;

  allOps.forEach(function(op, idx) {
    let typeName = OP_TYPES[op.type] || op.type;
    let bg = OP_COLORS[typeName] || (idx % 2 === 0 ? C.EVEN : C.ODD);
    let isIncome = (typeName === 'Купон' || typeName === 'Дивиденд' ||
                    typeName === 'Пополнение' || typeName === 'Погашение облигации');
    let isCost   = (typeName === 'Комиссия' || typeName === 'Вывод');

    let dateCell = sh.getRange(r, 1);
    dateCell.setNumberFormat('@');
    dateCell.setValue(Utilities.formatDate(op.date, tz, 'dd.MM.yyyy HH:mm'));
    dateCell.setBackground(bg);
    sh.getRange(r, 2, 1, COLS - 1).setValues([[
      op.accountName, typeName, op.instrumentName || '—',
      op.quantity || '—', op.price || '', op.amount, op.commission || '', op.note || ''
    ]]).setBackground(bg);

    sh.getRange(r, 6).setNumberFormat('#,##0.00 [$₽-ru-RU]');
    sh.getRange(r, 7).setNumberFormat('#,##0.00 [$₽-ru-RU]');
    if (op.commission) sh.getRange(r, 8).setNumberFormat('#,##0.00 [$₽-ru-RU]');

    let amountCell = sh.getRange(r, 7);
    if (isIncome) amountCell.setFontColor('#1b5e20').setFontWeight('bold');
    else if (isCost) amountCell.setFontColor('#b71c1c');

    r++;
  });
 
 if (allOps.length > 0) {
    sh.getRange(tableStartRow + 1, 9, allOps.length, 1).setWrap(true); // колонка «Примечание»
  }
  try {
    let existingFilter = sh.getFilter();
    if (existingFilter) existingFilter.remove();
    sh.getRange(tableStartRow, 1, allOps.length + 1, COLS).createFilter();
  } catch (e) {
    console.warn('История: не удалось создать фильтр: ' + e.message);
  }

  [90, 120, 100, 220, 70, 120, 130, 120, 150].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(6);
  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════════════
// СВОДКА (без изменений)
// ════════════════════════════════════════════════════════════════════

function buildSummary_(ops) {
  let s = {
    totalIn: 0, totalOut: 0, totalCoupons: 0, totalDivs: 0,
    totalFees: 0, totalBuys: 0, totalSells: 0, totalRepay: 0,
  };
  ops.forEach(function(op) {
    let t = op.type;
    if (t === 'OPERATION_TYPE_INPUT')           s.totalIn      += op.amount;
    if (t === 'OPERATION_TYPE_OUTPUT')          s.totalOut     += op.amount;
    if (t === 'OPERATION_TYPE_COUPON')          s.totalCoupons += op.amount;
    if (t === 'OPERATION_TYPE_DIVIDEND')        s.totalDivs    += op.amount;
    if (t === 'OPERATION_TYPE_BROKER_FEE')      s.totalFees    += op.amount;
    if (t === 'OPERATION_TYPE_BUY' || t === 'OPERATION_TYPE_BUY_CARD')   s.totalBuys++;
    if (t === 'OPERATION_TYPE_SELL' || t === 'OPERATION_TYPE_SELL_CARD') s.totalSells++;
    if (t === 'OPERATION_TYPE_BOND_REPAYMENT')  s.totalRepay   += op.amount;
  });
  return s;
}

function renderSummary_(sh, r, s, COLS) {
  mergedCell_(sh, r, 1, 1, COLS, '▌ СВОДКА ЗА ПЕРИОД', { bg: C.MID, fg: '#ffffff', bold: true });
  r++;

  let summaryItems = [
    ['💰 Пополнения',          rub_(Math.round(s.totalIn)),      '#f3e5f5'],
    ['📤 Выводы',              rub_(Math.round(s.totalOut)),     '#fff3e0'],
    ['🏦 Купоны получено',     rub_(Math.round(s.totalCoupons)), '#e8f5e9'],
    ['📈 Дивиденды получено',  rub_(Math.round(s.totalDivs)),    '#e8f5e9'],
    ['🔄 Погашения облигаций', rub_(Math.round(s.totalRepay)),   '#e1f5fe'],
    ['💸 Комиссии уплачено',   rub_(Math.round(s.totalFees)),    '#fce4ec'],
    ['🛒 Покупок совершено',   s.totalBuys + ' сделок',          '#e3f2fd'],
    ['💹 Продаж совершено',    s.totalSells + ' сделок',         '#e8f5e9'],
  ];

  for (let i = 0; i < summaryItems.length; i += 2) {
    let left  = summaryItems[i];
    let right = summaryItems[i+1] || ['', '', C.EVEN];
    let half  = Math.floor(COLS / 2);

    sh.getRange(r, 1, 1, half).merge()
      .setValue(left[0] + ':  ' + left[1]).setBackground(left[2]).setFontWeight('bold');
    sh.getRange(r, half + 1, 1, COLS - half).merge()
      .setValue(right[0] + (right[1] ? ':  ' + right[1] : '')).setBackground(right[2]).setFontWeight('bold');
    r++;
  }

  r++;
  mergedCell_(sh, r, 1, 1, COLS, '▌ РЕАЛИЗОВАННЫЙ ПАССИВНЫЙ ДОХОД ЗА ПЕРИОД', { bg: '#2e7d32', fg: '#ffffff', bold: true });
  r++;

  let passiveIncome = Math.round(s.totalCoupons + s.totalDivs);
  let passiveRows = [
    ['🏦 Купоны по облигациям', Math.round(s.totalCoupons), '#e8f5e9', false],
    ['📈 Дивиденды по акциям',  Math.round(s.totalDivs),    '#e8f5e9', false],
    ['📅 ИТОГО пассивный доход', passiveIncome,             '#a5d6a7', true],
    ['🔄 Погашения облигаций',  Math.round(s.totalRepay),   '#e1f5fe', false],
  ];
  passiveRows.forEach(function(item) {
    sh.getRange(r, 1, 1, COLS).merge()
      .setValue(item[0] + ':   ' + rub_(item[1]))
      .setBackground(item[2]).setFontWeight(item[3] ? 'bold' : 'normal').setFontSize(item[3] ? 12 : 10);
    r++;
  });

  return r;
}


// ════════════════════════════════════════════════════════════════════
// ПОЛУЧЕНИЕ ОПЕРАЦИЙ (для остальных модулей — без изменений, узкие окна)
// ════════════════════════════════════════════════════════════════════

function fetchOperations_(accountId, accountName, fromDate, toDate) {
  let ops = [];
  let cursor = '';
  let maxPages = 20;

  while (maxPages-- > 0) {
    let page = fetchOperationsPage_(accountId, accountName, fromDate, toDate, cursor);
    ops = ops.concat(page.ops);
    if (page.hasNext && page.nextCursor) cursor = page.nextCursor;
    else break;
  }
  return ops;
}


// ════════════════════════════════════════════════════════════════════
// ПОЛУЧЕНИЕ СЧЕТОВ
// ════════════════════════════════════════════════════════════════════

function getAccounts_() {
  try {
    let resp = tiFetch_('/tinkoff.public.invest.api.contract.v1.UsersService/GetAccounts', {});
    let accounts = (resp.accounts || []).filter(function(a) { return a.status === 'ACCOUNT_STATUS_OPEN'; });
    return accounts.map(function(a) {
      return { id: a.id, name: a.name || (a.type === 'ACCOUNT_TYPE_TINKOFF_IIS' ? 'ИИС' : 'Брокерский') };
    });
  } catch (e) {
    console.error('История: ошибка получения счетов: ' + e.message);
    return [];
  }
}