/**
 * history.js — История операций
 *
 * Автоматически подтягивает все операции со всех счетов
 * через T-Invest API (GetOperationsByCursor).
 *
 * Типы операций:
 *   Пополнение, Вывод, Покупка, Продажа, Купон, Дивиденд, Комиссия
 *
 * Добавить в onOpen():
 *   .addItem('📋  Обновить Историю операций', 'updateHistorySheet')
 *
 * Зависимости: tiFetch_(), moneyToNumber_() — tinvest.gs
 *              C, DST, rub_(), mergedCell_(), hdrRow_() — dashboard.gs
 */

// Название листа берётся из DST.HISTORY (dashboard.gs)
const HISTORY_MONTHS = 12; // глубина истории в месяцах (можно увеличить до 24)

// Маппинг типов операций из API → понятные названия
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

// Типы которые показываем (остальные фильтруем)
const OP_SHOW = [
  'OPERATION_TYPE_BUY', 'OPERATION_TYPE_SELL',
  'OPERATION_TYPE_BUY_CARD', 'OPERATION_TYPE_SELL_CARD',
  'OPERATION_TYPE_INPUT', 'OPERATION_TYPE_OUTPUT',
  'OPERATION_TYPE_COUPON', 'OPERATION_TYPE_DIVIDEND',
  'OPERATION_TYPE_BROKER_FEE',
  'OPERATION_TYPE_BOND_REPAYMENT',
];

// Цвета по типу операции
const OP_COLORS = {
  'Покупка':              '#e3f2fd',  // светло-синий
  'Продажа':              '#e8f5e9',  // светло-зелёный
  'Пополнение':           '#f3e5f5',  // светло-фиолетовый
  'Вывод':                '#fff3e0',  // светло-оранжевый
  'Купон':                '#e8f5e9',  // светло-зелёный
  'Дивиденд':             '#e8f5e9',  // светло-зелёный
  'Комиссия':             '#fce4ec',  // светло-красный
  'Погашение облигации':  '#e1f5fe',  // голубой
};


// ════════════════════════════════════════════════════════════════════
// ГЛАВНАЯ ФУНКЦИЯ
// ════════════════════════════════════════════════════════════════════

function updateHistorySheet() {
  let ss  = SpreadsheetApp.getActive();
  let sh  = ss.getSheetByName(DST.HISTORY);
  if (!sh) sh = ss.insertSheet(DST.HISTORY);
  sh.clearContents();
  sh.clearFormats();

  let tz  = Session.getScriptTimeZone();
  let now = new Date();
  let nowStr = Utilities.formatDate(now, tz, 'dd.MM.yyyy HH:mm');

  // Период выборки
  let fromDate = new Date(now.getTime() - HISTORY_MONTHS * 30 * 24 * 3600 * 1000);

  // Получаем все счета
  let accounts = getAccounts_();
  if (!accounts.length) {
    sh.getRange(1,1).setValue('⚠️ Не удалось получить список счетов.');
    return;
  }

  // Собираем операции со всех счетов
  let allOps = [];
  accounts.forEach(function(acc) {
    let ops = fetchOperations_(acc.id, acc.name, fromDate, now);
    allOps = allOps.concat(ops);
  });

  // Сортируем от новых к старым
  allOps.sort(function(a, b) { return b.date - a.date; });

  let COLS = 9;
  let r    = 1;

  // ── Шапка ─────────────────────────────────────────────────────────
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

  // ── Мини-сводка ──────────────────────────────────────────────────
  let summary = buildSummary_(allOps);
  r = renderSummary_(sh, r, summary, COLS);
  r++;

  // ── Таблица операций ─────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '▌ ВСЕ ОПЕРАЦИИ',
    { bg: C.MID, fg: '#ffffff', bold: true });
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

    // Надёжная запись даты: сначала формат @(текст), потом Utilities.formatDate
    // Прямая запись Date-объекта или строки → Sheets конвертирует в серийный номер
    let dateCell = sh.getRange(r, 1);
    dateCell.setNumberFormat('@');
    dateCell.setValue(Utilities.formatDate(op.date, tz, 'dd.MM.yyyy HH:mm'));
    dateCell.setBackground(bg);
    sh.getRange(r, 2, 1, COLS - 1).setValues([[
      op.accountName,
      typeName,
      op.instrumentName || '—',
      op.quantity || '—',
      op.price    || '',
      op.amount,
      op.commission || '',
      op.note || ''
    ]]).setBackground(bg);

    // Форматирование числовых колонок
    sh.getRange(r, 6).setNumberFormat('#,##0.00 [$₽-ru-RU]');
    sh.getRange(r, 7).setNumberFormat('#,##0.00 [$₽-ru-RU]');
    if (op.commission) sh.getRange(r, 8).setNumberFormat('#,##0.00 [$₽-ru-RU]');

    // Цвет суммы: доход = зелёный, расход = красный
    let amountCell = sh.getRange(r, 7);
    if (isIncome)     amountCell.setFontColor('#1b5e20').setFontWeight('bold');
    else if (isCost)  amountCell.setFontColor('#b71c1c');

    r++;
  });

  // Фильтр на таблицу
  try {
    let existingFilter = sh.getFilter();
    if (existingFilter) existingFilter.remove();
    sh.getRange(tableStartRow, 1, allOps.length + 1, COLS).createFilter();
  } catch(e) {
    console.warn('История: не удалось создать фильтр: ' + e.message);
  }

  // ── Ширина колонок ────────────────────────────────────────────────
  [90, 120, 100, 220, 70, 120, 130, 120, 150].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(6);
  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════════════
// ПОЛУЧЕНИЕ ОПЕРАЦИЙ ЧЕРЕЗ API (с пагинацией)
// ════════════════════════════════════════════════════════════════════

/**
 * Получает все операции по счёту за указанный период.
 * API возвращает данные страницами — обходим через cursor.
 */
function fetchOperations_(accountId, accountName, fromDate, toDate) {
  let ops    = [];
  let cursor = '';
  let maxPages = 20; // защита от бесконечного цикла

  while (maxPages-- > 0) {
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
    } catch(e) {
      console.warn('История: ошибка для счёта ' + accountId + ': ' + e.message);
      break;
    }

    let items = resp.items || [];
    items.forEach(function(item) {
      let op = parseOperation_(item, accountId, accountName);
      if (op) ops.push(op);
    });

    // Переходим на следующую страницу
    if (resp.hasNext && resp.nextCursor) {
      cursor = resp.nextCursor;
    } else {
      break;
    }
  }

  return ops;
}

/**
 * Парсит одну операцию из API-ответа.
 */
function parseOperation_(item, accountId, accountName) {
  let type = item.type || item.operationType || '';
  if (!type) return null;

  // Парсим дату
  let dateRaw = item.date;
  let date;
  if (typeof dateRaw === 'string') {
    date = new Date(dateRaw);
  } else if (dateRaw && dateRaw.seconds) {
    date = new Date(Number(dateRaw.seconds) * 1000);
  } else {
    return null;
  }
  if (isNaN(date.getTime())) return null;

  // Сумма операции
  let amount = 0;
  if (item.payment)    amount = moneyToNumber_(item.payment);
  else if (item.price) amount = moneyToNumber_(item.price);

  // Комиссия
  let commission = null;
  if (item.commission && moneyToNumber_(item.commission) !== 0) {
    commission = Math.abs(moneyToNumber_(item.commission));
  }

  // Цена за единицу
  let price = null;
  if (item.price && item.quantity) {
    price = moneyToNumber_(item.price);
  }

  // Количество бумаг
  let quantity = null;
  if (item.quantity && Number(item.quantity) > 0) {
    quantity = Number(item.quantity);
  }

  // Название инструмента
  let instrName = item.name || item.instrumentName || '';
  if (!instrName && item.figi) instrName = item.figi;

  return {
    date:           date,
    accountId:      accountId,
    accountName:    accountName,
    type:           type,
    instrumentName: instrName,
    quantity:       quantity,
    price:          price ? Math.round(price * 100) / 100 : null,
    amount:         Math.round(Math.abs(amount) * 100) / 100,
    commission:     commission ? Math.round(commission * 100) / 100 : null,
    note:           item.description || '',
  };
}


// ════════════════════════════════════════════════════════════════════
// МИНИ-СВОДКА
// ════════════════════════════════════════════════════════════════════

function buildSummary_(ops) {
  let s = {
    totalIn:       0,  // пополнения
    totalOut:      0,  // выводы
    totalCoupons:  0,  // купоны
    totalDivs:     0,  // дивиденды
    totalFees:     0,  // комиссии
    totalBuys:     0,  // покупки (кол-во сделок)
    totalSells:    0,  // продажи
    totalRepay:    0,  // погашения облигаций
  };

  ops.forEach(function(op) {
    let t = op.type;
    if (t === 'OPERATION_TYPE_INPUT')           s.totalIn      += op.amount;
    if (t === 'OPERATION_TYPE_OUTPUT')          s.totalOut     += op.amount;
    if (t === 'OPERATION_TYPE_COUPON')          s.totalCoupons += op.amount;
    if (t === 'OPERATION_TYPE_DIVIDEND')        s.totalDivs    += op.amount;
    if (t === 'OPERATION_TYPE_BROKER_FEE')      s.totalFees    += op.amount;
    if (t === 'OPERATION_TYPE_BUY' ||
        t === 'OPERATION_TYPE_BUY_CARD')        s.totalBuys++;
    if (t === 'OPERATION_TYPE_SELL' ||
        t === 'OPERATION_TYPE_SELL_CARD')       s.totalSells++;
    if (t === 'OPERATION_TYPE_BOND_REPAYMENT')  s.totalRepay   += op.amount;
  });

  return s;
}

function renderSummary_(sh, r, s, COLS) {
  mergedCell_(sh, r, 1, 1, COLS, '▌ СВОДКА ЗА ПЕРИОД',
    { bg: C.MID, fg: '#ffffff', bold: true });
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

  // Раскладываем по 2 в строку
  for (let i = 0; i < summaryItems.length; i += 2) {
    let left  = summaryItems[i];
    let right = summaryItems[i+1] || ['', '', C.EVEN];
    let half  = Math.floor(COLS / 2);

    sh.getRange(r, 1, 1, half).merge()
      .setValue(left[0] + ':  ' + left[1])
      .setBackground(left[2]).setFontWeight('bold');
    sh.getRange(r, half + 1, 1, COLS - half).merge()
      .setValue(right[0] + (right[1] ? ':  ' + right[1] : ''))
      .setBackground(right[2]).setFontWeight('bold');
    r++;
  }

  // ── Блок реализованного пассивного дохода ─────────────────────────
  r++;
  let passiveTotal = Math.round(s.totalCoupons + s.totalDivs + s.totalRepay); // используется в summary tiles
  mergedCell_(sh, r, 1, 1, COLS,
    '▌ РЕАЛИЗОВАННЫЙ ПАССИВНЫЙ ДОХОД ЗА ПЕРИОД',
    { bg: '#2e7d32', fg: '#ffffff', bold: true });
  r++;

  let passiveIncome = Math.round(s.totalCoupons + s.totalDivs); // только доход, без возврата капитала
  let passiveRows = [
    ['🏦 Купоны по облигациям',  Math.round(s.totalCoupons), '#e8f5e9', false],
    ['📈 Дивиденды по акциям',   Math.round(s.totalDivs),    '#e8f5e9', false],
    ['📅 ИТОГО пассивный доход', passiveIncome,              '#a5d6a7', true],
    ['🔄 Погашения облигаций',   Math.round(s.totalRepay),   '#e1f5fe', false],  // возврат номинала, не доход
  ];

  passiveRows.forEach(function(item) {
    sh.getRange(r, 1, 1, COLS).merge()
      .setValue(item[0] + ':   ' + rub_(item[1]))
      .setBackground(item[2])
      .setFontWeight(item[3] ? 'bold' : 'normal')
      .setFontSize(item[3] ? 12 : 10);
    r++;
  });

  return r;
}


// ════════════════════════════════════════════════════════════════════
// ПОЛУЧЕНИЕ СЧЕТОВ
// ════════════════════════════════════════════════════════════════════

function getAccounts_() {
  try {
    let resp     = tiFetch_('/tinkoff.public.invest.api.contract.v1.UsersService/GetAccounts', {});
    let accounts = (resp.accounts || []).filter(function(a) {
      return a.status === 'ACCOUNT_STATUS_OPEN';
    });
    return accounts.map(function(a) {
      return {
        id:   a.id,
        name: a.name || (a.type === 'ACCOUNT_TYPE_TINKOFF_IIS' ? 'ИИС' : 'Брокерский'),
      };
    });
  } catch(e) {
    console.error('История: ошибка получения счетов: ' + e.message);
    return [];
  }
}
