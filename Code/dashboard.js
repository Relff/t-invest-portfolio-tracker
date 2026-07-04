/**
 * ╔═══════════════════════════════════════════════════════════════════╗
 * ║  dashboard.js — Дашборд портфеля                                 ║
 * ║  Работает поверх Jeanorochka/tinvest-google-sheets-portfolio     ║
 * ║                                                                   ║
 * ║  УСТАНОВКА (5 шагов):                                            ║
 * ║  1. В Google Sheets: Расширения → Apps Script                    ║
 * ║  2. Создайте новый файл: + → Скрипт → назовите "dashboard"      ║
 * ║  3. Вставьте весь этот код                                       ║
 * ║  4. В файле tinvest.js найдите функцию onOpen() и удалите её    ║
 * ║     (dashboard.js создаёт общее меню вместо неё)                ║
 * ║  5. Обновите страницу — появится меню "Tinkoff"                  ║
 * ║                                                                   ║
 * ║  ПЕРВЫЙ ЗАПУСК:                                                   ║
 * ║  Tinkoff → ⚙️ Инициализировать Config                            ║
 * ║  Tinkoff → 🔄 Синхронизировать + обновить всё                    ║
 * ║                                                                   ║
 * ║  ТРИГГЕР (автообновление раз в час):                             ║
 * ║  Apps Script → Триггеры → + → syncAndRefresh →                  ║
 * ║  Time-driven → Hour timer → Every hour                           ║
 * ╚═══════════════════════════════════════════════════════════════════╝
 */

// ════════════════════════════════════════════════════════════════════
// 1. КОНСТАНТЫ
// ════════════════════════════════════════════════════════════════════

/** Листы из tinvest.js — НЕ переименовывать */
const SRC = {
  SHARES:  'Дан_Акции',
  BONDS:   'Дан_Облигации',
  ETFS:    'Дан_Фонды',
  MONEY:   'Дан_Деньги',
};

/** Листы дашборда */
const DST = {
  CONFIG:    'Настройки',
  DASHBOARD: 'Дашборд',
  REBALANCE: 'Ребалансировка',
  INCOME:    'Ожидаемый доход',
  CALENDAR:  'Календарь выплат',
  HISTORY:   'История операций',
  POSITIONS: 'Позиции',
};

/** Цвета */
const C = {
  DARK:     '#1a237e',   // тёмно-синий (шапка)
  MID:      '#283593',   // синий (секции)
  OK:       '#1b5e20',   // зелёный  — отклонение ≤ 1.5 пп
  WARN:     '#e65100',   // оранжевый — 1.5–3 пп
  CRIT:     '#b71c1c',   // красный   — > 3 пп
  ODD:      '#f5f5f5',
  EVEN:     '#ffffff',
  SKIP:     '#9e9e9e',   // серый — пропущенная позиция
  INPUT:    '#fff9c4',   // жёлтый — ячейка ввода
};

/** Пороги отклонения (процентных пунктов) */
const THR = { OK: 1.5, WARN: 3.0 };


// ════════════════════════════════════════════════════════════════════
// 2. МЕНЮ (заменяет onOpen из tinvest.js)
// ════════════════════════════════════════════════════════════════════


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tinkoff')
    .addItem('🔄  Синхронизировать позиции (T-Invest API)',   'syncTinkoffPositions')
    .addSeparator()
    .addItem('⚙️  Инициализировать Config (первый запуск)',   'initConfig')
    .addItem('📊  Обновить Dashboard',                        'updateDashboard')
    .addItem('💰  Пересчитать калькулятор пополнения',        'calculateRebalance')
    .addSeparator()
    .addItem('🚀  Синхронизировать + обновить всё',           'syncAndRefresh')
    .addSeparator()
    .addItem('💵  Обновить Ожидаемый доход',                  'updateIncomeSheet')
    .addItem('➕  Добавить блок дивидендов в Config',         'addDividendsBlock')
    .addSeparator()
    .addItem('📅  Обновить Календарь выплат',                 'updateCalendarSheet')
    .addItem('📋  Обновить Историю операций', 'updateHistorySheet')
    .addToUi();
}

// ════════════════════════════════════════════════════════════════════
// 3. ИНИЦИАЛИЗАЦИЯ CONFIG
// ════════════════════════════════════════════════════════════════════

function initConfig() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);

  if (sh) {
    let ans = SpreadsheetApp.getUi().alert(
      'Лист Config уже существует. Перезаписать?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (ans !== SpreadsheetApp.getUi().Button.YES) return;
    sh.clearContents();
    sh.clearFormats();
  } else {
    sh = ss.insertSheet(DST.CONFIG);
  }

  // ── Данные ────────────────────────────────────────────────────────
  let rows = [
    // Блок 1: Целевые доли по классам
    ['▌ ЦЕЛЕВАЯ СТРУКТУРА ПОРТФЕЛЯ (% от всего портфеля)', '', ''],
    ['Категория',        'Цель %', 'Счёт'],
    ['Акции',            40,       'ИИС-3'],
    ['Облигации',        40,       'ИИС-3'],
    ['Золото',           10,       'ИИС-3'],
    ['Замещайки',         5,       'ИИС-3'],
    ['Денежный рынок',    5,       'Брокерский'],
    ['', '', ''],

    // Блок 2: Целевые доли акций
    ['▌ ЦЕЛЕВЫЕ ДОЛИ АКЦИЙ (% от всего портфеля)', '', ''],
    ['Название (точно как в Positions_Aggregated)',  'Цель %', 'Тикер'],
    ['Сбербанк',                    6,     'SBER'],
    ['Лукойл',                      6,     'LKOH'],
    ['Т-Технологии',                5,     'T'],
    ['Полюс',                       5,     'PLZL'],
    ['ЯНДЕКС ао01',                 4.5,   'YDEX'],
    ['КЦ ИКС 5',                    4,     'X5'],
    ['ФосАгро',                     3.5,   'PHOR'],
    ['Positive Technologies',       3,     'POSI'],
    ['Озон ао01',                   3,     'OZON'],
    ['', '', ''],

    // Блок 3: Маппинг спецкатегорий
    ['▌ МАППИНГ СПЕЦКАТЕГОРИЙ — заполните точные названия из Positions_Aggregated', '', ''],
    ['Название инструмента',                    'Категория',         'Комментарий'],
    ['паи ЗПИФ нд Т-КапитЛужникиКолл',         'Золото',            '← ВИМ фонд золота'],
    ['',                                         'Замещайки',         '← Укажите замещающую облигацию'],
    ['',                                         'Денежный рынок',    '← Укажите фонд денежного рынка'],
  ];

  sh.getRange(1, 1, rows.length, 3).setValues(rows);

  // Форматирование заголовков блоков
  [[1, C.DARK], [9, C.DARK], [21, C.DARK]].forEach(function(pair) {
    sh.getRange(pair[0], 1, 1, 3).merge()
      .setBackground(pair[1]).setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(11);
  });

  // Форматирование строк-заголовков таблиц
  [[2, C.MID], [10, C.MID], [22, C.MID]].forEach(function(pair) {
    sh.getRange(pair[0], 1, 1, 3)
      .setBackground(pair[1]).setFontColor('#ffffff').setFontWeight('bold');
  });

  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 200);

  // Автоматически добавляем Блок 4 (дивиденды) при первоначальной настройке
  if (typeof addDividendsBlock === 'function') {
    addDividendsBlock();
  }

  SpreadsheetApp.getUi().alert(
    '✅ Config создан! Блок дивидендов добавлен автоматически.\n\n' +
    '⚠️  Важно: проверьте точные названия инструментов в Блоке 3 —\n' +
    'они должны совпадать с колонкой "name" в листе Позиции.'
  );
}


// ════════════════════════════════════════════════════════════════════
// 4. ЧТЕНИЕ CONFIG
// ════════════════════════════════════════════════════════════════════

function readConfig_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) throw new Error('Лист Config не найден. Запустите initConfig().');

  let v = sh.getDataRange().getValues();

  // Классы: строки 3–7 (индексы 2–6)
  let classTargets = {};
  for (let i = 2; i <= 6 && i < v.length; i++) {
    let cat = String(v[i][0]).trim();
    let pct = Number(v[i][1]);
    if (cat && !isNaN(pct) && pct > 0) classTargets[cat] = pct / 100;
  }

  // Акции: строки 11–19 (индексы 10–18)
  let stockTargets = {};
  for (let j = 10; j <= 18 && j < v.length; j++) {
    let sName = String(v[j][0]).trim();
    let sPct  = Number(v[j][1]);
    if (sName && !isNaN(sPct) && sPct > 0) stockTargets[sName] = sPct / 100;
  }

  // Маппинг: строки 23+ (индексы 22+)
  let mapping = {};
  for (let k = 22; k < v.length; k++) {
    let mName = String(v[k][0]).trim();
    if (mName.indexOf('ДИВИДЕНДЫ') >= 0) break;
    let mCat  = String(v[k][1]).trim();
    if (mName && mCat) mapping[mName] = mCat;
  }

  return { classTargets: classTargets, stockTargets: stockTargets, mapping: mapping };
}


// ════════════════════════════════════════════════════════════════════
// 5. ЧТЕНИЕ ПОЗИЦИЙ ИЗ ЛИСТОВ TINVEST.JS
// ════════════════════════════════════════════════════════════════════

function readPositions_(config) {
  let ss = SpreadsheetApp.getActive();
  let sources = [
    { sheet: SRC.SHARES, defaultCat: 'Акции' },
    { sheet: SRC.BONDS,  defaultCat: 'Облигации' },
    { sheet: SRC.ETFS,   defaultCat: 'Золото' },
    { sheet: SRC.MONEY,  defaultCat: 'Денежный рынок' },
  ];

  let positions = [];

  for (let s = 0; s < sources.length; s++) {
    let src = sources[s];
    let sh  = ss.getSheetByName(src.sheet);
    if (!sh) continue;

    let data = sh.getDataRange().getValues();
    if (data.length < 2) continue;

    // Индексируем заголовок
    let H = {};
    for (let hi = 0; hi < data[0].length; hi++) {
      H[String(data[0][hi]).trim()] = hi;
    }

    for (let ri = 1; ri < data.length; ri++) {
      let row      = data[ri];
      let name     = String(row[H['name']]   || '').trim();
      let ticker   = String(row[H['ticker']] || '').trim();
      let valueRub = Number(row[H['position_value_rub']]             || 0);
      let price    = Number(row[H['current_price_rub_per_piece']]    || 0);
      let qty      = Number(row[H['quantity_pcs']]                   || 0);
      let lot      = Number(row[H['lot']]    || 1);

      if (!name || valueRub === 0) continue;

      // Категория: сначала маппинг из Config, потом дефолт по типу
      let category = config.mapping[name] || src.defaultCat;

      positions.push({
        name:     name,
        ticker:   ticker,
        category: category,
        valueRub: valueRub,
        price:    price,
        qty:      qty,
        lot:      lot,
      });
    }
  }

  return positions;
}


// ════════════════════════════════════════════════════════════════════
// 6. ОБНОВЛЕНИЕ DASHBOARD
// ════════════════════════════════════════════════════════════════════

function updateDashboard() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) sh = ss.insertSheet(DST.DASHBOARD);
  sh.clearContents();
  sh.clearFormats();

  let config;
  try { config = readConfig_(); }
  catch (e) { sh.getRange(1, 1).setValue('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  if (!positions.length) {
    sh.getRange(1, 1).setValue('⚠️ Нет данных. Запустите syncTinkoffPositions() сначала.');
    return;
  }

  let totalRub = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);
  let tz       = Session.getScriptTimeZone();
  let now      = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
  let COLS     = 6;
  let r        = 1;

  // ── Шапка ─────────────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS,
    '📊  ДАШБОРД ПОРТФЕЛЯ',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });
  r++;

  mergedCell_(sh, r, 1, 1, COLS,
    'Обновлено: ' + now + '   ·   Общий портфель: ' + rub_(totalRub),
    { bg: '#263238', fg: '#b0bec5', align: 'center' });
  r += 2;

  // ── Блок: по классам ──────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '▌ РАСПРЕДЕЛЕНИЕ ПО КЛАССАМ',
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;

  hdrRow_(sh, r,
    ['Категория', 'Сумма, ₽', 'Текущий %', 'Цель %', 'Отклонение', 'Статус'],
    COLS);
  r++;

  let cats = Object.keys(config.classTargets);
  cats.forEach(function(cat, idx) {
    let actual  = positions.filter(function(p) { return p.category === cat; })
                           .reduce(function(s, p) { return s + p.valueRub; }, 0);
    let actPct  = totalRub > 0 ? actual / totalRub : 0;
    let tgtPct  = config.classTargets[cat] || 0;
    let diff    = actPct - tgtPct;
    let absDiff = Math.abs(diff) * 100;
    let txt     = absDiff <= THR.OK   ? '✅ Норма'
                : absDiff <= THR.WARN ? '⚠️ Умеренно'
                : '🔴 Требует внимания';
    let clr     = absDiff <= THR.OK ? C.OK : absDiff <= THR.WARN ? C.WARN : C.CRIT;
    let bg      = idx % 2 === 0 ? C.EVEN : C.ODD;

    sh.getRange(r, 1, 1, COLS)
      .setValues([[cat, actual, actPct, tgtPct, diff, txt]])
      .setBackground(bg);
    sh.getRange(r, 2).setNumberFormat('#,##0 [$₽-ru-RU]');
    sh.getRange(r, 3).setNumberFormat('0.0%');
    sh.getRange(r, 4).setNumberFormat('0.0%');
    sh.getRange(r, 5).setNumberFormat('+0.0%;-0.0%;0.0%')
                     .setFontColor(clr).setFontWeight('bold');
    sh.getRange(r, 6).setFontColor(clr).setFontWeight('bold');
    r++;
  });
  r++;

  // ── Блок: акции ───────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '▌ АКЦИИ — ДЕТАЛИЗАЦИЯ',
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;

  hdrRow_(sh, r,
    ['Название', 'Тикер', 'Сумма, ₽', 'Текущий %', 'Цель %', 'Отклонение'],
    COLS);
  r++;

  let shares = positions.filter(function(p) { return p.category === 'Акции'; });
  shares.sort(function(a, b) { return b.valueRub - a.valueRub; });

  shares.forEach(function(p, idx) {
    let actPct  = totalRub > 0 ? p.valueRub / totalRub : 0;
    let tgtPct  = findTarget_(p.name, p.ticker, config.stockTargets);
    let diff    = actPct - tgtPct;
    let absDiff = Math.abs(diff) * 100;
    let clr     = absDiff <= THR.OK ? C.OK : absDiff <= THR.WARN ? C.WARN : C.CRIT;
    let bg      = idx % 2 === 0 ? C.EVEN : C.ODD;

    sh.getRange(r, 1, 1, COLS)
      .setValues([[p.name, p.ticker, p.valueRub, actPct, tgtPct, diff]])
      .setBackground(bg);
    sh.getRange(r, 3).setNumberFormat('#,##0 [$₽-ru-RU]');
    sh.getRange(r, 4).setNumberFormat('0.0%');
    sh.getRange(r, 5).setNumberFormat('0.0%');
    sh.getRange(r, 6).setNumberFormat('+0.0%;-0.0%;0.0%')
                     .setFontColor(clr).setFontWeight('bold');
    r++;
  });

  // ── Ширина колонок ────────────────────────────────────────────────
  [260, 70, 155, 105, 80, 145].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(4);
  addDashboardCharts();
  SpreadsheetApp.flush();
}



// ════════════════════════════════════════════════════════════════════
// 8. ЕДИНАЯ СИНХРОНИЗАЦИЯ (для триггера)
// ════════════════════════════════════════════════════════════════════

/**
 * Запускайте этим триггером раз в час.
 * Rebalance не трогается — пересчитывается только вручную.
 */
function syncAndRefresh() {
  syncTinkoffPositions();  // из tinvest.js
  updateDashboard();
  hideDataSheets(false);  // скрываем технические листы (без всплывающего окна)
}


// ════════════════════════════════════════════════════════════════════
// 9. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ════════════════════════════════════════════════════════════════════

/**
 * Читает из листа Rebalance список имён, у которых чекбокс (колонка F) = true.
 * Вызывается ДО clearContents(), чтобы не потерять состояние.
 */
function readSkipped_(sh) {
  try {
    let data = sh.getDataRange().getValues();
    let result = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][5] === true) {
        let name = String(data[i][0]).trim();
        // Исключаем строки-заголовки
        if (name && name !== 'Категория' && name !== 'Название') {
          result.push(name);
        }
      }
    }
    return result;
  } catch(e) {
    return [];
  }
}

/**
 * Ищет целевую долю акции по имени (точное → частичное совпадение).
 */
function findTarget_(name, ticker, stockTargets) {
  if (stockTargets[name] !== undefined) return stockTargets[name];
  let nl = name.toLowerCase();
  let keys = Object.keys(stockTargets);
  for (let i = 0; i < keys.length; i++) {
    let k = keys[i].toLowerCase();
    if (nl.includes(k) || k.includes(nl)) return stockTargets[keys[i]];
  }
  return 0;
}

/**
 * Ищет позицию в массиве по имени (частичное совпадение).
 */
function matchPos_(positions, name) {
  let nl = name.toLowerCase();
  for (let i = 0; i < positions.length; i++) {
    let pnl = positions[i].name.toLowerCase();
    if (pnl === nl || pnl.includes(nl) || nl.includes(pnl)) return positions[i];
  }
  return null;
}

/** Форматирование числа в рубли (без Intl, работает везде). */
function rub_(amount) {
  let n = Math.round(amount);
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '\u00a0') + '\u00a0₽';
}

/** Создаёт объединённую ячейку с форматированием. */
function mergedCell_(sh, row, col, rows, cols, value, fmt) {
  let rng = sh.getRange(row, col, rows, cols).merge().setValue(value);
  if (fmt.bg)     rng.setBackground(fmt.bg);
  if (fmt.fg)     rng.setFontColor(fmt.fg);
  if (fmt.bold)   rng.setFontWeight('bold');
  if (fmt.size)   rng.setFontSize(fmt.size);
  if (fmt.align)  rng.setHorizontalAlignment(fmt.align);
  if (fmt.italic) rng.setFontStyle('italic');
}

/** Строка заголовка таблицы. */
function hdrRow_(sh, row, headers, cols) {
  sh.getRange(row, 1, 1, cols).setValues([headers])
    .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold');
}


// ════════════════════════════════════════════════════════════════════
// УМНОЕ РАСПРЕДЕЛЕНИЕ С УЧЁТОМ ЛОТОВ
// ════════════════════════════════════════════════════════════════════

function allocateWithLots_(budget, stockNeed, skipped) {
  let activeNames = Object.keys(stockNeed).filter(function(n) {
    return skipped.indexOf(n) === -1 && stockNeed[n].need > 0;
  });

  // Шаг 1: пропорциональные аллокации по дефициту
  let totalNeed = activeNames.reduce(function(s, n) { return s + stockNeed[n].need; }, 0);
  let allocs = {};
  activeNames.forEach(function(n) {
    allocs[n] = totalNeed > 0 ? (stockNeed[n].need / totalNeed) * budget : 0;
  });

  // Шаг 2: округляем вниз до целых лотов, считаем остаток
  let results = {};
  let remainder = budget;

  Object.keys(stockNeed).forEach(function(n) {
    let info    = stockNeed[n];
    let price   = info.price || 0;
    let lot     = info.lot   || 1;
    let lotCost = price * lot;

    if (skipped.indexOf(n) !== -1) {
      results[n] = { lots: 0, actualAlloc: 0, lotCost: lotCost, unknown: false };
      return;
    }
    if (price <= 0 || lotCost <= 0) {
      // Цена неизвестна (акция ещё не куплена) — показываем пропорциональный бюджет
      let myAllocUnk = allocs[n] || 0;
      results[n] = { lots: '?', actualAlloc: myAllocUnk, lotCost: 0, unknown: true };
      remainder  -= myAllocUnk;  // деньги "зарезервированы" под эту бумагу
      return;
    }

    let myAlloc     = allocs[n] || 0;
    let lots        = Math.floor(myAlloc / lotCost);
    let actualAlloc = lots * lotCost;
    results[n] = { lots: lots, actualAlloc: actualAlloc, lotCost: lotCost, unknown: false };
    remainder  -= actualAlloc;
  });

  // Шаг 3: итеративно докидываем остаток самым отстающим
  let maxIter = 20;
  while (remainder > 0.01 && maxIter > 0) {
    maxIter--;

    let candidates = Object.keys(stockNeed).filter(function(n) {
      let r = results[n];
      return r && !r.unknown && r.lotCost > 0 &&
             remainder >= r.lotCost &&
             skipped.indexOf(n) === -1 &&
             stockNeed[n].need > 0;
    });

    if (candidates.length === 0) break;

    candidates.sort(function(a, b) {
      return stockNeed[b].need - stockNeed[a].need;
    });

    let top = candidates[0];
    results[top].lots        += 1;
    results[top].actualAlloc += results[top].lotCost;
    remainder                -= results[top].lotCost;
  }

  return { results: results, remainder: Math.max(0, remainder) };
}

// ════════════════════════════════════════════════════════════════════
// 7. КАЛЬКУЛЯТОР ПОПОЛНЕНИЯ
// ════════════════════════════════════════════════════════════════════

function calculateRebalance() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.REBALANCE);
  if (!sh) sh = ss.insertSheet(DST.REBALANCE);

  // ⚠️ Читаем данные ДО очистки листа
  let amount  = 0;
  let skipped = [];
  try {
    amount  = Number(sh.getRange('B2').getValue()) || 0;
    skipped = readSkipped_(sh);
  } catch(e) {}

  sh.clearContents();
  sh.clearFormats();

  let config;
  try { config = readConfig_(); }
  catch(e) { sh.getRange(1, 1).setValue('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);
  let COLS      = 7;
  let r         = 1;

  // ── Шапка ─────────────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '💰  КАЛЬКУЛЯТОР ПОПОЛНЕНИЯ',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });
  r++;

  sh.getRange(r, 1).setValue('Сумма пополнения, ₽:').setFontWeight('bold');
  sh.getRange(r, 2)
    .setValue(amount || '')
    .setBackground(C.INPUT).setFontWeight('bold').setFontSize(12)
    .setNumberFormat('#,##0');
  sh.getRange(r, 3)
    .setValue('← введите сумму, затем Tinkoff → Пересчитать')
    .setFontColor('#9e9e9e').setFontStyle('italic');
  r += 2;

  if (amount <= 0) {
    mergedCell_(sh, r, 1, 1, COLS,
      '⬆️  Введите сумму пополнения в жёлтую ячейку B2 и нажмите «Пересчитать» в меню Tinkoff.',
      { fg: C.WARN, bold: true, align: 'center' });
    colWidths_(sh);
    return;
  }

  let newTotal = totalRub + amount;

  // ── Блок: классы ──────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '▌ РАСПРЕДЕЛЕНИЕ ПО КЛАССАМ',
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;
  hdrRow_(sh, r,
    ['Категория', 'Текущий %', 'Цель %', 'Отклонение', 'Рекомендуется, ₽', '⬛ Пропустить', ''],
    COLS);
  r++;

  let cats      = Object.keys(config.classTargets);
  let classNeed = {};
  cats.forEach(function(cat) {
    let actual = positions.filter(function(p) { return p.category === cat; })
                          .reduce(function(s, p) { return s + p.valueRub; }, 0);
    let need   = Math.max(0, newTotal * (config.classTargets[cat] || 0) - actual);
    classNeed[cat] = { need: need, actual: actual };
  });

  let activeCats   = cats.filter(function(c) { return skipped.indexOf(c) === -1; });
  let totalNeedCls = activeCats.reduce(function(s, c) { return s + classNeed[c].need; }, 0);
  let classAlloc   = {};
  cats.forEach(function(cat) {
    if (skipped.indexOf(cat) !== -1 || totalNeedCls === 0) {
      classAlloc[cat] = 0;
    } else {
      classAlloc[cat] = (classNeed[cat].need / totalNeedCls) * amount;
    }
  });

  cats.forEach(function(cat, idx) {
    let actual  = classNeed[cat].actual;
    let actPct  = totalRub > 0 ? actual / totalRub : 0;
    let tgtPct  = config.classTargets[cat] || 0;
    let diff    = actPct - tgtPct;
    let alloc   = classAlloc[cat];
    let isSkip  = skipped.indexOf(cat) !== -1;
    let bg      = idx % 2 === 0 ? C.EVEN : C.ODD;

    sh.getRange(r, 1, 1, 5)
      .setValues([[cat, actPct, tgtPct, diff, alloc]])
      .setBackground(bg);
    sh.getRange(r, 2).setNumberFormat('0.0%');
    sh.getRange(r, 3).setNumberFormat('0.0%');
    sh.getRange(r, 4).setNumberFormat('+0.0%;-0.0%;0.0%');
    sh.getRange(r, 5).setNumberFormat('#,##0 [$₽-ru-RU]');
    if (isSkip) sh.getRange(r, 1, 1, 5).setFontColor(C.SKIP);
    sh.getRange(r, 6).insertCheckboxes().setValue(isSkip);
    r++;
  });
  r++;

  // ── Блок: акции с умным распределением лотов ──────────────────────
  let stockBudget = classAlloc['Акции'] || 0;
  let sharePos    = positions.filter(function(p) { return p.category === 'Акции'; });

  let stockNeed = {};
  Object.keys(config.stockTargets).forEach(function(name) {
    let tgt    = config.stockTargets[name];
    let pos    = matchPos_(sharePos, name);
    let actual = pos ? pos.valueRub : 0;
    let need   = Math.max(0, newTotal * tgt - actual);
    let price  = pos ? pos.price : 0;
    let lot    = pos ? pos.lot   : 1;
    stockNeed[name] = { need: need, pos: pos, actual: actual, tgt: tgt, price: price, lot: lot };
  });

  let lotResult = allocateWithLots_(stockBudget, stockNeed, skipped);

  mergedCell_(sh, r, 1, 1, COLS,
    '▌ АКЦИИ — ЧТО ПОКУПАТЬ   (бюджет: ' + rub_(stockBudget) +
    (lotResult.remainder > 0 ? '   |   остаток: ' + rub_(lotResult.remainder) : '') + ')',
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;
  hdrRow_(sh, r,
    ['Название', 'Тикер', 'Текущий %', 'Цель %', 'Купить на, ₽', '⬛ Пропустить', '≈ Лотов'],
    COLS);
  r++;

  let stockOrder = Object.keys(stockNeed).sort(function(a, b) {
    return stockNeed[b].need - stockNeed[a].need;
  });

  stockOrder.forEach(function(name, idx) {
    let info   = stockNeed[name];
    let isSkip = skipped.indexOf(name) !== -1;
    let res    = lotResult.results[name] || { lots: 0, actualAlloc: 0, lotCost: 0 };
    let ticker = info.pos ? info.pos.ticker : '';
    let actPct = totalRub > 0 ? info.actual / totalRub : 0;
    let lots   = isSkip ? 0 : res.lots;
    let alloc  = isSkip ? 0 : res.actualAlloc;

    let bg = idx % 2 === 0 ? C.EVEN : C.ODD;
    sh.getRange(r, 1, 1, 5)
      .setValues([[name, ticker, actPct, info.tgt, alloc]])
      .setBackground(bg);
    sh.getRange(r, 3).setNumberFormat('0.0%');
    sh.getRange(r, 4).setNumberFormat('0.0%');
    sh.getRange(r, 5).setNumberFormat('#,##0 [$₽-ru-RU]');
    sh.getRange(r, 6).insertCheckboxes().setValue(isSkip);

    let lotsCell = sh.getRange(r, 7);
    if (isSkip) {
      lotsCell.setValue('—').setFontColor(C.SKIP);
    } else if (res.unknown) {
      lotsCell.setValue('? уточни лот').setFontColor(C.WARN).setFontStyle('italic');
    } else if (lots === 0 && info.need > 0) {
      lotsCell.setValue('0 ⚠️').setFontColor(C.CRIT).setFontWeight('bold');
    } else {
      lotsCell.setValue(lots).setFontColor(lots > 0 ? C.OK : '#666666');
    }

    if (isSkip) sh.getRange(r, 1, 1, 7).setFontColor(C.SKIP);
    r++;
  });

  // ── Блок: другие категории ───────────────────────────────────────
  let otherCats   = ['Золото', 'Замещайки', 'Денежный рынок'];
  let otherBudget = otherCats.reduce(function(s,c){ return s + (classAlloc[c]||0); }, 0);

  mergedCell_(sh, r, 1, 1, COLS,
    '▌ ДРУГИЕ КАТЕГОРИИ   (бюджет: ' + rub_(otherBudget) + ')',
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;
  hdrRow_(sh, r,
    ['Категория', 'Инструмент', 'Текущий %', 'Цель %', 'Рекомендуется, \u20bd', '\u2611 Пропустить', 'Комментарий'],
    COLS);
  r++;

  otherCats.forEach(function(cat, idx) {
    let alloc    = classAlloc[cat] || 0;
    let actual   = positions.filter(function(p){ return p.category === cat; })
                            .reduce(function(s,p){ return s + p.valueRub; }, 0);
    let actPct   = totalRub > 0 ? actual / totalRub : 0;
    let tgtPct   = config.classTargets[cat] || 0;
    let isSkip   = skipped.indexOf(cat) !== -1;
    let catPos   = positions.filter(function(p){ return p.category === cat; });
    // Берём инструмент с наибольшей ценой (фонд, а не кэш/дешёвый ETF)
    catPos.sort(function(a, b) { return (b.price || 0) - (a.price || 0); });
    let instrName  = catPos.length > 0 ? catPos[0].name : '—';
    let instrPrice = catPos.length > 0 ? catPos[0].price : 0;
    let instrLot   = catPos.length > 0 ? (catPos[0].lot || 1) : 1;
    let showAlloc  = isSkip ? 0 : alloc;
    let comment = '';

    if (isSkip) {
      comment = 'Пропущено';
    } else if (alloc <= 0) {
      comment = 'Категория выше цели';
    } else if (instrPrice > 0) {
      let lotCost = instrPrice * instrLot;
      let u = Math.floor(alloc / lotCost);
      comment = u > 0
        ? 'Купить ' + u + ' пай(ёв) · ' + instrName
        : 'Не хватает на 1 пай (~' + rub_(lotCost) + ')';
    } else if (cat === 'Замещайки') {
      comment = 'Рассмотреть замещающие облигации на ' + rub_(alloc);
    } else {
      comment = 'Укажите инструмент в Config → Блок 3';
    }

    let bg = idx % 2 === 0 ? C.EVEN : C.ODD;
    sh.getRange(r, 1, 1, COLS)
      .setValues([[cat, instrName, actPct, tgtPct, showAlloc, isSkip, comment]])
      .setBackground(bg);
    sh.getRange(r, 3).setNumberFormat('0.0%');
    sh.getRange(r, 4).setNumberFormat('0.0%');
    sh.getRange(r, 5).setNumberFormat('#,##0 [$\u20bd-ru-RU]');
    sh.getRange(r, 6).insertCheckboxes().setValue(isSkip);
    if (alloc <= 0 && !isSkip) {
      sh.getRange(r, 5).setFontColor(C.SKIP);
      sh.getRange(r, 7).setFontColor(C.SKIP).setFontStyle('italic');
    }
    r++;
  });

  // Облигации — предупреждение "не докупать"
  let bndActual = positions.filter(function(p){ return p.category === 'Облигации'; })
                           .reduce(function(s,p){ return s + p.valueRub; }, 0);
  let bndActPct = totalRub > 0 ? bndActual / totalRub : 0;
  let bndTgt    = config.classTargets['Облигации'] || 0;
  let overPct   = Math.round((bndActPct - bndTgt) * 100);
  sh.getRange(r, 1, 1, COLS)
    .setValues([['Облигации', '—', bndActPct, bndTgt, 0, false,
                 '\u26d4 Выше цели на +' + overPct + ' пп — НЕ ДОКУПАТЬ']])
    .setBackground('#fff3e0');
  sh.getRange(r, 3).setNumberFormat('0.0%');
  sh.getRange(r, 4).setNumberFormat('0.0%');
  sh.getRange(r, 5).setNumberFormat('#,##0 [$\u20bd-ru-RU]');
  sh.getRange(r, 7).setFontColor(C.CRIT).setFontWeight('bold');
  r++;

  colWidths_(sh);
  SpreadsheetApp.flush();
}


/** Стандартная ширина колонок для Rebalance. */
function colWidths_(sh) {
  [250, 80, 110, 85, 175, 125, 95].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });
}
function cleanupOldSheets() {
  let ss = SpreadsheetApp.getActive();
  let toDelete = [
    'Positions', 'Positions_Aggregated', 'Positions_SummaryByType',
    'Positions_Shares', 'Positions_Bonds', 'Positions_ETFs',
    'Positions_Currencies', 'Positions_Futures', 'Positions_Other',
    'Positions_Money', 'Лист1', 'Лист2'
  ];
  toDelete.forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
  SpreadsheetApp.getUi().alert('✅ Старые листы удалены!');
}

function hideDataSheets(showAlert) {
  let ss = SpreadsheetApp.getActive();
  const keepVisible = [DST.CONFIG, DST.DASHBOARD, DST.REBALANCE, DST.INCOME, DST.CALENDAR, DST.HISTORY];
  ss.getSheets().forEach(function(sh) {
    let name = sh.getName();
    if (keepVisible.indexOf(name) === -1) {
      sh.hideSheet();
    }
  });
  if (showAlert !== false) SpreadsheetApp.getUi().alert('✅ Технические листы скрыты.');
}