/**
 * ╔═══════════════════════════════════════════════════════════════════╗
 * ║  dashboard.js — Дашборд портфеля                                 ║
 * ╚═══════════════════════════════════════════════════════════════════╝
 */

// ════════════════════════════════════════════════════════════════════
// 1. КОНСТАНТЫ
// ════════════════════════════════════════════════════════════════════

const SRC = {
  SHARES:  '_Дан_Акции',
  BONDS:   '_Дан_Облигации',
  ETFS:    '_Дан_Фонды',
  MONEY:   '_Дан_Деньги',
};

const DST = {
  CONFIG:    'Настройки',
  DASHBOARD: 'Дашборд',
  REBALANCE: 'Ребалансировка',
  INCOME:    'Ожидаемый доход',
  CALENDAR:  'Календарь выплат',
  HISTORY:   'История операций',
  POSITIONS: '_Позиции',
};

const C = {
  DARK:     '#1a237e',
  MID:      '#283593',
  OK:       '#1b5e20',
  WARN:     '#e65100',
  CRIT:     '#b71c1c',
  ODD:      '#f5f5f5',
  EVEN:     '#ffffff',
  SKIP:     '#9e9e9e',
  INPUT:    '#fff9c4',
};

const THR = { OK: 1.5, WARN: 3.0 }; // оставлено для совместимости, больше не используется напрямую
const THR_525 = { critAbs: 5.0, critRel: 0.25, warnAbs: 2.5, warnRel: 0.125 };

function deviationStatus_(actPct, tgtPct) {
  let params = readAdvancedParams_();
  let absDiffPP = Math.abs(actPct - tgtPct) * 100;
  let relDiff   = tgtPct > 0 ? Math.abs(actPct - tgtPct) / tgtPct : 0;

  let isCrit = absDiffPP >= params.thrAbs || relDiff >= params.thrRel;
  let isWarn = absDiffPP >= params.thrAbs / 2 || relDiff >= params.thrRel / 2;

  if (isCrit) return { clr: C.CRIT, txt: '🔴 Требует внимания' };
  if (isWarn) return { clr: C.WARN, txt: '⚠️ Умеренно' };
  return { clr: C.OK, txt: '✅ Норма' };
}


// ════════════════════════════════════════════════════════════════════
// 2. МЕНЮ
// ════════════════════════════════════════════════════════════════════

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('Tinkoff');

  menu.addItem('🔄  Синхронизировать позиции', 'syncTinkoffPositions')
      .addItem('🚀  Синхронизировать + обновить всё', 'syncAndRefresh')
      .addSeparator();

  menu.addSubMenu(ui.createMenu('📊 Дашборд и ребаланс')
        .addItem('Обновить Dashboard', 'updateDashboard')
        .addItem('Рассчитать доходность (XIRR + IMOEX)', 'calculateAnalytics')
        .addItem('Пересчитать калькулятор пополнения', 'calculateRebalance')
        .addItem('Прогресс к цели', 'calculateGoalProgress'))
      .addSubMenu(ui.createMenu('💰 Доход и история')
        .addItem('Обновить Ожидаемый доход', 'updateIncomeSheet')
        .addItem('Обновить Календарь выплат', 'updateCalendarSheet')
        .addItem('Обновить Историю операций', 'updateHistorySheet'))
      .addSubMenu(ui.createMenu('🎯 Аналитика (вручную)')
        .addItem('Средняя цена и P/L', 'calculateAveragePriceAndPL')
        .addItem('Yield on Cost', 'calculateYieldOnCost')
        .addItem('Health check концентрации', 'calculateConcentrationHealth')
        .addItem('ИИС-3 — вычет за год', 'calculateIisDeductionUsage')
        .addItem('Льгота на долгосрочное владение (ЛДВ)', 'calculateLdvEligibility')
        .addItem('Дисциплина пополнений', 'calculateContributionDiscipline'))
      .addSubMenu(ui.createMenu('📄 Отчёты')
        .addItem('Сформировать годовой отчёт', 'generateAnnualReport')
        .addItem('Открыть HTML-дашборд', 'showHtmlDashboard'))
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ Настройки')
        .addItem('Инициализировать Config', 'initConfig')
        .addItem('Добавить блок дивидендов', 'addDividendsBlock')
        .addItem('Добавить блок продвинутых параметров', 'addAdvancedParamsBlock')
        .addItem('Добавить блок ИИС-3', 'addIisBlock')
        .addItem('Добавить блок цели портфеля', 'addGoalBlock')
        .addItem('Проверить подключение Telegram', 'testTelegramConnection'));
        

  menu.addToUi();
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

  let rows = [
    ['▌ ЦЕЛЕВАЯ СТРУКТУРА ПОРТФЕЛЯ (% от всего портфеля)', '', ''],
    ['Категория',        'Цель %', 'Счёт'],
    ['Акции',            '',       ''],
    ['Облигации',        '',       ''],
    ['Золото',           '',       ''],
    ['Замещайки',        '',       ''],
    ['Денежный рынок',   '',       ''],
    ['', '', ''],

    ['▌ ЦЕЛЕВЫЕ ДОЛИ АКЦИЙ (% от всего портфеля)', '', ''],
    ['Название (точно как в разделе «Акции — детализация» на Дашборде)',  'Цель %', 'Тикер'],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],

    ['▌ МАППИНГ СПЕЦКАТЕГОРИЙ — заполните точные названия из раздела «Акции — детализация» на Дашборде', '', ''],
    ['Название инструмента',                    'Категория',         'Комментарий'],
    ['',                                         'Золото',            '← Укажите, где храните золото (например, фонд ВИМ или другой)'],
    ['',                                         'Замещайки',         '← Укажите замещающую облигацию'],
    ['',                                         'Денежный рынок',    '← Укажите фонд денежного рынка'],
  ];

  sh.getRange(1, 1, rows.length, 3).setValues(rows);

  [[1, C.DARK], [9, C.DARK], [23, C.DARK]].forEach(function(pair) {
    sh.getRange(pair[0], 1, 1, 3).merge()
      .setBackground(pair[1]).setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(11);
  });

  [[2, C.MID], [10, C.MID], [24, C.MID]].forEach(function(pair) {
    sh.getRange(pair[0], 1, 1, 3)
      .setBackground(pair[1]).setFontColor('#ffffff').setFontWeight('bold');
  });

  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 260);
  sh.getRange(1, 3, sh.getLastRow(), 1).setWrap(true);

  if (typeof addDividendsBlock === 'function') {
    addDividendsBlock();
  }
  if (typeof addAdvancedParamsBlock === 'function') addAdvancedParamsBlock();
  if (typeof addGoalBlock === 'function') addGoalBlock();
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

  let classTargets = {};
  for (let i = 2; i <= 6 && i < v.length; i++) {
    let cat = String(v[i][0]).trim();
    let pct = Number(v[i][1]);
    if (cat && !isNaN(pct) && pct > 0) classTargets[cat] = pct / 100;
  }

  let stockTargets = {};
  for (let j = 10; j <= 21 && j < v.length; j++) {
    let sName = String(v[j][0]).trim();
    let sPct  = Number(v[j][1]);
    if (sName && !isNaN(sPct) && sPct > 0) stockTargets[sName] = sPct / 100;
  }

  let mapping = {};
  for (let k = 24; k < v.length; k++) {
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

  mergedCell_(sh, r, 1, 1, COLS,
    '📊  ДАШБОРД ПОРТФЕЛЯ',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });
  r++;

  mergedCell_(sh, r, 1, 1, COLS,
    'Обновлено: ' + now + '   ·   Общий портфель: ' + rub_(totalRub),
    { bg: '#263238', fg: '#b0bec5', align: 'center' });
  r += 2;

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
    let status  = deviationStatus_(actPct, tgtPct);
    let bg      = idx % 2 === 0 ? C.EVEN : C.ODD;

    sh.getRange(r, 1, 1, COLS)
      .setValues([[cat, actual, actPct, tgtPct, diff, status.txt]])
      .setBackground(bg);
    sh.getRange(r, 2).setNumberFormat('#,##0 [$₽-ru-RU]');
    sh.getRange(r, 3).setNumberFormat('0.0%');
    sh.getRange(r, 4).setNumberFormat('0.0%');
    sh.getRange(r, 5).setNumberFormat('+0.0%;-0.0%;0.0%')
                     .setFontColor(status.clr).setFontWeight('bold');
    sh.getRange(r, 6).setFontColor(status.clr).setFontWeight('bold');
    r++;
  });
  r++;

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
    let status  = deviationStatus_(actPct, tgtPct);
    let bg      = idx % 2 === 0 ? C.EVEN : C.ODD;

    sh.getRange(r, 1, 1, COLS)
      .setValues([[p.name, p.ticker, p.valueRub, actPct, tgtPct, diff]])
      .setBackground(bg);
    sh.getRange(r, 3).setNumberFormat('#,##0 [$₽-ru-RU]');
    sh.getRange(r, 4).setNumberFormat('0.0%');
    sh.getRange(r, 5).setNumberFormat('0.0%');
    sh.getRange(r, 6).setNumberFormat('+0.0%;-0.0%;0.0%')
                     .setFontColor(status.clr).setFontWeight('bold');
    r++;
  });

  [260, 70, 155, 105, 80, 145].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(4);
  addDashboardCharts();
  calculateAnalytics();
  calculateContributionDiscipline();
  calculateConcentrationHealth();
  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════════════
// 8. ЕДИНАЯ СИНХРОНИЗАЦИЯ (для триггера)
// ════════════════════════════════════════════════════════════════════

function syncAndRefresh() {
  syncTinkoffPositions();
  updateDashboard();
  updateIncomeSheet();
  updateCalendarSheet();
  hideDataSheets(false);
  checkAndNotifyDeviations_();
  notifySyncComplete_();
}


// ════════════════════════════════════════════════════════════════════
// 9. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ════════════════════════════════════════════════════════════════════

function readSkipped_(sh) {
  try {
    let data = sh.getDataRange().getValues();
    let result = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][5] === true) {
        let name = String(data[i][0]).trim();
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

function matchPos_(positions, name) {
  let nl = name.toLowerCase();
  for (let i = 0; i < positions.length; i++) {
    let pnl = positions[i].name.toLowerCase();
    if (pnl === nl || pnl.includes(nl) || nl.includes(pnl)) return positions[i];
  }
  return null;
}

function rub_(amount) {
  let n = Math.round(amount);
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '\u00a0') + '\u00a0₽';
}

function mergedCell_(sh, row, col, rows, cols, value, fmt) {
  let rng = sh.getRange(row, col, rows, cols).merge().setValue(value);
  if (fmt.bg)     rng.setBackground(fmt.bg);
  if (fmt.fg)     rng.setFontColor(fmt.fg);
  if (fmt.bold)   rng.setFontWeight('bold');
  if (fmt.size)   rng.setFontSize(fmt.size);
  if (fmt.align)  rng.setHorizontalAlignment(fmt.align);
  if (fmt.italic) rng.setFontStyle('italic');
}

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

  let totalNeed = activeNames.reduce(function(s, n) { return s + stockNeed[n].need; }, 0);
  let allocs = {};
  activeNames.forEach(function(n) {
    allocs[n] = totalNeed > 0 ? (stockNeed[n].need / totalNeed) * budget : 0;
  });

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
      let myAllocUnk = allocs[n] || 0;
      results[n] = { lots: '?', actualAlloc: myAllocUnk, lotCost: 0, unknown: true };
      remainder  -= myAllocUnk;
      return;
    }

    let myAlloc     = allocs[n] || 0;
    let lots        = Math.floor(myAlloc / lotCost);
    let actualAlloc = lots * lotCost;
    results[n] = { lots: lots, actualAlloc: actualAlloc, lotCost: lotCost, unknown: false };
    remainder  -= actualAlloc;
  });

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

  mergedCell_(sh, r, 1, 1, COLS, '▌ РАСПРЕДЕЛЕНИЕ ПО КЛАССАМ',
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;
  hdrRow_(sh, r,
    ['Категория', 'Текущий %', 'Цель %', 'Отклонение', 'Рекомендуется, ₽', '⬛ Пропустить', 'Статус'],
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
    let status  = deviationStatus_(actPct, tgtPct);
    let bg      = idx % 2 === 0 ? C.EVEN : C.ODD;

    sh.getRange(r, 1, 1, 5)
      .setValues([[cat, actPct, tgtPct, diff, alloc]])
      .setBackground(bg);
    sh.getRange(r, 2).setNumberFormat('0.0%');
    sh.getRange(r, 3).setNumberFormat('0.0%');
    sh.getRange(r, 4).setNumberFormat('+0.0%;-0.0%;0.0%')
                     .setFontColor(status.clr).setFontWeight('bold');
    sh.getRange(r, 5).setNumberFormat('#,##0 [$₽-ru-RU]');
    if (isSkip) sh.getRange(r, 1, 1, 5).setFontColor(C.SKIP);
    sh.getRange(r, 6).insertCheckboxes().setValue(isSkip);
    sh.getRange(r, 7).setValue(status.txt).setFontColor(status.clr).setFontWeight('bold');
    r++;
  });
  r++;

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

  sh.getRange(1, 7, sh.getLastRow(), 1).setWrap(true); // колонка «Комментарий»
  colWidths_(sh);
}


function colWidths_(sh) {
  [250, 80, 110, 85, 175, 125, 260].forEach(function(w, i) {
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