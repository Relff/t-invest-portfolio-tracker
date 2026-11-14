/**

 * ╔═══════════════════════════════════════════════════════════════════╗

 * ║  dashboard.js — Дашборд портфеля                                 ║

 * ╚═══════════════════════════════════════════════════════════════════╝

 */



// ════════════════════════════════════════════════════════════════════

// 1. КОНСТАНТЫ

// ════════════════════════════════════════════════════════════════════



const SRC = {

  SHARES:     '_Дан_Акции',

  BONDS:      '_Дан_Облигации',

  ETFS:       '_Дан_Фонды',

  CURRENCIES: '_Дан_Валюта',

  MONEY:      '_Дан_Деньги',

};



const DST = {

  CONFIG:    '⚙️ Настройки',

  DASHBOARD: '📊 Дашборд',

  REBALANCE: '💰 Ребалансировка',

  INCOME:    '💵 Ожидаемый доход',

  CALENDAR:  '📅 Календарь выплат',

  HISTORY:   '📜 История операций',

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



/**

 * Разовая миграция: переименовывает уже существующие вкладки листов под

 * новые имена с эмодзи (см. DST, CORR_SHEET, REBAL_HISTORY_SHEET).

 * Без нужды не запускать больше одного раза — но и повторный запуск

 * безопасен: если вкладка уже переименована, просто пропускается.

 */

function renameSheetTabsWithEmoji() {

  let ss = SpreadsheetApp.getActive();

  let renames = [

    ['Настройки', DST.CONFIG],

    ['Дашборд', DST.DASHBOARD],

    ['Ребалансировка', DST.REBALANCE],

    ['Ожидаемый доход', DST.INCOME],

    ['Календарь выплат', DST.CALENDAR],

    ['История операций', DST.HISTORY],

    ['Корреляция акций', CORR_SHEET],

    ['История ребалансировок', REBAL_HISTORY_SHEET],

  ];

  let renamed = [];

  renames.forEach(function(pair) {

    let oldName = pair[0], newName = pair[1];

    if (oldName === newName) return;

    let sh = ss.getSheetByName(oldName);

    if (sh) { sh.setName(newName); renamed.push(newName); }

  });

  SpreadsheetApp.getUi().alert(renamed.length

    ? '✅ Переименовано вкладок: ' + renamed.length + '\n\n' + renamed.join('\n')

    : 'Нечего переименовывать — либо уже переименовано, либо листы ещё не созданы.');

}



/**

 * Одна кнопка вместо десяти: пересчитывает всё содержимое листа «Дашборд»

 * разом (P/L, Yield on Cost, Health check, ЛДВ, Дисциплина, XIRR+IMOEX

 * + опциональные блоки). Опциональные блоки (ИИС-3, сектора, цель

 * портфеля) тихо пропускаются, если ещё не настроены в Config — иначе

 * при каждом обычном обновлении заваливало бы тремя алертами

 * «сначала настрой», что убивало бы весь смысл кнопки «в один клик».

 */

function updateDashboardAll() {

  if (typeof updateCashSummary_ === 'function') updateCashSummary_();

  updateDashboard(); // уже сама пересчитывает Аналитику, Дисциплину и Health check внутри себя

  calculateAveragePriceAndPL();

  calculateYieldOnCost();

  calculateLdvEligibility();



  if (readIisAccountName_())               calculateIisDeductionUsage();

  if (Object.keys(readSectorMap_()).length) calculateSectorDiversification();

  if (readGoalTarget_() > 0)                calculateGoalProgress();



  SpreadsheetApp.getUi().alert('✅ Дашборд полностью обновлён.');

}



function onOpen() {

  let ui = SpreadsheetApp.getUi();

  let menu = ui.createMenu('Tinkoff');



  menu.addItem('🔄  Синхронизировать позиции', 'syncTinkoffPositions')

      .addItem('🚀  Синхронизировать + обновить всё', 'syncAndRefresh')

      .addSeparator()

      .addItem('🔄  Обновить Dashboard (всё)', 'updateDashboardAll')
      .addItem('💵  Обновить сводку наличных', 'updateCashSummary_')

      .addSeparator();



  menu.addSubMenu(ui.createMenu('📊 Ребаланс и отчёты')

        .addItem('💰 Пересчитать калькулятор пополнения', 'calculateRebalance')

        .addItem('📋 История ребалансировок: сравнить с фактом', 'compareRebalanceExecution')

        .addItem('📄 Сформировать годовой отчёт', 'generateAnnualReport')

        .addItem('📈 Открыть HTML-дашборд', 'showHtmlDashboard'))

      .addSubMenu(ui.createMenu('💰 Доход и история')

        .addItem('💵 Обновить Ожидаемый доход', 'updateIncomeSheet')

        .addItem('📅 Обновить Календарь выплат', 'updateCalendarSheet')

        .addItem('📜 Обновить Историю операций', 'updateHistorySheet'))

      .addSubMenu(ui.createMenu('🧮 Тяжёлая аналитика (не для частого использования)')

        .addItem('🔗 Корреляция между акциями', 'calculateStockCorrelation')

        .addItem('🥇 Доп. бенчмарки (RGBI, золото)', 'calculateExtraBenchmarks'))

      .addSubMenu(ui.createMenu('⚙️ Настройки (Config)')

        .addItem('🆕 Инициализировать Config', 'initConfig')

        .addItem('💵 Добавить блок дивидендов', 'addDividendsBlock')

        .addItem('🎛️ Добавить блок продвинутых параметров', 'addAdvancedParamsBlock')

        .addItem('🏦 Добавить блок ИИС-3', 'addIisBlock')

        .addItem('🎯 Добавить блок цели портфеля', 'addGoalBlock')

        .addItem('🏭 Добавить блок секторов акций', 'addSectorsBlock'))

      .addSubMenu(ui.createMenu('🤖 Telegram')

        .addItem('📡 Проверить подключение', 'testTelegramConnection'))

      .addItem('ℹ️ О трекере', 'showAboutTracker');



  menu.addToUi();

}



function showAboutTracker() {

  SpreadsheetApp.getUi().alert(aboutTrackerText_());

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

    // Старые объединения переживают clearContents/clearFormats и могут
    // конфликтовать с новой структурой Config. Поэтому сначала снимаем их.
    if (sh.getMaxRows() > 0 && sh.getMaxColumns() > 0) {
      sh.getRange(1, 1, sh.getMaxRows(), Math.min(3, sh.getMaxColumns())).breakApart();
    }
    sh.clearContents();
    sh.clearFormats();
  } else {
    sh = ss.insertSheet(DST.CONFIG);
  }

  let rows = [
    ['▌ ЦЕЛЕВАЯ СТРУКТУРА ПОРТФЕЛЯ (% от всего портфеля)', '', ''],
    ['Категория', 'Цель %', 'Счёт'],
    ['Акции', '', ''],
    ['Облигации', '', ''],
    ['Золото', '', ''],
    ['Замещайки', '', ''],
    ['Денежный рынок', '', ''],
    ['Кэш', '', ''],
    ['', '', ''],

    ['▌ ЦЕЛЕВЫЕ ДОЛИ АКЦИЙ (% от всего портфеля)', '', ''],
    ['Название (точно как в разделе «Акции — детализация» на Дашборде)', 'Цель %', 'Тикер'],
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

    ['▌ МАППИНГ СПЕЦКАТЕГОРИЙ — заполните точные названия инструментов', '', ''],
    ['Название инструмента', 'Категория', 'Комментарий'],
    ['', 'Золото', '← Укажите инструмент, который нужно считать золотом'],
    ['', 'Замещайки', '← Укажите замещающую облигацию'],
    ['', 'Денежный рынок', '← Укажите фонд денежного рынка'],
  ];

  sh.getRange(1, 1, rows.length, 3).setValues(rows);

  // Оформление больше не зависит от конкретных номеров строк.
  // Можно добавлять/убирать категории и строки акций — стили не съедут.
  rows.forEach(function(row, idx) {
    let rowNumber = idx + 1;
    let title = String(row[0] || '').trim();

    if (title.indexOf('▌') === 0) {
      sh.getRange(rowNumber, 1, 1, 3).merge()
        .setBackground(C.DARK)
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontSize(11);

      // Строка сразу после заголовка секции — шапка таблицы.
      if (rowNumber < rows.length) {
        sh.getRange(rowNumber + 1, 1, 1, 3)
          .setBackground(C.MID)
          .setFontColor('#ffffff')
          .setFontWeight('bold');
      }
    }
  });

  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 300);
  sh.getRange(1, 3, sh.getLastRow(), 1).setWrap(true);

  if (typeof addDividendsBlock === 'function') addDividendsBlock();
  if (typeof addAdvancedParamsBlock === 'function') addAdvancedParamsBlock();
  if (typeof addGoalBlock === 'function') addGoalBlock();

  SpreadsheetApp.getUi().alert(
    '✅ Config создан!\n\n' +
    'Проценты можно вводить как 40% или как число 40 — readConfig_() понимает оба варианта.\n' +
    'Названия секций и количество строк теперь можно менять без привязки к номерам строк.'
  );
}





// ════════════════════════════════════════════════════════════════════

// 4. ЧТЕНИЕ CONFIG

// ════════════════════════════════════════════════════════════════════



function normalizeTicker_(ticker) {
  return String(ticker || '')
    .trim()
    .toUpperCase()
    .replace(/@$/, '');
}

function parseTargetPercent_(value) {
  if (value === '' || value === null || value === undefined) return 0;

  // На случай текстового значения вроде "40%" / "40,00%".
  if (typeof value === 'string') {
    let s = value.trim().replace(/\s/g, '').replace(',', '.');
    let hasPercent = s.endsWith('%');
    if (hasPercent) s = s.slice(0, -1);

    let n = Number(s);
    if (!isFinite(n) || n <= 0) return 0;
    return hasPercent ? n / 100 : (n > 1 ? n / 100 : n);
  }

  // Google Sheets хранит 40% как 0.4. Если пользователь ввёл просто 40,
  // превращаем его в 0.4. Это позволяет использовать оба формата.
  let n = Number(value);
  if (!isFinite(n) || n <= 0) return 0;
  return n > 1 ? n / 100 : n;
}

function findConfigSection_(values, titlePart) {
  let needle = String(titlePart || '').trim().toUpperCase();
  for (let i = 0; i < values.length; i++) {
    let title = String(values[i][0] || '').trim().toUpperCase();
    if (title.indexOf(needle) >= 0) return i;
  }
  return -1;
}

function findNextConfigSection_(values, sectionRow) {
  for (let i = sectionRow + 1; i < values.length; i++) {
    let title = String(values[i][0] || '').trim();
    if (title.indexOf('▌') === 0) return i;
  }
  return values.length;
}

function readConfig_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) throw new Error('Лист Config не найден. Запустите initConfig().');

  let v = sh.getDataRange().getValues();

  let classSection = findConfigSection_(v, 'ЦЕЛЕВАЯ СТРУКТУРА ПОРТФЕЛЯ');
  let stockSection = findConfigSection_(v, 'ЦЕЛЕВЫЕ ДОЛИ АКЦИЙ');
  let mapSection   = findConfigSection_(v, 'МАППИНГ СПЕЦКАТЕГОРИЙ');

  if (classSection < 0) throw new Error('В Config не найден блок «ЦЕЛЕВАЯ СТРУКТУРА ПОРТФЕЛЯ».');
  if (stockSection < 0) throw new Error('В Config не найден блок «ЦЕЛЕВЫЕ ДОЛИ АКЦИЙ».');
  if (mapSection < 0)   throw new Error('В Config не найден блок «МАППИНГ СПЕЦКАТЕГОРИЙ».');

  // ── 1. Целевые доли классов ────────────────────────────────────────
  let classTargets = {};
  let classEnd = findNextConfigSection_(v, classSection);

  // +2: пропускаем строку заголовка секции и шапку колонок.
  for (let i = classSection + 2; i < classEnd; i++) {
    let cat = String(v[i][0] || '').trim();
    let pct = parseTargetPercent_(v[i][1]);

    if (cat && pct > 0) {
      classTargets[cat] = pct;
    }
  }

  // ── 2. Целевые доли отдельных акций ───────────────────────────────
  let stockTargets = {};          // название -> доля
  let stockTickers = {};          // название -> тикер
  let stockTargetsByTicker = {};  // нормализованный тикер -> доля
  let stockEnd = findNextConfigSection_(v, stockSection);

  for (let i = stockSection + 2; i < stockEnd; i++) {
    let name   = String(v[i][0] || '').trim();
    let pct    = parseTargetPercent_(v[i][1]);
    let ticker = String(v[i][2] || '').trim();

    if (!name || pct <= 0) continue;

    stockTargets[name] = pct;

    if (ticker) {
      stockTickers[name] = ticker;
      stockTargetsByTicker[normalizeTicker_(ticker)] = pct;
    }
  }

  // ── 3. Маппинг спецкатегорий ──────────────────────────────────────
  let mapping = {};
  let mapEnd = findNextConfigSection_(v, mapSection);

  for (let i = mapSection + 2; i < mapEnd; i++) {
    let name = String(v[i][0] || '').trim();
    let cat  = String(v[i][1] || '').trim();

    if (name && cat) mapping[name] = cat;
  }

  return {
    classTargets: classTargets,
    stockTargets: stockTargets,
    stockTickers: stockTickers,
    stockTargetsByTicker: stockTargetsByTicker,
    mapping: mapping,
  };
}





// ════════════════════════════════════════════════════════════════════

// 5. ЧТЕНИЕ ПОЗИЦИЙ ИЗ ЛИСТОВ TINVEST.JS

// ════════════════════════════════════════════════════════════════════



function readPositions_(config) {

  let ss = SpreadsheetApp.getActive();



  /*

   * Источники автоматически полученных данных.

   *

   * ВАЖНО:

   * - ETF по умолчанию считаем акциями, а не золотом.

   *   Специальные фонды вроде TMON будут переклассифицированы ниже.

   * - Валюты по умолчанию считаем кэшем.

   * - Деньги на брокерском счёте тоже считаем кэшем.

   */

  let sources = [

    { sheet: SRC.SHARES,     defaultCat: 'Акции' },

    { sheet: SRC.BONDS,      defaultCat: 'Облигации' },

    { sheet: SRC.ETFS,       defaultCat: 'Акции' },

    { sheet: SRC.CURRENCIES, defaultCat: 'Кэш' },

    { sheet: SRC.MONEY,      defaultCat: 'Кэш' },

  ];



  let positions = [];



  for (let s = 0; s < sources.length; s++) {

    let src = sources[s];

    let sh = ss.getSheetByName(src.sheet);



    if (!sh) continue;



    let data = sh.getDataRange().getValues();

    if (data.length < 2) continue;



    // Индексы колонок по их названиям

    let H = {};

    for (let hi = 0; hi < data[0].length; hi++) {

      H[String(data[0][hi]).trim()] = hi;

    }



    for (let ri = 1; ri < data.length; ri++) {

      let row = data[ri];



      let name = String(row[H['name']] || '').trim();

      let ticker = String(row[H['ticker']] || '').trim();



      let valueRub = Number(row[H['position_value_rub']] || 0);

      let price = Number(row[H['current_price_rub_per_piece']] || 0);

      let qty = Number(row[H['quantity_pcs']] || 0);

      let lot = Number(row[H['lot']] || 1);



      if (!name || valueRub === 0) continue;



      /*

       * RUB приходит из T-Invest дважды:

       *

       * _Дан_Валюта:

       *   Российский рубль / RUB000UTSTOM

       *

       * _Дан_Деньги:

       *   Деньги (RUB)

       *

       * Оставляем только _Дан_Деньги.

       */

      if (

        src.sheet === SRC.CURRENCIES &&

        (ticker === 'RUB000UTSTOM' || ticker === 'RUB')

      ) {

        continue;

      }



      // Категория по умолчанию

      let category = src.defaultCat;



      /*

       * Встроенные правила для наиболее очевидных инструментов.

       * Благодаря этому они будут работать даже без ручного mapping.

       */

      let normalizedTicker = ticker.replace(/@$/, '').toUpperCase();



      // Биржевое золото Московской биржи

      if (normalizedTicker === 'GLDRUB_TOM') {

        category = 'Золото';

      }



      // TMON — фонд денежного рынка

      if (normalizedTicker === 'TMON') {

        category = 'Денежный рынок';

      }



      /*

       * Пользовательский mapping имеет высший приоритет.

       *

       * Например:

       * "Денежный рынок" -> "Денежный рынок"

       * "Крупнейшие компании РФ" -> "Акции"

       * "Золото" -> "Золото"

       */

      if (config.mapping && config.mapping[name]) {

        category = config.mapping[name];

      }



      positions.push({

        name: name,

        ticker: ticker,

        category: category,

        valueRub: valueRub,

        price: price,

        qty: qty,

        lot: lot,

        source: 'T-Invest',

      });

    }

  }



  // Ручные денежные активы читаются из отдельного модуля cash.js.
  // Это позволяет журналу A:G жить независимо от автоматической сводки I:R.
  if (typeof getManualCashPositions_ === 'function') {
    positions = positions.concat(getManualCashPositions_());
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



  // Версия трекера — пользователь сам вписывает и меняет её в этой ячейке

  // (строка 3, колонка A), не в коде. Читаем ДО очистки листа и запишем

  // обратно после пересборки — иначе ежедневный clearContents() стирал бы

  // ручную правку. Если ячейка ещё пустая (первый запуск) — берём дефолт.

  let versionCell = sh.getRange(3, 1);

  let savedVersion = String(versionCell.getValue() || '').trim();

  let trackerVersion = savedVersion || TRACKER_VERSION_DEFAULT;



  // clearContents()/clearFormats() НЕ снимают объединение ячеек — это

  // отдельное структурное свойство. Без явного breakApart() старое

  // объединение заголовка (ещё на 6 колонок, с прошлой версии) конфликтует

  // с новым на всю ширину (13 колонок) — Sheets API кидает ошибку

  // "необходимо выделить весь диапазон", хотя вручную это разрешено.

  if (sh.getMaxRows() > 0 && sh.getMaxColumns() > 0) {

    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();

  }

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

  let FULL_W   = 13; // левая колонка (1-6) + отступ (7) + правая колонка (8-13) — вся ширина дашборда

  let r        = 1;



  mergedCell_(sh, r, 1, 1, FULL_W,

    '📊  ДАШБОРД ПОРТФЕЛЯ',

    { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });

  r++;



  // Сумма портфеля — самая важная цифра на листе, выделяем её крупнее

  // и белым на фоне остального (более блёклого) текста подзаголовка.

  let prefixText = 'Обновлено: ' + now + '   ·   Общий портфель: ';

  let valueText  = rub_(totalRub);

  let subtitleText = prefixText + valueText;

  let richText = SpreadsheetApp.newRichTextValue()

    .setText(subtitleText)

    .setTextStyle(0, prefixText.length,

      SpreadsheetApp.newTextStyle().setForegroundColor('#b0bec5').setFontSize(10).build())

    .setTextStyle(prefixText.length, subtitleText.length,

      SpreadsheetApp.newTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(13).build())

    .build();

  sh.getRange(r, 1, 1, FULL_W).merge().setRichTextValue(richText)

    .setBackground('#263238').setHorizontalAlignment('center');

  r++;



  // Версия трекера — можно менять прямо в этой ячейке (не в коде).

  // "Версия" в updates.js читает именно отсюда, а не константу из кода —

  // так уведомления об обновлении зависят от того, когда ТЫ решил, что

  // обновился, а не от любой правки в самом коде.

  sh.getRange(r, 1).setValue(trackerVersion)

    .setFontSize(8).setFontColor('#9e9e9e').setFontStyle('italic')

    .setNote('Версия трекера — меняй здесь вручную после того, как реально обновил файлы. Используется для проверки обновлений (updates.js).');

  r++;



  mergedCell_(sh, r, 1, 1, COLS, '▌ РАСПРЕДЕЛЕНИЕ ПО КЛАССАМ',

    { bg: C.MID, fg: '#ffffff', bold: true });

  r++;



  hdrRow_(sh, r,

    ['Категория', 'Сумма, ₽', 'Текущий %', 'Цель %', 'Отклонение', 'Статус'],

    COLS);

  r++;



  let cats = Object.keys(config.classTargets);



  // Показываем также реальные категории,

  // даже если для них не задана целевая доля.

  positions.forEach(function(p) {

    if (cats.indexOf(p.category) === -1) {

      cats.push(p.category);

    }

  });

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

    let tgtPct  = findTarget_(p.name, p.ticker, config);

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

  sh.setColumnWidth(7, 24); // узкий отступ между левой и правой колонкой

  [180, 90, 90, 90, 90, 90].forEach(function(w, i) {

    sh.setColumnWidth(8 + i, w);

  });

  // Тонкая граница между колонками — на масштабе 60% пустой отступ сам

  // по себе плохо читается, с чёткой линией сразу видно, где раздел.

  sh.getRange(4, 6, 297, 1).setBorder(null, null, null, true, null, null,

    '#7986cb', SpreadsheetApp.BorderStyle.SOLID);

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

  let t0 = new Date().getTime();

  function lap_(label) {

    let now = new Date().getTime();

    Logger.log('syncAndRefresh: ' + label + ' — ' + ((now - t0) / 1000).toFixed(1) + ' сек от старта');

  }

  syncTinkoffPositions();
  lap_('syncTinkoffPositions завершён');

  if (typeof updateCashSummary_ === 'function') {
    updateCashSummary_();
    lap_('updateCashSummary_ завершён');
  }

  updateDashboard();
  lap_('updateDashboard завершён');

  calculateAveragePriceAndPL();
  lap_('calculateAveragePriceAndPL завершён');

  calculateYieldOnCost();
  lap_('calculateYieldOnCost завершён');

  calculateLdvEligibility();
  lap_('calculateLdvEligibility завершён');

  if (readIisAccountName_()) {
    calculateIisDeductionUsage();
    lap_('calculateIisDeductionUsage завершён');
  }

  if (Object.keys(readSectorMap_()).length) {
    calculateSectorDiversification();
    lap_('calculateSectorDiversification завершён');
  }

  if (readGoalTarget_() > 0) {
    calculateGoalProgress();
    lap_('calculateGoalProgress завершён');
  }

  updateIncomeSheet();
  lap_('updateIncomeSheet завершён');

  updateCalendarSheet();
  lap_('updateCalendarSheet завершён');

  hideDataSheets(false);
  lap_('hideDataSheets завершён');

  checkAndNotifyDeviations_();
  lap_('checkAndNotifyDeviations_ завершён');

  checkIisDividendHint_();
  lap_('checkIisDividendHint_ завершён');

  const updateInfo = checkForUpdates_();

  if (updateInfo.hasUpdate) {

    SpreadsheetApp.getActive().toast('Доступна новая версия: ' + updateInfo.latestVersion, 'Обновление трекера', 10);

  }

  lap_('checkForUpdates_ завершён');

  notifySyncComplete_();        lap_('notifySyncComplete_ завершён — всё готово');

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



function findTarget_(name, ticker, config) {
  let stockTargets = config.stockTargets || {};

  // 1. Точное совпадение по названию.
  if (stockTargets[name] !== undefined) return stockTargets[name];

  // 2. Точное совпадение по тикеру — надёжнее названия, если оно изменилось
  //    или где-то было сокращено для отображения.
  let tickerKey = normalizeTicker_(ticker);
  if (tickerKey && config.stockTargetsByTicker &&
      config.stockTargetsByTicker[tickerKey] !== undefined) {
    return config.stockTargetsByTicker[tickerKey];
  }

  // 3. Мягкое совпадение по названию как последний fallback.
  let nl = String(name || '').trim().toLowerCase();
  let keys = Object.keys(stockTargets);
  for (let i = 0; i < keys.length; i++) {
    let k = keys[i].toLowerCase();
    if (nl === k || nl.indexOf(k) >= 0 || k.indexOf(nl) >= 0) {
      return stockTargets[keys[i]];
    }
  }

  return 0;
}



function matchPos_(positions, name, ticker) {
  let tickerKey = normalizeTicker_(ticker);

  // Сначала тикер: он стабильнее отображаемого названия.
  if (tickerKey) {
    for (let i = 0; i < positions.length; i++) {
      if (normalizeTicker_(positions[i].ticker) === tickerKey) return positions[i];
    }
  }

  // Затем название.
  let nl = String(name || '').trim().toLowerCase();
  for (let i = 0; i < positions.length; i++) {
    let pnl = String(positions[i].name || '').trim().toLowerCase();
    if (pnl === nl || pnl.indexOf(nl) >= 0 || nl.indexOf(pnl) >= 0) return positions[i];
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



  let batchDate = new Date(); // метка этой конкретной пачки рекомендаций — для истории ребалансировок



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

    let pos    = matchPos_(sharePos, name, (config.stockTickers || {})[name] || '');

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



  let rebalanceLogItems = [];



  stockOrder.forEach(function(name, idx) {

    let info   = stockNeed[name];

    let isSkip = skipped.indexOf(name) !== -1;

    let res    = lotResult.results[name] || { lots: 0, actualAlloc: 0, lotCost: 0 };

    let ticker = info.pos ? info.pos.ticker : '';

    let actPct = totalRub > 0 ? info.actual / totalRub : 0;

    let lots   = isSkip ? 0 : res.lots;

    let alloc  = isSkip ? 0 : res.actualAlloc;



    if (!isSkip && alloc > 0) rebalanceLogItems.push({ name: name, amount: alloc });



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



  logRebalanceRecommendation_(batchDate, rebalanceLogItems);



  let otherCats   = ['Золото', 'Замещайки', 'Денежный рынок', 'Кэш'];

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

      comment = 'Категория на цели или выше цели';

    } else if (cat === 'Кэш') {

      instrName = '—';

      comment = 'Оставить ' + rub_(alloc) + ' в кэше';

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



  const keepVisible = [

    DST.CONFIG,

    DST.DASHBOARD,

    DST.REBALANCE,

    DST.INCOME,

    DST.CALENDAR,

    DST.HISTORY,

    '💵 Наличные',

  ];



  ss.getSheets().forEach(function(sh) {

    let name = sh.getName();



    if (keepVisible.indexOf(name) === -1) {

      sh.hideSheet();

    }

  });



  if (showAlert !== false) {

    SpreadsheetApp.getUi().alert('✅ Технические листы скрыты.');

  }

}