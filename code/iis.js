/**
 * iis.js — ИИС-3: использование годового лимита налогового вычета
 *
 * У ИИС-3 нет лимита на сумму пополнения, но вычет 13-15% считается
 * только с суммы до 400 000 ₽ в год. Эта секция показывает, сколько
 * из этого лимита уже выбрано в текущем календарном году, чтобы
 * не потерять часть льготы к концу года.
 *
 * Зависимости: getAccounts_(), fetchOperations_() — history.js
 *              mergedCell_(), renderTile_(), rub_(), DST, C — dashboard.js / sections.js
 */

const IIS_DEDUCTION_LIMIT   = 400000; // актуально на 2026 год, если лимит изменится — поправь тут
const IIS_SECTION_TITLE     = '▌ ИИС-3 — ВЫЧЕТ ЗА ГОД';
const IIS_PROP              = 'IIS_DEDUCTION_DATA';

/**
 * Тихий расчёт без UI (SpreadsheetApp.getUi().alert упал бы, если вызвать
 * из триггера без интерфейса — например, из ежедневной проверки в
 * telegram.js). Возвращает объект с данными или null, если счёт ИИС-3
 * не настроен / не найден — без единого alert.
 */
function computeIisDeductionData_() {
  let accountName = readIisAccountName_();
  if (!accountName) return null;

  let accounts = getAccounts_();
  let account = accounts.find(function(a) { return a.name.trim() === accountName.trim(); });
  if (!account) return null;

  let now = new Date();
  let yearStart = new Date(now.getFullYear(), 0, 1);

  let contributed = 0;
  let ops = fetchOperations_(account.id, account.name, yearStart, now);
  ops.forEach(function(op) {
    if (op.type === 'OPERATION_TYPE_INPUT') contributed += op.amount;
  });

  let pct = Math.min(contributed / IIS_DEDUCTION_LIMIT, 1);
  let remaining = Math.max(IIS_DEDUCTION_LIMIT - contributed, 0);

  let yearEnd = new Date(now.getFullYear(), 11, 31);
  let daysLeft = Math.max(0, Math.ceil((yearEnd - now) / (24 * 3600 * 1000)));

  // Минимальный срок владения — считаем, только если дата открытия указана
  let openDate = readIisOpenDate_();
  let holdInfo = null;
  if (openDate) {
    let minYears = iisMinHoldYears_(openDate.getFullYear());
    let unlockDate = new Date(openDate.getFullYear() + minYears, openDate.getMonth(), openDate.getDate());
    let daysUntilUnlock = Math.ceil((unlockDate - now) / (24 * 3600 * 1000));
    holdInfo = {
      openDate: openDate.getTime(),
      minYears: minYears,
      unlockDate: unlockDate.getTime(),
      daysUntilUnlock: daysUntilUnlock, // отрицательное или 0 — минимальный срок уже прошёл
    };
  }

  return {
    accountName: account.name, year: now.getFullYear(),
    contributed: Math.round(contributed), remaining: Math.round(remaining), pct: pct,
    daysLeft: daysLeft, holdInfo: holdInfo,
  };
}

function calculateIisDeductionUsage() {
  let ui = SpreadsheetApp.getUi();
  let accountName = readIisAccountName_();
  if (!accountName) {
    ui.alert('Сначала укажи название своего счёта ИИС-3 в Config → Блок «ИИС-3».');
    return;
  }

  let accounts = getAccounts_();
  let account = accounts.find(function(a) { return a.name.trim() === accountName.trim(); });
  if (!account) {
    ui.alert('Счёт «' + accountName + '» не найден среди твоих открытых счетов. Проверь точное название в Config.');
    return;
  }

  let data = computeIisDeductionData_();
  if (!data) { ui.alert('Не получилось посчитать — проверь настройки ИИС-3 в Config.'); return; }

  PropertiesService.getScriptProperties().setProperty(IIS_PROP, JSON.stringify(data));

  redrawIisSection_();
}

function readIisAccountName_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) return '';
  let v = sh.getDataRange().getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'Название счёта ИИС-3') {
      return String(v[i][1]).trim();
    }
  }
  return '';
}

function readIisOpenDate_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) return null;
  let v = sh.getDataRange().getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'Дата открытия ИИС-3') {
      let raw = v[i][1];
      if (!raw) return null;
      let d = raw instanceof Date ? raw : new Date(String(raw).split('.').reverse().join('-'));
      return isNaN(d.getTime()) ? null : d;
    }
  }
  return null;
}

/**
 * Минимальный срок владения ИИС-3 зависит от года открытия (по действующим
 * на 2026 год правилам): 5 лет для счетов, открытых в 2024-2026 годах,
 * дальше +1 год за каждый год открытия вплоть до 10 лет к 2031 году
 * и позже. Если правила ещё раз изменятся — поправь формулу здесь.
 */
function iisMinHoldYears_(openYear) {
  if (openYear <= 2026) return 5;
  return Math.min(10, 5 + (openYear - 2026));
}

function addIisBlock() {
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) { ui.alert('⚠️ Сначала запустите initConfig()'); return; }

  let vals = sh.getDataRange().getValues();
  let blockHeaderRow = -1, accountRow = -1, dateRow = -1;
  for (let i = 0; i < vals.length; i++) {
    let label = String(vals[i][0]).trim();
    if (label.indexOf('ИИС-3') >= 0 && label.indexOf('▌') === 0) blockHeaderRow = i;
    if (label === 'Название счёта ИИС-3') accountRow = i;
    if (label === 'Дата открытия ИИС-3') dateRow = i;
  }

  if (blockHeaderRow === -1) {
    // Блока ещё нет вообще — создаём с нуля, как раньше
    let lastRow = sh.getLastRow() + 2;
    let block = [
      ['▌ ИИС-3 — отслеживание вычета', '', ''],
      ['Название счёта ИИС-3', '', '← впиши точное название счёта, как оно указано в Т-Инвестициях'],
      ['Дата открытия ИИС-3', '', '← необязательно, но без неё не посчитать минимальный срок владения (напр. 15.03.2024)'],
    ];
    sh.getRange(lastRow, 1, block.length, 3).setValues(block);
    sh.getRange(lastRow, 1, 1, 3).merge()
      .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
    sh.getRange(lastRow + 1, 2).setBackground(C.INPUT).setFontWeight('bold');
    sh.getRange(lastRow + 2, 2).setBackground(C.INPUT).setFontWeight('bold').setNumberFormat('dd.mm.yyyy');
    ui.alert('✅ Блок «ИИС-3» добавлен. Впиши название счёта и запусти «ИИС-3 — вычет за год».');
    return;
  }

  if (dateRow !== -1) {
    ui.alert('Блок «ИИС-3» уже содержит поле «Дата открытия ИИС-3» — добавлять нечего.');
    return;
  }

  // Блок уже есть (с прошлой версии), но нового поля с датой в нём нет —
  // дописываем строку сразу после «Название счёта ИИС-3», не трогая
  // остальной Config и уже введённое название счёта.
  let insertAfterRow = (accountRow !== -1 ? accountRow : blockHeaderRow) + 1; // 0-indexed → 1-indexed строка ПОСЛЕ
  sh.insertRowAfter(insertAfterRow);
  let newRow = insertAfterRow + 1;
  sh.getRange(newRow, 1, 1, 3).setValues([[
    'Дата открытия ИИС-3', '', '← необязательно, но без неё не посчитать минимальный срок владения (напр. 15.03.2024)'
  ]]);
  sh.getRange(newRow, 2).setBackground(C.INPUT).setFontWeight('bold').setNumberFormat('dd.mm.yyyy');

  ui.alert('✅ Добавил поле «Дата открытия ИИС-3» в существующий блок. Впиши дату и запусти «ИИС-3 — вычет за год».');
}

function redrawIisSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(IIS_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);

  renderSection_(sh, IIS_SECTION_TITLE, function(sh, r, COLS, colStart) {
    r = renderTile_(sh, r, COLS,
      '🏦 Использовано вычета за ' + d.year + ' год (' + d.accountName + ')',
      (d.pct * 100).toFixed(1) + '%', '#1565c0', colStart);

    // Прогресс-бар с рыжей обводкой вокруг закрашенной части
    let barLen = 24;
    let filled = Math.round(d.pct * barLen);
    let bar = '█'.repeat(filled) + '░'.repeat(barLen - filled);

    let barCell = sh.getRange(r, colStart, 1, 3).merge();
    barCell.setValue(bar)
      .setFontFamily('Courier New').setFontSize(11).setFontColor('#e65100').setBackground(C.EVEN);
    barCell.setBorder(true, true, true, true, false, false, '#e65100', SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(r, colStart + 3, 1, COLS - 3).merge()
      .setValue(rub_(d.contributed) + ' из ' + rub_(IIS_DEDUCTION_LIMIT))
      .setFontWeight('bold').setHorizontalAlignment('right').setBackground(C.EVEN);
    r++;

    if (d.remaining > 0) {
      sh.getRange(r, colStart, 1, COLS).merge()
        .setValue('Осталось довнести до конца года для максимального вычета: ' + rub_(d.remaining) +
                  '  ·  осталось дней: ' + (d.daysLeft !== undefined ? d.daysLeft : '—'))
        .setFontColor(C.WARN).setFontStyle('italic').setFontSize(10).setBackground(C.EVEN);
    } else {
      sh.getRange(r, colStart, 1, COLS).merge()
        .setValue('✅ Лимит вычета за этот год полностью выбран')
        .setFontColor(C.OK).setFontWeight('bold').setFontSize(10).setBackground(C.EVEN);
    }
    r++;

    let tz = Session.getScriptTimeZone();
    if (d.holdInfo) {
      let h = d.holdInfo;
      let unlockStr = Utilities.formatDate(new Date(h.unlockDate), tz, 'dd.MM.yyyy');
      if (h.daysUntilUnlock > 0) {
        let years = Math.floor(h.daysUntilUnlock / 365);
        let restDays = h.daysUntilUnlock % 365;
        sh.getRange(r, colStart, 1, COLS).merge()
          .setValue('⏳ Мин. срок владения — ' + h.minYears + ' лет, до ' + unlockStr +
                    ' осталось ~' + years + ' г. ' + restDays + ' дн. (закрыть раньше — потерять все льготы)')
          .setFontColor('#9e9e9e').setFontStyle('italic').setFontSize(9).setBackground(C.EVEN);
      } else {
        sh.getRange(r, colStart, 1, COLS).merge()
          .setValue('✅ Минимальный срок владения (' + h.minYears + ' лет) уже прошёл — ' + unlockStr)
          .setFontColor(C.OK).setFontStyle('italic').setFontSize(9).setBackground(C.EVEN);
      }
      r++;
    } else {
      sh.getRange(r, colStart, 1, COLS).merge()
        .setValue('ℹ️ Впиши «Дата открытия ИИС-3» в Config, чтобы видеть минимальный срок владения')
        .setFontColor('#9e9e9e').setFontStyle('italic').setFontSize(9).setBackground(C.EVEN);
      r++;
    }
    return r;
  }, 'left');
}