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

function calculateIisDeductionUsage() {
  let accountName = readIisAccountName_();
  if (!accountName) {
    SpreadsheetApp.getUi().alert('Сначала укажи название своего счёта ИИС-3 в Config → Блок «ИИС-3».');
    return;
  }

  let accounts = getAccounts_();
  let account = accounts.find(function(a) { return a.name.trim() === accountName.trim(); });
  if (!account) {
    SpreadsheetApp.getUi().alert('Счёт «' + accountName + '» не найден среди твоих открытых счетов. Проверь точное название в Config.');
    return;
  }

  let now = new Date();
  let yearStart = new Date(now.getFullYear(), 0, 1);

  let contributed = 0;
  let ops = fetchOperations_(account.id, account.name, yearStart, now);
  ops.forEach(function(op) {
    if (op.type === 'OPERATION_TYPE_INPUT') contributed += op.amount;
  });

  let pct = Math.min(contributed / IIS_DEDUCTION_LIMIT, 1);
  let remaining = Math.max(IIS_DEDUCTION_LIMIT - contributed, 0);

  PropertiesService.getScriptProperties().setProperty(IIS_PROP, JSON.stringify({
    accountName: account.name, year: now.getFullYear(),
    contributed: Math.round(contributed), remaining: Math.round(remaining), pct: pct,
  }));

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

function addIisBlock() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) { SpreadsheetApp.getUi().alert('⚠️ Сначала запустите initConfig()'); return; }

  let vals = sh.getDataRange().getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).indexOf('ИИС-3') >= 0) {
      SpreadsheetApp.getUi().alert('Блок «ИИС-3» уже существует.');
      return;
    }
  }

  let lastRow = sh.getLastRow() + 2;
  let block = [
    ['▌ ИИС-3 — отслеживание вычета', '', ''],
    ['Название счёта ИИС-3', '', '← впиши точное название счёта, как оно указано в Т-Инвестициях'],
  ];
  sh.getRange(lastRow, 1, block.length, 3).setValues(block);
  sh.getRange(lastRow, 1, 1, 3).merge()
    .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.getRange(lastRow + 1, 2).setBackground(C.INPUT).setFontWeight('bold');

  SpreadsheetApp.getUi().alert('✅ Блок «ИИС-3» добавлен. Впиши название счёта и запусти «ИИС-3 — вычет за год».');
}

function redrawIisSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(IIS_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);

  renderSection_(sh, IIS_SECTION_TITLE, function(sh, r, COLS) {
    r = renderTile_(sh, r, COLS,
      '🏦 Использовано вычета за ' + d.year + ' год (' + d.accountName + ')',
      (d.pct * 100).toFixed(1) + '%', '#1565c0');

    // Прогресс-бар с рыжей обводкой вокруг закрашенной части
    let barLen = 24;
    let filled = Math.round(d.pct * barLen);
    let bar = '█'.repeat(filled) + '░'.repeat(barLen - filled);

    let barCell = sh.getRange(r, 1, 1, 3).merge();
    barCell.setValue(bar)
      .setFontFamily('Courier New').setFontSize(11).setFontColor('#e65100').setBackground(C.EVEN);
    barCell.setBorder(true, true, true, true, false, false, '#e65100', SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(r, 4, 1, COLS - 3).merge()
      .setValue(rub_(d.contributed) + ' из ' + rub_(IIS_DEDUCTION_LIMIT))
      .setFontWeight('bold').setHorizontalAlignment('right').setBackground(C.EVEN);
    r++;

    if (d.remaining > 0) {
      sh.getRange(r, 1, 1, COLS).merge()
        .setValue('Осталось довнести до конца года для максимального вычета: ' + rub_(d.remaining))
        .setFontColor(C.WARN).setFontStyle('italic').setFontSize(10).setBackground(C.EVEN);
    } else {
      sh.getRange(r, 1, 1, COLS).merge()
        .setValue('✅ Лимит вычета за этот год полностью выбран')
        .setFontColor(C.OK).setFontWeight('bold').setFontSize(10).setBackground(C.EVEN);
    }
    return r + 1;
  });
}