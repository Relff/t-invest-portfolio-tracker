/**
 * goal.js — Прогресс к цели (целевая сумма портфеля)
 *
 * Зависимости: readConfig_(), readPositions_(), rub_(), DST, C — dashboard.js
 *              renderSection_() — sections.js
 */

const GOAL_PROP          = 'GOAL_PROGRESS_DATA';
const GOAL_SECTION_TITLE = '▌ ПРОГРЕСС К ЦЕЛИ';

function calculateGoalProgress() {
  let config;
  try { config = readConfig_(); }
  catch (e) { SpreadsheetApp.getUi().alert('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);

  let targetSum = readGoalTarget_();
  if (!targetSum || targetSum <= 0) {
    SpreadsheetApp.getUi().alert('Целевая сумма не задана. Заполни блок «Цель портфеля» в Config.');
    return;
  }

  let pct = Math.min(totalRub / targetSum, 1);
  PropertiesService.getScriptProperties().setProperty(GOAL_PROP, JSON.stringify({
    total: Math.round(totalRub), target: Math.round(targetSum), pct: pct,
  }));
  redrawGoalSection_();
}

function readGoalTarget_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) return 0;
  let v = sh.getDataRange().getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'Целевая сумма, ₽') {
      return Number(v[i][1]) || 0;
    }
  }
  return 0;
}

function addGoalBlock() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) { SpreadsheetApp.getUi().alert('⚠️ Сначала запустите initConfig()'); return; }

  let vals = sh.getDataRange().getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).indexOf('ЦЕЛЬ ПОРТФЕЛЯ') >= 0) {
      SpreadsheetApp.getUi().alert('Блок «Цель портфеля» уже существует.');
      return;
    }
  }

  let lastRow = sh.getLastRow() + 2;
  let block = [
    ['▌ ЦЕЛЬ ПОРТФЕЛЯ (опционально)', '', ''],
    ['Целевая сумма, ₽', 0, '← впиши сумму, к которой идёшь (например, 5000000)'],
  ];
  sh.getRange(lastRow, 1, block.length, 3).setValues(block);
  sh.getRange(lastRow, 1, 1, 3).merge()
    .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.getRange(lastRow + 1, 2).setBackground(C.INPUT).setFontWeight('bold');

  SpreadsheetApp.getUi().alert('✅ Блок «Цель портфеля» добавлен. Впиши сумму в ячейку и запусти «Прогресс к цели».');
}

function redrawGoalSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(GOAL_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);

  renderSection_(sh, GOAL_SECTION_TITLE, function(sh, r, COLS) {
    let barLen  = 20;
    let filled  = Math.round(d.pct * barLen);
    let bar     = '█'.repeat(filled) + '░'.repeat(barLen - filled);
    let clr     = d.pct >= 1 ? C.OK : d.pct >= 0.5 ? C.WARN : C.CRIT;

    sh.getRange(r, 1, 1, 3).merge().setValue('Прогресс к цели (' + rub_(d.target) + ')').setFontWeight('bold');
    sh.getRange(r, 4, 1, COLS - 3).merge().setValue((d.pct * 100).toFixed(1) + '%')
      .setFontColor(clr).setFontWeight('bold').setHorizontalAlignment('right');
    sh.getRange(r, 1, 1, COLS).setBackground(C.EVEN);
    r++;

    sh.getRange(r, 1, 1, COLS).merge()
      .setValue(bar + '  ' + rub_(d.total) + ' из ' + rub_(d.target))
      .setFontFamily('Courier New').setFontSize(10).setFontColor(clr).setBackground(C.EVEN);
  });
}