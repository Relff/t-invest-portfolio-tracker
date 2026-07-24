/**
 * sections.js — общий механизм секций на листе «Дашборд»
 */
function renderSection_(sh, title, contentFn) {
  let COLS = 6;
  let lastRow = sh.getLastRow();
  let data = lastRow > 0 ? sh.getRange(1, 1, lastRow, 1).getValues() : [];
  let sectionRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).indexOf(title) === 0) { sectionRow = i + 1; break; }
  }
  if (sectionRow > 0) {
    let rng = sh.getRange(sectionRow, 1, lastRow - sectionRow + 1, sh.getMaxColumns());
    rng.breakApart();
    rng.clear();
  }

  let r = sh.getLastRow() + 2;
  mergedCell_(sh, r, 1, 1, COLS, title, { bg: C.MID, fg: '#ffffff', bold: true });
  r++;
  let endRow = contentFn(sh, r, COLS) || r;

  // Перенос текста для всей секции целиком, чтобы длинные названия не вылезали
  sh.getRange(sh.getLastRow() > endRow ? sh.getLastRow() : endRow, 1, 1, COLS); // safety no-op
  sh.getRange(sectionRow > 0 ? sectionRow : r - 1, 1, sh.getLastRow() - (sectionRow > 0 ? sectionRow : r - 1) + 1, COLS)
    .setWrap(true);
}

/**
 * Крупная цветная плашка для ключевой метрики — по паттерну карточек
 * из income.js. Возвращает следующую свободную строку.
 */
function renderTile_(sh, r, COLS, label, value, bg) {
  sh.getRange(r, 1, 1, COLS).merge()
    .setValue(label + ':  ' + value)
    .setBackground(bg || '#1565c0').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  return r + 1;
}