/**
 * sections.js — общий механизм секций на листе «Дашборд»
 *
 * Два независимых «трека» — левая колонка (track='left', столбцы A-F)
 * и правая (track='right', столбцы H-M, с одним пустым столбцом G между
 * ними как визуальный отступ). У каждого трека своя высота — они растут
 * вниз независимо друг от друга, поэтому весь дашборд теперь помещается
 * в разы меньше строк по вертикали.
 */
const TRACK_COL_START_ = { left: 1, right: 8 };

// Шапка «ДАШБОРД ПОРТФЕЛЯ» + «Обновлено...» — 2 строки, растянутые на всю
// ширину листа (обе колонки). У объединённой ячейки значение хранится
// только в первой (левой) ячейке — колонки H-M внутри неё физически
// пустые при чтении. Поэтому trackLastRow_() для правого трека их не
// увидит — нужен явный "пол", иначе первая правая секция попытается
// начаться со 2-й строки, которая уже занята объединением шапки.
const HEADER_ROWS_ = 2;

function trackColStart_(track) {
  return TRACK_COL_START_[track] || TRACK_COL_START_.left;
}

/**
 * Находит последнюю занятую строку СТРОГО в пределах столбцов одного
 * трека (не по всему листу — sh.getLastRow() тут не подходит, он смотрит
 * на весь лист сразу и не различает треки).
 */
function trackLastRow_(sh, colStart, colSpan) {
  let maxRows = sh.getMaxRows();
  if (maxRows === 0) return 0;
  let vals = sh.getRange(1, colStart, maxRows, colSpan).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    for (let j = 0; j < vals[i].length; j++) {
      if (vals[i][j] !== '' && vals[i][j] !== null) return i + 1;
    }
  }
  return 0;
}

function renderSection_(sh, title, contentFn, track, headerBg) {
  track = track || 'left';
  let COLS = 6;
  let colStart = trackColStart_(track);

  let lastRow = trackLastRow_(sh, colStart, COLS);
  let data = lastRow > 0 ? sh.getRange(1, colStart, lastRow, 1).getValues() : [];
  let sectionRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).indexOf(title) === 0) { sectionRow = i + 1; break; }
  }
  if (sectionRow > 0) {
    let rng = sh.getRange(sectionRow, colStart, lastRow - sectionRow + 1, COLS);
    rng.breakApart();
    rng.clear();
    lastRow = trackLastRow_(sh, colStart, COLS);
  }

  // Новая секция никогда не начинается раньше, чем закончится шапка,
  // даже если trackLastRow_ для этого трека честно вернул 0.
  let r = Math.max(lastRow, HEADER_ROWS_) + 2;
  mergedCell_(sh, r, colStart, 1, COLS, title, { bg: headerBg || C.MID, fg: '#ffffff', bold: true });
  r++;
  let endRow = contentFn(sh, r, COLS, colStart) || r;

  // Перенос текста для всей секции целиком, чтобы длинные названия не вылезали
  let finalLastRow = trackLastRow_(sh, colStart, COLS);
  let wrapFrom = sectionRow > 0 ? sectionRow : r - 1;
  sh.getRange(wrapFrom, colStart, Math.max(finalLastRow, endRow) - wrapFrom + 1, COLS).setWrap(true);
}

/**
 * Крупная цветная плашка для ключевой метрики — по паттерну карточек
 * из income.js. Возвращает следующую свободную строку.
 */
function renderTile_(sh, r, COLS, label, value, bg, colStart) {
  colStart = colStart || 1;
  sh.getRange(r, colStart, 1, COLS).merge()
    .setValue(label + ':  ' + value)
    .setBackground(bg || '#1565c0').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  return r + 1;
}