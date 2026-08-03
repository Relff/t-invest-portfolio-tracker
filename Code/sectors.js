/**
 * sectors.js — Секторная диверсификация акций
 *
 * Мосбиржа не отдаёт единую отраслевую классификацию через T-Invest API
 * (у разных источников разные сектора для одной и той же бумаги), поэтому
 * сектор задаётся вручную в Config — то же решение, что уже используется
 * для «Маппинга спецкатегорий»: точность важнее автоматизации.
 *
 * При первом добавлении блок сам подставляет список текущих акций из
 * портфеля — остаётся только вписать сектор напротив каждой.
 *
 * Зависимости: readConfig_(), readPositions_(), rub_(), DST, C — dashboard.js
 *              renderSection_() — sections.js
 */

const SECTOR_PROP          = 'SECTOR_DIVERSIFICATION_DATA';
const SECTOR_SECTION_TITLE = '▌ СЕКТОРНАЯ ДИВЕРСИФИКАЦИЯ (АКЦИИ)';

function readSectorMap_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  let map = {};
  if (!sh) return map;

  let v = sh.getDataRange().getValues();
  let inBlock = false;
  for (let i = 0; i < v.length; i++) {
    let col0 = String(v[i][0]).trim();
    if (col0.indexOf('СЕКТОРЫ АКЦИЙ') >= 0) { inBlock = true; continue; }
    if (!inBlock) continue;
    if (col0 === 'Название акции') continue; // строка заголовка таблицы
    if (col0 === '') break; // блок закончился
    let sector = String(v[i][1] || '').trim();
    if (sector) map[col0] = sector;
  }
  return map;
}

function addSectorsBlock() {
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) { ui.alert('⚠️ Сначала запустите initConfig()'); return; }

  let vals = sh.getDataRange().getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).indexOf('СЕКТОРЫ АКЦИЙ') >= 0) {
      ui.alert('Блок «Секторы акций» уже существует.');
      return;
    }
  }

  // Пытаемся сразу подставить текущие акции из портфеля — если он уже
  // синхронизирован. Если ещё нет — просто добавляем пустой блок,
  // без ошибок.
  let shareNames = [];
  try {
    let config    = readConfig_();
    let positions = readPositions_(config);
    shareNames = positions.filter(function(p) { return p.category === 'Акции'; })
                           .map(function(p) { return p.name; });
  } catch (e) { /* портфель ещё не синхронизирован — оставляем блок пустым */ }

  let lastRow = sh.getLastRow() + 2;
  let block = [
    ['▌ СЕКТОРЫ АКЦИЙ (опционально)', '', ''],
    ['Впиши сектор для каждой акции свободным текстом — например «Финансы», «Нефть и газ», «Металлургия», «IT»', '', ''],
    ['Название акции', 'Сектор', ''],
  ];
  shareNames.forEach(function(name) { block.push([name, '', '']); });

  sh.getRange(lastRow, 1, block.length, 3).setValues(block);
  sh.getRange(lastRow, 1, 1, 3).merge()
    .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.getRange(lastRow + 1, 1, 1, 3).merge()
    .setFontColor('#9e9e9e').setFontStyle('italic');
  sh.getRange(lastRow + 2, 1, 1, 3)
    .setBackground(C.MID).setFontColor('#ffffff').setFontWeight('bold');
  if (shareNames.length) {
    sh.getRange(lastRow + 3, 2, shareNames.length, 1).setBackground(C.INPUT);
  }

  ui.alert('✅ Блок «Секторы акций» добавлен' +
    (shareNames.length ? ' с заготовкой из ' + shareNames.length + ' текущих акций' : '') +
    '. Впиши сектор напротив каждой (колонка «Сектор») и запусти «Секторная диверсификация».');
}

function calculateSectorDiversification() {
  let ui = SpreadsheetApp.getUi();
  let config;
  try { config = readConfig_(); }
  catch (e) { ui.alert('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  let shares    = positions.filter(function(p) { return p.category === 'Акции'; });
  if (!shares.length) { ui.alert('Нет акций в портфеле.'); return; }

  let sectorMap = readSectorMap_();
  if (!Object.keys(sectorMap).length) {
    ui.alert('Сначала заполни блок «Секторы акций» в Config (Tinkoff → ⚙️ Настройки → Добавить блок секторов акций).');
    return;
  }

  let sharesTotal = shares.reduce(function(s, p) { return s + p.valueRub; }, 0);
  let bySector = {};
  let unmapped = 0;

  shares.forEach(function(p) {
    let sector = sectorMap[p.name];
    if (!sector) { unmapped += p.valueRub; return; }
    bySector[sector] = (bySector[sector] || 0) + p.valueRub;
  });

  let rows = Object.keys(bySector).map(function(sector) {
    return {
      sector: sector,
      valueRub: Math.round(bySector[sector]),
      pct: sharesTotal > 0 ? bySector[sector] / sharesTotal : 0,
    };
  }).sort(function(a, b) { return b.valueRub - a.valueRub; });

  PropertiesService.getScriptProperties().setProperty(SECTOR_PROP, JSON.stringify({
    rows: rows, unmapped: Math.round(unmapped), sharesTotal: Math.round(sharesTotal),
  }));

  redrawSectorSection_();
}

function redrawSectorSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(SECTOR_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);

  renderSection_(sh, SECTOR_SECTION_TITLE, function(sh, r, COLS, colStart) {
    sh.getRange(r, colStart, 1, 3).merge().setValue('Сектор');
    sh.getRange(r, colStart + 3, 1, COLS - 3).merge().setValue('Доля от акций').setHorizontalAlignment('right');
    sh.getRange(r, colStart, 1, COLS).setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold');
    r++;

    let barLen = 18;
    d.rows.forEach(function(row, idx) {
      let bg     = idx % 2 === 0 ? C.EVEN : C.ODD;
      let filled = Math.round(row.pct * barLen);
      let bar    = '█'.repeat(filled) + '░'.repeat(barLen - filled);

      sh.getRange(r, colStart, 1, 3).merge().setValue(row.sector).setBackground(bg);
      sh.getRange(r, colStart + 3, 1, COLS - 3).merge()
        .setValue(bar + '  ' + (row.pct * 100).toFixed(1) + '%  ·  ' + rub_(row.valueRub))
        .setFontFamily('Courier New').setFontSize(9).setBackground(bg).setHorizontalAlignment('right');
      r++;
    });

    if (d.unmapped > 0) {
      let unmappedPct = d.sharesTotal > 0 ? (d.unmapped / d.sharesTotal * 100).toFixed(1) : '0';
      sh.getRange(r, colStart, 1, COLS).merge()
        .setValue('⚠️ Без сектора: ' + rub_(d.unmapped) + ' (' + unmappedPct + '%) — дозаполни блок «Секторы акций» в Config')
        .setFontColor(C.WARN).setFontStyle('italic').setFontSize(9);
      r++;
    }
    return r;
  }, 'right');
}
