/**
 * charts.js — Визуализация для листа Dashboard
 *
 * Круговая и столбчатая диаграммы («Распределение портфеля», «Текущее vs
 * Цель») переехали в HTML-дашборд (htmldashboard.js) — там они и так уже
 * были продублированы (пирог целиком, столбцы по сути то же самое, что
 * радар «Факт vs Цель»). Держать одно и то же в двух местах смысла не было.
 *
 * Здесь остались только текстовые прогресс-бары по акциям — это не
 * отдельный графический объект, а часть самой таблицы (юникод-полоски
 * прямо в ячейках), поэтому их решили не трогать.
 */

function addDashboardCharts() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;

  var data = sh.getDataRange().getValues();
  var stockData = extractStockData_(data);

  if (stockData.length > 0) {
    buildStockProgress_(sh, stockData, data.length);
  }

  SpreadsheetApp.flush();
}

function extractStockData_(data) {
  var result = [];
  var inStockSection = false;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var col0 = String(row[0] || '').trim();
    if (col0.indexOf('АКЦИИ') >= 0 && col0.indexOf('ДЕТАЛИЗАЦИЯ') >= 0) { inStockSection = true; continue; }
    if (inStockSection && col0 === 'Название') continue;
    if (inStockSection && col0 && col0 !== '' && row[2] > 0) {
      result.push({
        name: col0.length > 20 ? col0.substring(0, 20) + '…' : col0,
        actual: Number(row[3] || 0) * 100,
        target: Number(row[4] || 0) * 100,
      });
    }
    if (inStockSection && col0 === '' && result.length > 0) break;
  }
  return result;
}

function buildStockProgress_(sh, stockData, startRow) {
  var pr = startRow + 3;
  sh.getRange(pr, 1, 1, 6).merge()
    .setValue('▌ АКЦИИ — ПРОГРЕСС К ЦЕЛИ')
    .setBackground(C.MID).setFontColor('#ffffff').setFontWeight('bold');
  pr++;
  sh.getRange(pr, 1).setValue('Акция');
  sh.getRange(pr, 2).setValue('Текущий %');
  sh.getRange(pr, 3).setValue('Цель %');
  sh.getRange(pr, 4, 1, 3).merge().setValue('Прогресс');
  sh.getRange(pr, 1, 1, 6).setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold');
  pr++;

  stockData.forEach(function(stock, idx) {
    var bg = idx % 2 === 0 ? C.EVEN : C.ODD;
    var progress = stock.target > 0 ? Math.min(stock.actual / stock.target, 1) : 0;
    var bars = Math.round(progress * 20);
    var barStr = '█'.repeat(bars) + '░'.repeat(20 - bars);
    var barColor = progress >= 0.9 ? '#1b5e20' : progress >= 0.5 ? '#f57f17' : '#b71c1c';

    sh.getRange(pr, 1).setValue(stock.name).setBackground(bg);
    sh.getRange(pr, 2).setValue(stock.actual / 100).setNumberFormat('0.0%').setBackground(bg);
    sh.getRange(pr, 3).setValue(stock.target / 100).setNumberFormat('0.0%').setBackground(bg);
    sh.getRange(pr, 4, 1, 3).merge()
      .setValue(barStr + '  ' + Math.round(progress * 100) + '%')
      .setBackground(bg).setFontColor(barColor)
      .setFontFamily('Courier New').setFontSize(9);
    pr++;
  });
}

/**
 * Разовая уборка для тех, у кого на листе Дашборд ещё сидят старые
 * круговая/столбчатая диаграммы с прошлой версии — запусти один раз
 * из редактора Apps Script, дальше эта функция не понадобится, новые
 * графики на лист больше не добавляются.
 */
function removeAllCharts() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  var count = sh.getCharts().length;
  sh.getCharts().forEach(function(c) { sh.removeChart(c); });
  PropertiesService.getScriptProperties().deleteProperty('DASHBOARD_CHARTS_INITIALIZED');
  var helper = ss.getSheetByName('_chart_data');
  if (helper) ss.deleteSheet(helper);
  SpreadsheetApp.getUi().alert('✅ Удалено диаграмм: ' + count +
    '. Больше графики на этом листе не создаются — смотри их в HTML-дашборде (Tinkoff → 📊 Ребаланс и отчёты → Открыть HTML-дашборд).');
}
