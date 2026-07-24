/**
 * charts.js — Визуализация для листа Dashboard
 *
 * Графики создаются один раз и больше не пересоздаются — это отслеживается
 * флагом в Script Properties, а не сравнением заголовков (оказалось ненадёжным).
 * Позиция и размер, куда бы их ни перетащили руками, сохраняются насовсем.
 * Сбросить можно только через removeAllCharts() — она же сбрасывает флаг.
 */

const CHARTS_INIT_PROP = 'DASHBOARD_CHARTS_INITIALIZED';

function addDashboardCharts() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;

  var helperSheet = getOrCreateHelper_(ss);
  helperSheet.clearContents();

  var data = sh.getDataRange().getValues();
  var classData  = extractClassData_(data);
  var stockData  = extractStockData_(data);

  if (classData.length === 0) return;

  // Данные под графиками обновляем всегда — сами графики не трогаем,
  // если уже были созданы раньше (флаг в Properties).
  writePieData_(helperSheet, classData);
  writeBarData_(helperSheet, classData);

  var initialized = PropertiesService.getScriptProperties().getProperty(CHARTS_INIT_PROP) === 'true';
  if (!initialized) {
    buildPieChart_(sh, helperSheet);
    buildBarChart_(sh, helperSheet);
    PropertiesService.getScriptProperties().setProperty(CHARTS_INIT_PROP, 'true');
  }

  if (stockData.length > 0) {
    buildStockProgress_(sh, stockData, data.length);
  }

  SpreadsheetApp.flush();
}

function writePieData_(helperSheet, classData) {
  var pieRows = [['Класс', 'Сумма, ₽']];
  classData.forEach(function(d) {
    if (d.valueRub > 0) pieRows.push([d.name, d.valueRub]);
  });
  helperSheet.getRange(1, 1, pieRows.length, 2).setValues(pieRows);
}

function writeBarData_(helperSheet, classData) {
  var barRows = [['Класс', 'Текущий %', 'Цель %']];
  classData.forEach(function(d) {
    barRows.push([d.name, Math.round(d.actual * 10) / 10, Math.round(d.target * 10) / 10]);
  });
  helperSheet.getRange(20, 1, barRows.length, 3).setValues(barRows);
}

function buildPieChart_(sh, helperSheet) {
  var lastRow = helperSheet.getRange(1, 1, 1, 1).getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  var pieRange = helperSheet.getRange(1, 1, Math.max(lastRow, 2), 2);

  var chart = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(pieRange)
    .setOption('title', 'Распределение портфеля')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1a237e' })
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .setOption('pieSliceTextStyle', { fontSize: 10 })
    .setOption('colors', ['#1565c0', '#0d47a1', '#ffd54f', '#43a047', '#ef6c00'])
    .setOption('backgroundColor', '#ffffff')
    .setOption('width', 320)
    .setOption('height', 240)
    .setPosition(1, 8, 0, 0)
    .build();

  sh.insertChart(chart);
}

function buildBarChart_(sh, helperSheet) {
  var lastRow = helperSheet.getRange(20, 1, 1, 1).getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
  var barRange = helperSheet.getRange(20, 1, Math.max(lastRow - 19, 2), 3);

  var chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(barRange)
    .setOption('title', 'Текущее vs Цель (%)')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1a237e' })
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#1565c0', '#ffd54f'])
    .setOption('backgroundColor', '#ffffff')
    .setOption('vAxis', { title: '%', minValue: 0, maxValue: 90 })
    .setOption('hAxis', { textStyle: { fontSize: 9 } })
    .setOption('width', 340)
    .setOption('height', 240)
    .setPosition(1, 14, 0, 0)
    .build();

  sh.insertChart(chart);
}

function extractClassData_(data) {
  var result = [];
  var inClassSection = false;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var col0 = String(row[0] || '').trim();
    if (col0.indexOf('РАСПРЕДЕЛЕНИЕ ПО КЛАССАМ') >= 0) { inClassSection = true; continue; }
    if (inClassSection && col0 === 'Категория') continue;
    if (inClassSection && col0 && row[1] > 0) {
      if (col0.indexOf('АКЦИИ') >= 0 || col0.indexOf('▌') >= 0) break;
      result.push({
        name: col0,
        actual: Number(row[2] || 0) * 100,
        target: Number(row[3] || 0) * 100,
        valueRub: Number(row[1] || 0),
      });
    }
    if (inClassSection && col0 === '' && result.length > 0) break;
  }
  return result;
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
  sh.getRange(pr, 1, 1, 8).merge()
    .setValue('▌ АКЦИИ — ПРОГРЕСС К ЦЕЛИ')
    .setBackground(C.MID).setFontColor('#ffffff').setFontWeight('bold');
  pr++;
  sh.getRange(pr, 1).setValue('Акция');
  sh.getRange(pr, 2).setValue('Текущий %');
  sh.getRange(pr, 3).setValue('Цель %');
  sh.getRange(pr, 4, 1, 5).merge().setValue('Прогресс');
  sh.getRange(pr, 1, 1, 8).setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold');
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
    sh.getRange(pr, 4, 1, 5).merge()
      .setValue(barStr + '  ' + Math.round(progress * 100) + '%')
      .setBackground(bg).setFontColor(barColor)
      .setFontFamily('Courier New').setFontSize(9);
    pr++;
  });
}

function getOrCreateHelper_(ss) {
  var name = '_chart_data';
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.hideSheet();
  }
  return sh;
}

function removeAllCharts() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  sh.getCharts().forEach(function(c) { sh.removeChart(c); });
  PropertiesService.getScriptProperties().deleteProperty(CHARTS_INIT_PROP);
  SpreadsheetApp.getUi().alert('✅ Все диаграммы удалены. Флаг сброшен — при следующем обновлении Dashboard графики создадутся заново.');
}