/**
 * dashboard_charts.js — Визуализация для листа Dashboard
 *
 * Добавить в Apps Script как новый файл "charts".
 * Вызывается автоматически из updateDashboard().
 *
 * Что добавляет:
 *   1. Круговая диаграмма — текущее распределение по классам
 *   2. Столбчатая диаграмма — Текущее % vs Цель % по классам
 *   3. Горизонтальные прогресс-бары по акциям
 */

// ════════════════════════════════════════════════════════════════════
// ГЛАВНАЯ ФУНКЦИЯ — вызывается из updateDashboard()
// Добавь в конец updateDashboard(): addDashboardCharts();
// ════════════════════════════════════════════════════════════════════

function addDashboardCharts() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;

  // Удаляем старые диаграммы
  let charts = sh.getCharts();
  charts.forEach(function(c) { sh.removeChart(c); });

  // Очищаем вспомогательный лист ОДИН РАЗ (чтобы данные пирога не затёрлись баром)
  getOrCreateHelper_(ss).clearContents();

  // Читаем данные из листа
  let data = sh.getDataRange().getValues();

  // Находим строки с данными по классам (ищем по структуре)
  let classData  = extractClassData_(data);
  let stockData  = extractStockData_(data);

  if (classData.length === 0) return;

  // Диаграмма 1: Круговая — текущее распределение
  buildPieChart_(ss, sh, classData);

  // Диаграмма 2: Столбчатая — текущее vs цель
  buildBarChart_(ss, sh, classData);

  // Диаграмма 3: Прогресс-бары акций (через мини-таблицу)
  if (stockData.length > 0) {
    buildStockProgress_(sh, stockData, data.length);
  }

  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════════════
// ИЗВЛЕЧЕНИЕ ДАННЫХ ИЗ ЛИСТА
// ════════════════════════════════════════════════════════════════════

function extractClassData_(data) {
  let result = [];
  let inClassSection = false;

  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let col0 = String(row[0] || '').trim();

    // Ищем заголовок секции классов
    if (col0.indexOf('РАСПРЕДЕЛЕНИЕ ПО КЛАССАМ') >= 0) {
      inClassSection = true;
      continue;
    }

    // Строка заголовка таблицы
    if (inClassSection && col0 === 'Категория') continue;

    // Данные по классам
    if (inClassSection && col0 && row[1] > 0) {
      // Проверяем что это строка данных (не заголовок секции акций)
      if (col0.indexOf('АКЦИИ') >= 0 || col0.indexOf('▌') >= 0) break;
      result.push({
        name:    col0,
        actual:  Number(row[2] || 0) * 100,  // текущий %
        target:  Number(row[3] || 0) * 100,  // целевой %
        valueRub: Number(row[1] || 0),
      });
    }

    // Конец секции классов
    if (inClassSection && col0 === '' && result.length > 0) break;
  }

  return result;
}

function extractStockData_(data) {
  let result = [];
  let inStockSection = false;

  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let col0 = String(row[0] || '').trim();

    if (col0.indexOf('АКЦИИ') >= 0 && col0.indexOf('ДЕТАЛИЗАЦИЯ') >= 0) {
      inStockSection = true;
      continue;
    }
    if (inStockSection && col0 === 'Название') continue;

    if (inStockSection && col0 && col0 !== '' && row[2] > 0) {
      result.push({
        name:    col0.length > 20 ? col0.substring(0, 20) + '…' : col0,
        actual:  Number(row[3] || 0) * 100,
        target:  Number(row[4] || 0) * 100,
      });
    }

    if (inStockSection && col0 === '' && result.length > 0) break;
  }

  return result;
}


// ════════════════════════════════════════════════════════════════════
// ДИАГРАММА 1: КРУГОВАЯ — ТЕКУЩЕЕ РАСПРЕДЕЛЕНИЕ
// ════════════════════════════════════════════════════════════════════

function buildPieChart_(ss, sh, classData) {
  // Создаём вспомогательный диапазон данных для диаграммы
  let helperSheet = getOrCreateHelper_(ss);
  let pieRows = [['Класс', 'Сумма, ₽']];
  classData.forEach(function(d) {
    if (d.valueRub > 0) pieRows.push([d.name, d.valueRub]);
  });

  let pieRange = helperSheet.getRange(1, 1, pieRows.length, 2);
  pieRange.setValues(pieRows);

  let chart = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(pieRange)
    .setOption('title', 'Распределение портфеля')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1a237e' })
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .setOption('pieSliceTextStyle', { fontSize: 10 })
    .setOption('colors', ['#1565c0', '#0d47a1', '#ffd54f', '#43a047', '#ef6c00'])
    .setOption('backgroundColor', '#ffffff')
    .setPosition(5, 8, 5, 5)   // строка 5, колонка H, отступы 5px
    .build();

  sh.insertChart(chart);
}


// ════════════════════════════════════════════════════════════════════
// ДИАГРАММА 2: СТОЛБЧАТАЯ — ТЕКУЩЕЕ vs ЦЕЛЬ
// ════════════════════════════════════════════════════════════════════

function buildBarChart_(ss, sh, classData) {
  let helperSheet = getOrCreateHelper_(ss);
  let barRows = [['Класс', 'Текущий %', 'Цель %']];
  classData.forEach(function(d) {
    barRows.push([d.name, Math.round(d.actual * 10) / 10, Math.round(d.target * 10) / 10]);
  });

  let barRange = helperSheet.getRange(20, 1, barRows.length, 3);
  barRange.setValues(barRows);

  let chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(barRange)
    .setOption('title', 'Текущее vs Цель (%)')
    .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#1a237e' })
    .setOption('legend', { position: 'top' })
    .setOption('colors', ['#1565c0', '#ffd54f'])
    .setOption('backgroundColor', '#ffffff')
    .setOption('vAxis', { title: '%', minValue: 0, maxValue: 90 })
    .setOption('hAxis', { textStyle: { fontSize: 9 } })
    .setPosition(5, 14, 5, 5)  // строка 5, колонка N
    .build();

  sh.insertChart(chart);
}


// ════════════════════════════════════════════════════════════════════
// ДИАГРАММА 3: ПРОГРЕСС-БАРЫ АКЦИЙ
// ════════════════════════════════════════════════════════════════════

function buildStockProgress_(sh, stockData, startRow) {
  // Рисуем мини-таблицу прогресс-баров через условное форматирование
  let pr = startRow + 3; // строка после основной таблицы

  sh.getRange(pr, 1, 1, 8).merge()
    .setValue('▌ АКЦИИ — ПРОГРЕСС К ЦЕЛИ')
    .setBackground(C.MID).setFontColor('#ffffff').setFontWeight('bold');
  pr++;

  sh.getRange(pr, 1).setValue('Акция');
  sh.getRange(pr, 2).setValue('Текущий %');
  sh.getRange(pr, 3).setValue('Цель %');
  sh.getRange(pr, 4, 1, 5).merge().setValue('Прогресс');
  sh.getRange(pr, 1, 1, 8)
    .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold');
  pr++;

  stockData.forEach(function(stock, idx) {
    let bg = idx % 2 === 0 ? C.EVEN : C.ODD;
    let progress = stock.target > 0 ? Math.min(stock.actual / stock.target, 1) : 0;
    let pctText  = (stock.actual).toFixed(1) + '% / ' + (stock.target).toFixed(1) + '%';
    let bars     = Math.round(progress * 20); // макс 20 блоков
    let barStr   = '█'.repeat(bars) + '░'.repeat(20 - bars);
    let barColor = progress >= 0.9 ? '#1b5e20' : progress >= 0.5 ? '#f57f17' : '#b71c1c';

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


// ════════════════════════════════════════════════════════════════════
// ВСПОМОГАТЕЛЬНЫЙ СКРЫТЫЙ ЛИСТ ДЛЯ ДАННЫХ ДИАГРАММ
// ════════════════════════════════════════════════════════════════════

function getOrCreateHelper_(ss) {
  let name = '_chart_data';
  let sh   = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.hideSheet();
  }
  // Не очищаем здесь - очистка делается один раз в addDashboardCharts
  return sh;
}


// ════════════════════════════════════════════════════════════════════
// УДАЛИТЬ ВСЕ ДИАГРАММЫ (утилита на случай если надо сбросить)
// ════════════════════════════════════════════════════════════════════

function removeAllCharts() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  sh.getCharts().forEach(function(c) { sh.removeChart(c); });
  SpreadsheetApp.getUi().alert('✅ Все диаграммы удалены.');
}
