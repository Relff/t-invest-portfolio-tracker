/**
 * report.js — Годовой отчёт одной кнопкой (Google Doc)
 *
 * Собирает воедино уже посчитанные метрики (XIRR, бенчмарк, дисциплина,
 * концентрация — из Script Properties) плюс свежую статистику операций
 * за последние 12 месяцев, оформляет как Google-документ.
 *
 * Зависимости: getAccounts_(), fetchOperations_(), buildSummary_() — history.js
 *              readConfig_(), readPositions_(), rub_() — dashboard.js
 *              ANALYTICS_XIRR_PROP, ANALYTICS_BENCH_PROP — benchmark.js
 *              DISCIPLINE_PROP — discipline.js
 *              HEALTH_PROP — health.js
 *              IIS_PROP, IIS_DEDUCTION_LIMIT — iis.js
 *              LDV_PROP — ldv.js
 */

function generateAnnualReport() {
  let ui = SpreadsheetApp.getUi();
  let props = PropertiesService.getScriptProperties();

  let config;
  try { config = readConfig_(); }
  catch (e) { ui.alert('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);

  let toDate   = new Date();
  let fromDate = new Date(toDate.getFullYear() - 1, toDate.getMonth(), toDate.getDate());

  let allOps = [];
  getAccounts_().forEach(function(acc) {
    allOps = allOps.concat(fetchOperations_(acc.id, acc.name, fromDate, toDate));
  });
  let summary = buildSummary_(allOps);

  // Отделяем автоматические операции овернайта (денежный рынок, ВИМ-фонды,
  // ликвидность, валютные облигации) от реальных осознанных сделок
  let AUTO_KEYWORDS = ['денежный рынок', 'вим', 'ликвидность', 'локальные валютные облигации'];
  let realBuys = 0, realSells = 0, autoBuys = 0, autoSells = 0;

  allOps.forEach(function(op) {
    let isBuy  = op.type === 'OPERATION_TYPE_BUY'  || op.type === 'OPERATION_TYPE_BUY_CARD';
    let isSell = op.type === 'OPERATION_TYPE_SELL' || op.type === 'OPERATION_TYPE_SELL_CARD';
    if (!isBuy && !isSell) return;

    let nameLower = (op.instrumentName || '').toLowerCase();
    let isAuto = AUTO_KEYWORDS.some(function(kw) { return nameLower.indexOf(kw) >= 0; });

    if (isAuto) { if (isBuy) autoBuys++; else autoSells++; }
    else        { if (isBuy) realBuys++; else realSells++; }
  });

  let doc = DocumentApp.create('Годовой отчёт по портфелю — ' +
    Utilities.formatDate(toDate, Session.getScriptTimeZone(), 'dd.MM.yyyy'));
  let body = doc.getBody();

  body.appendParagraph('Годовой отчёт по инвестиционному портфелю')
    .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Сформирован: ' +
    Utilities.formatDate(toDate, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') +
    '   ·   Период: последние 12 месяцев')
    .setForegroundColor('#666666');

  // ── Итоговая стоимость ──────────────────────────────────────────
  body.appendParagraph('Портфель сейчас').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Общая стоимость: ' + rub_(totalRub)).setBold(false);

  // ── Доходность (последние сохранённые значения) ──────────────────
  body.appendParagraph('Доходность').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  let xirrRaw = props.getProperty('ANALYTICS_XIRR_VALUE');
  if (xirrRaw) {
    let xirr = JSON.parse(xirrRaw);
    body.appendParagraph('XIRR (доходность с учётом дат пополнений): ' +
      (xirr.value !== null ? xirr.value.toFixed(1) + '% годовых' : 'н/д'));
  } else {
    body.appendParagraph('XIRR не рассчитан — запусти «Рассчитать XIRR» перед формированием отчёта.');
  }

  let benchRaw = props.getProperty('ANALYTICS_BENCH_VALUE');
  if (benchRaw) {
    let b = JSON.parse(benchRaw);
    let diff = b.actual - b.hypothetical;
    body.appendParagraph('Сравнение с IMOEX: ваш портфель ' + rub_(b.actual) +
      ', при покупке индекса на те же суммы было бы ' + rub_(b.hypothetical) +
      ' (' + (diff >= 0 ? 'обгон' : 'отставание') + ' на ' + rub_(Math.abs(diff)) + ')');
  }

  // ── Операции за год ────────────────────────────────────────────
  body.appendParagraph('Активность за год').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  let opsTable = body.appendTable([
    ['Показатель', 'Значение'],
    ['Пополнено', rub_(Math.round(summary.totalIn))],
    ['Выведено', rub_(Math.round(summary.totalOut))],
    ['Получено купонов', rub_(Math.round(summary.totalCoupons))],
    ['Получено дивидендов', rub_(Math.round(summary.totalDivs))],
    ['Покупок (акции + облигации)', String(realBuys)],
    ['Продаж (акции + облигации)', String(realSells)],
    ['Операций денежного рынка/овернайта', String(autoBuys + autoSells) + ' (автоматические, не торговые решения)'],
  ]);
  opsTable.getRow(0).editAsText().setBold(true);

  // ── Дисциплина пополнений ─────────────────────────────────────
  let discRaw = props.getProperty('DISCIPLINE_DATA');
  if (discRaw) {
    let d = JSON.parse(discRaw);
    body.appendParagraph('Дисциплина пополнений').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Месяцев с чистым пополнением: ' + d.positiveCount + ' из ' + d.months.length);
    body.appendParagraph('Среднее сальдо в активный месяц: ' + rub_(d.avgNet));
  }

  // ── Концентрация ─────────────────────────────────────────────
  let healthRaw = props.getProperty('HEALTH_CONCENTRATION_DATA');
  if (healthRaw) {
    let h = JSON.parse(healthRaw);
    body.appendParagraph('Концентрация портфеля').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if (h.overall.top1Name) {
      body.appendParagraph('Крупнейшая позиция: ' + h.overall.top1Name +
        ' — ' + (h.overall.top1Pct * 100).toFixed(1) + '% от портфеля');
      body.appendParagraph('Топ-3 позиции: ' + (h.overall.top3Pct * 100).toFixed(1) + '% от портфеля');
    }
  }

  // ── ИИС-3 — вычет за год ─────────────────────────────────────
  let iisRaw = props.getProperty(IIS_PROP);
  if (iisRaw) {
    let iis = JSON.parse(iisRaw);
    body.appendParagraph('ИИС-3 — вычет за год').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Счёт: ' + iis.accountName + ' (' + iis.year + ' год)');
    body.appendParagraph('Использовано вычета: ' + (iis.pct * 100).toFixed(1) + '% — ' +
      rub_(iis.contributed) + ' из ' + rub_(IIS_DEDUCTION_LIMIT));
    body.appendParagraph(iis.remaining > 0
      ? 'Осталось довнести до конца года для максимального вычета: ' + rub_(iis.remaining)
      : '✅ Лимит вычета за этот год полностью выбран');
  }

  // ── ЛДВ — льгота на долгосрочное владение ────────────────────
  let ldvRaw = props.getProperty(LDV_PROP);
  if (ldvRaw) {
    let ldvResults = JSON.parse(ldvRaw);
    if (ldvResults.length) {
      body.appendParagraph('Льгота на долгосрочное владение (ЛДВ)').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      let tz = Session.getScriptTimeZone();
      let ldvRows = [['Акция', 'Есть льгота, шт.', 'Пока нет льготы, шт.', 'Ближайший льготный транш']];
      ldvResults.forEach(function(r) {
        let dateStr = r.notYetQty > 0 && r.nextEligibleDate
          ? Utilities.formatDate(new Date(r.nextEligibleDate), tz, 'dd.MM.yyyy')
          : '—';
        ldvRows.push([r.name, String(r.eligibleQty), String(r.notYetQty), dateStr]);
      });
      let ldvTable = body.appendTable(ldvRows);
      ldvTable.getRow(0).editAsText().setBold(true);
    }
  }

  // ── Распределение по классам ────────────────────────────────
  body.appendParagraph('Распределение по классам активов').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  let cats = Object.keys(config.classTargets);
  let allocRows = [['Категория', 'Сумма', 'Текущий %', 'Цель %']];
  cats.forEach(function(cat) {
    let actual = positions.filter(function(p) { return p.category === cat; })
                          .reduce(function(s, p) { return s + p.valueRub; }, 0);
    let actPct = totalRub > 0 ? actual / totalRub : 0;
    let tgtPct = config.classTargets[cat] || 0;
    allocRows.push([cat, rub_(actual), (actPct * 100).toFixed(1) + '%', (tgtPct * 100).toFixed(1) + '%']);
  });
  let allocTable = body.appendTable(allocRows);
  allocTable.getRow(0).editAsText().setBold(true);

  // ── Футер ────────────────────────────────────────────────────
  body.appendParagraph('');
  let footer = body.appendParagraph(
    'Сформировано с помощью t-invest-portfolio-tracker — github.com/Relff/t-invest-portfolio-tracker');
  footer.setForegroundColor('#999999');
  footer.editAsText().setFontSize(0, footer.getText().length - 1, 9);

  doc.saveAndClose();

  let url = doc.getUrl();
  let html = HtmlService.createHtmlOutput(
    '<p>Отчёт готов: <a href="' + url + '" target="_blank">открыть документ</a></p>'
  ).setWidth(350).setHeight(80);
  ui.showModalDialog(html, 'Годовой отчёт сформирован');
}