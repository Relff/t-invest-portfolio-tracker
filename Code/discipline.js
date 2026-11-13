/**
 * discipline.js — Дисциплина пополнений (месячное сальдо)
 *
 * Исправлено отображение месяцев: значения записываются как текст,
 * чтобы Google Sheets не превращал их в даты.
 */

const DISCIPLINE_PROP          = 'DISCIPLINE_DATA';
const DISCIPLINE_SECTION_TITLE = '▌ ДИСЦИПЛИНА ПОПОЛНЕНИЙ (месячное сальдо)';
const DISCIPLINE_MONTHS_BACK   = 12;
/* Produced by Relferium — t-invest-portfolio-tracker */

function calculateContributionDiscipline() {
  let toDate   = new Date();
  let fromDate = new Date(toDate.getFullYear(), toDate.getMonth() - DISCIPLINE_MONTHS_BACK + 1, 1);

  let netByMonth = {};

  getAccounts_().forEach(function(acc) {
    let ops = fetchOperations_(acc.id, acc.name, fromDate, toDate);
    ops.forEach(function(op) {
      if (op.type !== 'OPERATION_TYPE_INPUT' && op.type !== 'OPERATION_TYPE_OUTPUT') return;
      let key = op.date.getFullYear() + '-' + op.date.getMonth();
      if (!netByMonth[key]) netByMonth[key] = { in: 0, out: 0 };
      if (op.type === 'OPERATION_TYPE_INPUT')  netByMonth[key].in  += op.amount;
      if (op.type === 'OPERATION_TYPE_OUTPUT') netByMonth[key].out += op.amount;
    });
  });

  let months = [];
  for (let i = DISCIPLINE_MONTHS_BACK - 1; i >= 0; i--) {
    let d = new Date(toDate.getFullYear(), toDate.getMonth() - i, 1);
    let key = d.getFullYear() + '-' + d.getMonth();
    let grid = netByMonth[key] || { in: 0, out: 0 };
    months.push({
      label: d,
      inSum:  grid.in,
      outSum: grid.out,
      net:    grid.in - grid.out,
    });
  }

  let activeMonths = months.filter(function(m) { return m.net !== 0; });
  let positiveMonths = months.filter(function(m) { return m.net > 0; });
  let avgNet = activeMonths.length
    ? Math.round(activeMonths.reduce(function(s, m) { return s + m.net; }, 0) / activeMonths.length)
    : 0;

  let lastPositive = null;
  for (let i = months.length - 1; i >= 0; i--) {
    if (months[i].net > 0) { lastPositive = months[i].label; break; }
  }

  PropertiesService.getScriptProperties().setProperty(DISCIPLINE_PROP, JSON.stringify({
    months: months.map(function(m) {
      return { y: m.label.getFullYear(), mo: m.label.getMonth(), net: Math.round(m.net) };
    }),
    positiveCount: positiveMonths.length,
    avgNet: avgNet,
    lastPositiveY:  lastPositive ? lastPositive.getFullYear() : null,
    lastPositiveMo: lastPositive ? lastPositive.getMonth()   : null,
  }));

  redrawDisciplineSection_();
}

function redrawDisciplineSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(DISCIPLINE_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);
  let MONTH_NAMES = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];

  renderSection_(sh, DISCIPLINE_SECTION_TITLE, function(sh, r, COLS, colStart) {
    sh.getRange(r, colStart, 1, 3).merge().setValue('Месяцев с чистым пополнением').setFontWeight('bold');
    sh.getRange(r, colStart + 3, 1, COLS - 3).merge().setValue(d.positiveCount + ' из ' + d.months.length)
      .setFontWeight('bold').setHorizontalAlignment('right')
      .setFontColor(d.positiveCount >= 10 ? C.OK : d.positiveCount >= 6 ? C.WARN : C.CRIT);
    sh.getRange(r, colStart, 1, COLS).setBackground(C.EVEN);
    r++;

    sh.getRange(r, colStart, 1, 3).merge().setValue('Среднее сальдо в активный месяц').setFontWeight('bold');
    sh.getRange(r, colStart + 3, 1, COLS - 3).merge().setValue(rub_(d.avgNet))
      .setFontWeight('bold').setHorizontalAlignment('right');
    sh.getRange(r, colStart, 1, COLS).setBackground(C.EVEN);
    r++; r++;

    sh.getRange(r, colStart).setValue('Месяц').setFontWeight('bold');
    sh.getRange(r, colStart + 3, 1, COLS - 3).merge().setValue('Чистое сальдо, ₽').setFontWeight('bold').setHorizontalAlignment('right');
    sh.getRange(r, colStart, 1, COLS).setBackground(C.MID).setFontColor('#ffffff');
    r++;

    d.months.forEach(function(m, idx) {
      let bg = idx % 2 === 0 ? C.EVEN : C.ODD;
      let clr = m.net > 0 ? C.OK : m.net < 0 ? C.CRIT : C.SKIP;
      let monthCell = sh.getRange(r, colStart);

      // Записываем как текст, чтобы Sheets не превращал в дату
      monthCell
        .setNumberFormat('@')
        .setValue(MONTH_NAMES[m.mo] + ' ' + m.y)
        .setBackground(bg)
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle')
        .setFontSize(10)
        .setFontWeight('normal')
        .setFontStyle('normal')
        .setFontColor('#000000');

      sh.getRange(r, colStart + 3, 1, COLS - 3)
        .merge()
        .setValue(rub_(m.net))
        .setBackground(bg)
        .setFontColor(clr)
        .setFontWeight('bold')
        .setFontSize(10)
        .setFontStyle('normal')
        .setHorizontalAlignment('right')
        .setVerticalAlignment('middle');

      sh.getRange(r, colStart, 1, COLS).setBackground(bg);
      r++;
    });
  }, 'right');
}
