/**
 * goal.js — Прогресс к цели (целевая сумма портфеля) + прогноз даты
 * достижения по текущему темпу (XIRR + средняя дисциплина пополнений)
 *
 * Зависимости: readConfig_(), readPositions_(), rub_(), DST, C — dashboard.js
 *              renderSection_() — sections.js
 *              ANALYTICS_XIRR_VALUE — xirr.js
 *              DISCIPLINE_DATA — discipline.js
 */

const GOAL_PROP          = 'GOAL_PROGRESS_DATA';
const GOAL_SECTION_TITLE = '▌ ПРОГРЕСС К ЦЕЛИ';

/**
 * Прогноз даты достижения цели — симуляция месяц за месяцем: текущий
 * капитал растёт по ставке XIRR (уже посчитанной в xirr.js), плюс каждый
 * месяц добавляется среднее пополнение (уже посчитанное в discipline.js).
 * Ничего заново не запрашивает у API — только то, что уже посчитано.
 *
 * Это ПРОСТАЯ линейная проекция, не Monte Carlo и не учёт волатильности —
 * честная оценка «при сохранении текущего темпа», не гарантия.
 *
 * Возвращает null, если данных недостаточно (XIRR или дисциплина ещё
 * не считались) или если цель недостижима за разумный горизонт (50 лет)
 * при нулевом/отрицательном росте и нулевых пополнениях.
 */
function projectGoalDate_(totalRub, targetSum) {
  let props   = PropertiesService.getScriptProperties();
  let xirrRaw = props.getProperty('ANALYTICS_XIRR_VALUE');
  let discRaw = props.getProperty('DISCIPLINE_DATA');
  if (!xirrRaw || !discRaw) return null;

  let xirr = JSON.parse(xirrRaw).value; // % годовых, может быть null
  let disc = JSON.parse(discRaw);
  let monthlyContribution = Math.max(0, disc.avgNet || 0); // отрицательное пополнение не считаем

  let monthlyRate = (xirr === null ? 0 : xirr / 100) / 12;
  if (monthlyRate <= 0 && monthlyContribution <= 0) return null; // ни роста, ни пополнений — недостижимо

  let balance = totalRub;
  let months  = 0;
  let MAX_MONTHS = 600; // 50 лет — дальше прогноз уже бессмысленен
  while (balance < targetSum && months < MAX_MONTHS) {
    balance = balance * (1 + monthlyRate) + monthlyContribution;
    months++;
  }
  if (balance < targetSum) return null; // не достигается даже за 50 лет

  let now = new Date();
  let targetDate = new Date(now.getFullYear(), now.getMonth() + months, now.getDate());

  return {
    months: months,
    years: Math.floor(months / 12),
    remMonths: months % 12,
    date: targetDate.getTime(),
    xirrUsed: xirr,
    monthlyContribution: Math.round(monthlyContribution),
  };
}

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
  let projection = pct >= 1 ? null : projectGoalDate_(totalRub, targetSum);

  PropertiesService.getScriptProperties().setProperty(GOAL_PROP, JSON.stringify({
    total: Math.round(totalRub), target: Math.round(targetSum), pct: pct,
    projection: projection,
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

  renderSection_(sh, GOAL_SECTION_TITLE, function(sh, r, COLS, colStart) {
    let barLen  = 20;
    let filled  = Math.round(d.pct * barLen);
    let bar     = '█'.repeat(filled) + '░'.repeat(barLen - filled);
    let clr     = d.pct >= 1 ? C.OK : d.pct >= 0.5 ? C.WARN : C.CRIT;

    sh.getRange(r, colStart, 1, 3).merge().setValue('Прогресс к цели (' + rub_(d.target) + ')').setFontWeight('bold');
    sh.getRange(r, colStart + 3, 1, COLS - 3).merge().setValue((d.pct * 100).toFixed(1) + '%')
      .setFontColor(clr).setFontWeight('bold').setHorizontalAlignment('right');
    sh.getRange(r, colStart, 1, COLS).setBackground(C.EVEN);
    r++;

    sh.getRange(r, colStart, 1, COLS).merge()
      .setValue(bar + '  ' + rub_(d.total) + ' из ' + rub_(d.target))
      .setFontFamily('Courier New').setFontSize(10).setFontColor(clr).setBackground(C.EVEN);
    r++;

    if (d.pct >= 1) {
      sh.getRange(r, colStart, 1, COLS).merge()
        .setValue('✅ Цель уже достигнута')
        .setFontColor(C.OK).setFontWeight('bold').setFontSize(10).setBackground(C.EVEN);
    } else if (d.projection) {
      let p = d.projection;
      let tz = Session.getScriptTimeZone();
      let dateStr = Utilities.formatDate(new Date(p.date), tz, 'dd.MM.yyyy');
      let xirrStr = p.xirrUsed !== null ? p.xirrUsed.toFixed(1) + '%' : '0%';
      sh.getRange(r, colStart, 1, COLS).merge()
        .setValue('📅 При текущем темпе (XIRR ' + xirrStr + ', пополнение ' + rub_(p.monthlyContribution) +
                  '/мес) — цель к ' + dateStr + ' (~' + p.years + ' г. ' + p.remMonths + ' мес.)')
        .setFontColor('#9e9e9e').setFontStyle('italic').setFontSize(9).setBackground(C.EVEN);
    } else {
      sh.getRange(r, colStart, 1, COLS).merge()
        .setValue('ℹ️ Для прогноза даты нужны посчитанные XIRR и дисциплина пополнений (или темп сейчас нулевой/отрицательный)')
        .setFontColor('#9e9e9e').setFontStyle('italic').setFontSize(9).setBackground(C.EVEN);
    }
  }, 'left');
}