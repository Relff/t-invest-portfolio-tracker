/**
 * xirr.js — Расчёт XIRR и простого результата за выбранный период
 *
 * Зависимости:
 *   getAccounts_(), fetchOperations_() — history.js
 *   readConfig_(), readPositions_() — dashboard.js
 *   redrawAnalyticsSection_(), ANALYTICS_XIRR_PROP — benchmark.js
 */

const XIRR_HELPER_SHEET = '_XIRR_calc';
const XIRR_YEARS_BACK   = 5;


/**
 * Читает параметр:
 *   XIRR — считать с даты | 22.06.2026
 *
 * Если строка отсутствует или значение пустое — возвращает null,
 * и используется старое окно XIRR_YEARS_BACK.
 */
function readXirrStartDate_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);

  if (!sh) return null;

  let values = sh.getDataRange().getValues();

  for (let i = 0; i < values.length; i++) {
    let label = String(values[i][0] || '').trim();

    if (label !== 'XIRR — считать с даты') {
      continue;
    }

    let raw = values[i][1];

    // Настоящая дата Google Sheets
    if (raw instanceof Date && !isNaN(raw.getTime())) {
      return raw;
    }

    // Текстовая дата ДД.ММ.ГГГГ
    let text = String(raw || '').trim();
    if (!text) return null;

    let match = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);

    if (!match) {
      throw new Error(
        'Некорректная дата XIRR: "' +
        text +
        '". Используйте формат ДД.ММ.ГГГГ.'
      );
    }

    let day = Number(match[1]);
    let month = Number(match[2]) - 1;
    let year = Number(match[3]);

    let date = new Date(year, month, day);

    if (
      date.getFullYear() !== year ||
      date.getMonth() !== month ||
      date.getDate() !== day
    ) {
      throw new Error('Некорректная дата XIRR: ' + text);
    }

    return date;
  }

  return null;
}


/**
 * Полный расчёт статистики периода.
 *
 * Возвращает:
 * {
 *   xirrPct,
 *   currentValue,
 *   totalInputs,
 *   totalOutputs,
 *   profitRub,
 *   simpleReturnPct,
 *   startDate
 * }
 *
 * simpleReturnPct — НЕ годовая доходность:
 *
 *   (текущая стоимость + выводы - пополнения) / пополнения
 *
 * Это удобно как простой ответ на вопрос:
 * «Сколько денег и процентов я фактически заработал за этот период?»
 */
function computeXirrStats_() {
  let ss = SpreadsheetApp.getActive();

  let toDate = new Date();

  let configuredStartDate = readXirrStartDate_();

  let fromDate = configuredStartDate
    ? configuredStartDate
    : new Date(
        toDate.getTime() -
        XIRR_YEARS_BACK * 365 * 24 * 3600 * 1000
      );

  if (fromDate >= toDate) {
    throw new Error(
      'Дата начала XIRR должна быть раньше сегодняшней даты.'
    );
  }

  let flows = [];
  let totalInputs = 0;
  let totalOutputs = 0;

  getAccounts_().forEach(function(acc) {
    let ops = fetchOperations_(
      acc.id,
      acc.name,
      fromDate,
      toDate
    );

    ops.forEach(function(op) {
      if (op.type === 'OPERATION_TYPE_INPUT') {
        let amount = Math.abs(op.amount);

        flows.push({
          date: op.date,
          amount: -amount
        });

        totalInputs += amount;
      }

      if (op.type === 'OPERATION_TYPE_OUTPUT') {
        let amount = Math.abs(op.amount);

        flows.push({
          date: op.date,
          amount: amount
        });

        totalOutputs += amount;
      }
    });
  });

  let config = readConfig_();

  // Только инвестиционный портфель.
  // Физические USD/EUR сюда не входят.
  let positions = readPositions_(config)
    .filter(function(p) {
      return p.source !== 'Manual';
    });

  let totalRub = positions.reduce(function(s, p) {
    return s + p.valueRub;
  }, 0);

  let profitRub =
    totalRub +
    totalOutputs -
    totalInputs;

  let simpleReturnPct =
    totalInputs > 0
      ? profitRub / totalInputs * 100
      : null;

  // Без хотя бы одного внешнего пополнения XIRR считать бессмысленно.
  if (flows.length < 1) {
    return {
      xirrPct: null,
      currentValue: totalRub,
      totalInputs: totalInputs,
      totalOutputs: totalOutputs,
      profitRub: profitRub,
      simpleReturnPct: simpleReturnPct,
      startDate: fromDate
    };
  }

  flows.push({
    date: toDate,
    amount: totalRub
  });

  flows.sort(function(a, b) {
    return a.date - b.date;
  });

  let sh = ss.getSheetByName(XIRR_HELPER_SHEET);

  if (!sh) {
    sh = ss.insertSheet(XIRR_HELPER_SHEET);
  }

  sh.clearContents();

  sh.getRange(1, 1, flows.length, 1)
    .setValues(
      flows.map(function(f) {
        return [f.amount];
      })
    );

  sh.getRange(1, 2, flows.length, 1)
    .setValues(
      flows.map(function(f) {
        return [f.date];
      })
    );

  sh.getRange(1, 2, flows.length, 1)
    .setNumberFormat('dd.mm.yyyy');

  let formulaCell = sh.getRange(1, 4);

  formulaCell.setFormula(
    '=XIRR(A1:A' +
    flows.length +
    '; B1:B' +
    flows.length +
    ')'
  );

  SpreadsheetApp.flush();

  let xirrValue = formulaCell.getValue();

  sh.hideSheet();

  let xirrPct =
    typeof xirrValue === 'number'
      ? xirrValue * 100
      : null;

  return {
    xirrPct: xirrPct,
    currentValue: totalRub,
    totalInputs: totalInputs,
    totalOutputs: totalOutputs,
    profitRub: profitRub,
    simpleReturnPct: simpleReturnPct,
    startDate: fromDate
  };
}


/**
 * Совместимость с остальными модулями (например snapshot.js).
 * Возвращает только XIRR в % годовых.
 */
function computeXirrValue_() {
  return computeXirrStats_().xirrPct;
}


function calculateXIRR() {
  let stats;

  try {
    stats = computeXirrStats_();
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Ошибка расчёта XIRR: ' + e.message
    );
    return;
  }

  if (stats.xirrPct === null) {
    SpreadsheetApp.getUi().alert(
      'Недостаточно данных: после выбранной даты нет ни одного пополнения, либо XIRR не удалось посчитать.'
    );
  }

  writeXirrToDashboard_(stats);
}


/**
 * Сохраняем не только XIRR, но и простой фактический результат периода.
 */
function writeXirrToDashboard_(stats) {
  let tz = Session.getScriptTimeZone();

  // Обратная совместимость на случай вызова со старым числом.
  if (typeof stats === 'number' || stats === null) {
    stats = {
      xirrPct: typeof stats === 'number' ? stats : null,
      currentValue: null,
      totalInputs: null,
      totalOutputs: null,
      profitRub: null,
      simpleReturnPct: null,
      startDate: null
    };
  }

  let startDateStr =
    stats.startDate instanceof Date &&
    !isNaN(stats.startDate.getTime())
      ? Utilities.formatDate(
          stats.startDate,
          tz,
          'dd.MM.yyyy'
        )
      : '';

  PropertiesService
    .getScriptProperties()
    .setProperty(
      ANALYTICS_XIRR_PROP,
      JSON.stringify({
        value: stats.xirrPct,
        startDate: startDateStr,
        currentValue: stats.currentValue,
        totalInputs: stats.totalInputs,
        totalOutputs: stats.totalOutputs,
        profitRub: stats.profitRub,
        simpleReturnPct: stats.simpleReturnPct
      })
    );
}
