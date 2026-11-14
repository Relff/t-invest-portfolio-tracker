/**
 * advparams.js — Продвинутые параметры (пороги 5/25, health check, FIFO-окно)
 *
 * Читает блок «▌ ПРОДВИНУТЫЕ ПАРАМЕТРЫ» из Config по подписям строк.
 * Если блока нет, строка не найдена, или значение некорректно (не число,
 * ≤ 0) — используется дефолт из кода, ничего не падает и не ломается.
 */

const ADV_PARAMS_DEFAULTS_ = {
  thrAbs: 5.0, thrRel: 0.25,
  healthTop1Warn: 15, healthTop1Crit: 25,
  healthTop3Warn: 35, healthTop3Crit: 50,
  fifoYears: 5,
  inflationPct: 6.0, // на конец июля 2026 годовая инфляция в РФ ~5.9% (ЦБ) — обнови при желании
};

const ADV_PARAMS_LABELS_ = {
  'Правило 5/25 — абсолютное (п.п.)':      'thrAbs',
  'Правило 5/25 — относительное':          'thrRel',
  'Health: топ-1 предупреждение (%)':      'healthTop1Warn',
  'Health: топ-1 критично (%)':            'healthTop1Crit',
  'Health: топ-3 предупреждение (%)':      'healthTop3Warn',
  'Health: топ-3 критично (%)':            'healthTop3Crit',
  'FIFO — глубина истории (лет)':          'fifoYears',
  'Инфляция, % годовых':                   'inflationPct',
};

function readAdvancedParams_() {
  let result = Object.assign({}, ADV_PARAMS_DEFAULTS_);
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) return result;

  let v = sh.getDataRange().getValues();
  for (let i = 0; i < v.length; i++) {
    let label = String(v[i][0]).trim();
    let key = ADV_PARAMS_LABELS_[label];
    if (!key) continue;
    let raw = Number(v[i][1]);
    if (!isNaN(raw) && raw > 0) result[key] = raw;
  }
  return result;
}

function addAdvancedParamsBlock() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) { SpreadsheetApp.getUi().alert('⚠️ Сначала запустите initConfig()'); return; }

  let vals = sh.getDataRange().getValues();
  let blockHeaderRow = -1, lastParamRow = -1, hasInflation = false;
  for (let i = 0; i < vals.length; i++) {
    let label = String(vals[i][0]).trim();
    if (label.indexOf('ПРОДВИНУТЫЕ ПАРАМЕТРЫ') >= 0) blockHeaderRow = i;
    if (blockHeaderRow !== -1 && ADV_PARAMS_LABELS_[label]) lastParamRow = i; // последняя строка-параметр внутри блока
    if (label === 'Инфляция, % годовых') hasInflation = true;
  }

  if (blockHeaderRow !== -1) {
    if (hasInflation) {
      SpreadsheetApp.getUi().alert('Блок «Продвинутые параметры» уже содержит поле «Инфляция, % годовых» — добавлять нечего.');
      return;
    }
    // Блок уже есть (с прошлой версии) — дописываем только новую строку
    // сразу после последнего существующего параметра, остальное не трогаем.
    let insertAfterRow = lastParamRow + 1; // 0-indexed → 1-indexed строка ПОСЛЕ
    sh.insertRowAfter(insertAfterRow);
    let newRow = insertAfterRow + 1;
    sh.getRange(newRow, 1, 1, 3).setValues([[
      'Инфляция, % годовых', ADV_PARAMS_DEFAULTS_.inflationPct,
      'Для расчёта реальной доходности (XIRR с поправкой на инфляцию) — обновляй периодически по данным ЦБ РФ'
    ]]);
    SpreadsheetApp.getUi().alert('✅ Добавил поле «Инфляция, % годовых» (по умолчанию ' +
      ADV_PARAMS_DEFAULTS_.inflationPct + '%) в существующий блок «Продвинутые параметры».');
    return;
  }

  let lastRow = sh.getLastRow() + 2;
  let d = ADV_PARAMS_DEFAULTS_;
  let block = [
    ['▌ ПРОДВИНУТЫЕ ПАРАМЕТРЫ', '', ''],
    ['(можно не трогать — стоят разумные значения по умолчанию)', '', ''],
    ['Параметр', 'Значение', 'Пояснение'],
    ['Правило 5/25 — абсолютное (п.п.)',      d.thrAbs,          'Критично, если отклонение ≥ N процентных пунктов'],
    ['Правило 5/25 — относительное',          d.thrRel,          'Критично, если отклонение ≥ 25% от целевой доли'],
    ['Health: топ-1 предупреждение (%)',      d.healthTop1Warn,  'Жёлтый статус, если одна позиция ≥ N%'],
    ['Health: топ-1 критично (%)',            d.healthTop1Crit,  'Красный статус, если одна позиция ≥ N%'],
    ['Health: топ-3 предупреждение (%)',      d.healthTop3Warn,  ''],
    ['Health: топ-3 критично (%)',            d.healthTop3Crit,  ''],
    ['FIFO — глубина истории (лет)',          d.fifoYears,       'Сколько лет назад искать сделки для средней цены'],
    ['Инфляция, % годовых',                   d.inflationPct,    'Для расчёта реальной доходности (XIRR с поправкой на инфляцию) — обновляй периодически по данным ЦБ РФ'],
  ];

  sh.getRange(lastRow, 1, block.length, 3).setValues(block);
  sh.getRange(lastRow, 1, 1, 3).merge()
    .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.getRange(lastRow + 1, 1, 1, 3).merge()
    .setFontColor('#9e9e9e').setFontStyle('italic');
  sh.getRange(lastRow + 2, 1, 1, 3)
    .setBackground(C.MID).setFontColor('#ffffff').setFontWeight('bold');

  SpreadsheetApp.getUi().alert('✅ Блок «Продвинутые параметры» добавлен.');
}