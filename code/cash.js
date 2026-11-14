/**
 * cash.js — журнал и сводка ручных денежных активов
 *
 * Лист: 💵 Наличные
 *
 * ЖУРНАЛ (ручной ввод, НЕ перезаписывается скриптом):
 * A Дата
 * B Операция      Покупка / Продажа / Пополнение / Снятие / Проценты
 * C Актив         USD / EUR / Вклад RUB / Сейф RUB / ...
 * D Количество
 * E Курс сделки ₽
 * F Сумма ₽
 * G Комментарий
 *
 * СВОДКА (строится автоматически, I:R):
 * I  Актив
 * J  Текущий курс
 * K  Кол-во
 * L  На сумму
 * M  % от наличных
 * N  Средний курс покупки
 * O  P/L ₽
 * P  P/L %
 * Q  Реализ. P/L
 * R  Проценты / доход
 *
 * ВАЖНО:
 * updateCashSummary_() никогда не очищает и не перезаписывает A:G.
 * При обновлении пересобирается только автоматическая сводка справа.
 * Ручные текущие курсы из J сохраняются между обновлениями.
 */

const CASH_SHEET_ = '💵 Наличные';
const CASH_SUMMARY_COL_ = 9;   // I
const CASH_SUMMARY_COLS_ = 10; // I:R
const CASH_EPS_ = 0.0000001;


/**
 * Определяет базовую валюту по названию актива.
 *
 * Примеры:
 * USD        -> USD
 * EUR        -> EUR
 * Вклад RUB  -> RUB
 * Сейф RUB   -> RUB
 */
function cashBaseCurrency_(asset) {
  let s = String(asset || '').trim().toUpperCase();
  if (!s) return '';

  let m = s.match(/(?:^|\s)([A-Z]{3})$/);
  return m ? m[1] : s;
}


/**
 * Считывает текущие курсы из существующей автоматической сводки.
 *
 * Это позволяет пользователю вручную менять J («Текущий курс»):
 * при следующем updateCashSummary_ значение не потеряется.
 */
function readCashCurrentRates_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CASH_SHEET_);
  let result = {
    byAsset: {},
    byCurrency: {}
  };

  if (!sh || sh.getLastRow() < 2) return result;

  let rows = Math.max(0, sh.getLastRow() - 1);
  if (!rows) return result;

  let data = sh.getRange(2, CASH_SUMMARY_COL_, rows, 2).getValues();

  data.forEach(function(row) {
    let asset = String(row[0] || '').trim();
    let rate = Number(row[1] || 0);

    if (!asset || rate <= 0) return;

    result.byAsset[asset.toUpperCase()] = rate;

    let base = cashBaseCurrency_(asset);
    if (base && !result.byCurrency[base]) {
      result.byCurrency[base] = rate;
    }
  });

  return result;
}


/**
 * Преобразует строку журнала в количество и рублёвую сумму.
 * Можно заполнять и D/E, и F.
 *
 * Для RUB-активов:
 * - если D пусто, а F заполнено, количество = F;
 * - курс автоматически 1.
 */
function normalizeCashOperation_(asset, qtyRaw, rateRaw, amountRaw) {
  let base = cashBaseCurrency_(asset);

  let qty = Number(qtyRaw || 0);
  let rate = Number(rateRaw || 0);
  let amount = Number(amountRaw || 0);

  if (base === 'RUB') {
    if (rate <= 0) rate = 1;
    if (qty <= 0 && amount > 0) qty = amount;
    if (amount <= 0 && qty > 0) amount = qty;
  } else {
    if (amount <= 0 && qty > 0 && rate > 0) {
      amount = qty * rate;
    }

    if (qty <= 0 && amount > 0 && rate > 0) {
      qty = amount / rate;
    }
  }

  return {
    base: base,
    qty: qty,
    rate: rate,
    amount: amount
  };
}


/**
 * Разбирает журнал A:G.
 *
 * Поддерживаемые операции:
 *   Покупка     + остаток, увеличивает себестоимость
 *   Продажа     - остаток, считает реализованный P/L
 *   Пополнение  + остаток, увеличивает себестоимость
 *   Снятие      - остаток, НЕ считается продажей
 *   Проценты    + остаток, НЕ увеличивает себестоимость,
 *                отдельно накапливается как доход
 */
function readCashLedger_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CASH_SHEET_);

  if (!sh || sh.getLastRow() < 2) {
    return {};
  }

  let data = sh
    .getRange(2, 1, sh.getLastRow() - 1, 7)
    .getValues();

  let assets = {};

  function getAsset_(name) {
    let key = String(name || '').trim();

    if (!assets[key]) {
      assets[key] = {
        asset: key,
        baseCurrency: cashBaseCurrency_(key),
        qty: 0,
        costBasisRub: 0,
        realizedPL: 0,
        incomeRub: 0,
        lastDealRate: 0
      };
    }

    return assets[key];
  }

  data.forEach(function(row, index) {
    let operation = String(row[1] || '').trim().toLowerCase();
    let asset = String(row[2] || '').trim();

    if (!operation || !asset) return;

    let normalized = normalizeCashOperation_(
      asset,
      row[3],
      row[4],
      row[5]
    );

    let qty = normalized.qty;
    let rate = normalized.rate;
    let amountRub = normalized.amount;

    if (qty <= 0) return;

    let a = getAsset_(asset);

    if (rate > 0) {
      a.lastDealRate = rate;
    }

    // Покупка / обычное пополнение капитала
    if (operation === 'покупка' || operation === 'пополнение') {
      let cost = amountRub > 0
        ? amountRub
        : (rate > 0 ? qty * rate : 0);

      // Для нерублёвого актива желательно иметь курс сделки.
      // Если его нет, остаток всё равно учитываем, но cost basis
      // может остаться неполным до ручного заполнения курса/суммы.
      a.qty += qty;
      a.costBasisRub += Math.max(0, cost);
      return;
    }

    // Проценты — это уже заработанный доход.
    if (operation === 'проценты') {
      let income = amountRub > 0
        ? amountRub
        : (rate > 0 ? qty * rate : 0);

      a.qty += qty;
      a.incomeRub += Math.max(0, income);
      return;
    }

    // Продажа / снятие уменьшают остаток.
    if (operation === 'продажа' || operation === 'снятие') {
      if (qty > a.qty + CASH_EPS_) {
        throw new Error(
          'Наличные, строка ' +
          (index + 2) +
          ': попытка списать ' +
          qty +
          ' ' +
          asset +
          ', но доступно только ' +
          a.qty
        );
      }

      let avgCostRate =
        a.qty > CASH_EPS_
          ? a.costBasisRub / a.qty
          : 0;

      // Продажа фиксирует P/L.
      if (operation === 'продажа') {
        let proceeds = amountRub > 0
          ? amountRub
          : (rate > 0 ? qty * rate : 0);

        if (proceeds > 0 && avgCostRate > 0) {
          a.realizedPL += proceeds - avgCostRate * qty;
        }
      }

      // И продажа, и снятие убирают пропорциональную себестоимость.
      a.costBasisRub = Math.max(
        0,
        a.costBasisRub - avgCostRate * qty
      );

      a.qty -= qty;

      if (Math.abs(a.qty) < CASH_EPS_) {
        a.qty = 0;
        a.costBasisRub = 0;
      }

      return;
    }

    throw new Error(
      'Наличные, строка ' +
      (index + 2) +
      ': неизвестная операция "' +
      row[1] +
      '". Допустимо: Покупка, Продажа, Пополнение, Снятие, Проценты.'
    );
  });

  return assets;
}


/**
 * Добавляет текущую оценку активов.
 *
 * Для RUB-подобных активов курс всегда 1.
 * Для остальных:
 *   1) ручной курс из текущей сводки по точному активу;
 *   2) курс из сводки по базовой валюте;
 *   3) курс последней операции.
 */
function enrichCashAssets_(assets) {
  let rates = readCashCurrentRates_();

  Object.keys(assets).forEach(function(key) {
    let a = assets[key];

    let currentRate = 0;

    if (a.baseCurrency === 'RUB') {
      currentRate = 1;
    } else {
      currentRate =
        rates.byAsset[String(a.asset).toUpperCase()] ||
        rates.byCurrency[a.baseCurrency] ||
        a.lastDealRate ||
        0;
    }

    a.currentRate = currentRate;
    a.valueRub = currentRate > 0
      ? a.qty * currentRate
      : 0;

    // Для рублёвого актива курсовой P/L не имеет смысла:
    // доход по накопительному счёту показываем отдельно.
    if (a.baseCurrency === 'RUB') {
      a.avgBuyRate = 1;
      a.unrealizedPL = 0;
      a.unrealizedPct = 0;
    } else {
      a.avgBuyRate =
        a.qty > CASH_EPS_ && a.costBasisRub > 0
          ? a.costBasisRub / a.qty
          : 0;

      a.unrealizedPL =
        a.valueRub - a.costBasisRub;

      a.unrealizedPct =
        a.costBasisRub > 0
          ? a.unrealizedPL / a.costBasisRub
          : 0;
    }
  });

  return assets;
}


/**
 * Позиции для общего портфеля.
 *
 * Используется dashboard.js.
 */
function getManualCashPositions_() {
  let assets = enrichCashAssets_(readCashLedger_());
  let positions = [];

  Object.keys(assets).forEach(function(key) {
    let a = assets[key];

    if (a.qty <= CASH_EPS_ || a.currentRate <= 0) {
      return;
    }

    let displayName =
      a.asset.toUpperCase() === a.baseCurrency
        ? 'Наличные ' + a.asset
        : a.asset;

    positions.push({
      name: displayName,
      ticker: a.asset,
      category: 'Кэш',
      valueRub: a.valueRub,
      price: a.currentRate,
      qty: a.qty,
      lot: 1,
      source: 'Manual',
      avgBuyRate: a.avgBuyRate,
      realizedPL: a.realizedPL,
      incomeRub: a.incomeRub
    });
  });

  return positions;
}


/**
 * Полностью пересобирает только автоматическую таблицу I:R.
 *
 * A:G — журнал пользователя — не меняется вообще.
 */
function updateCashSummary_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CASH_SHEET_);

  if (!sh) {
    sh = ss.insertSheet(CASH_SHEET_);
  }

  let assets = enrichCashAssets_(readCashLedger_());

  let active = Object.keys(assets)
    .map(function(key) { return assets[key]; })
    .filter(function(a) {
      return a.qty > CASH_EPS_;
    })
    .sort(function(a, b) {
      return b.valueRub - a.valueRub;
    });

  let totalCashRub = active.reduce(function(sum, a) {
    return sum + (a.valueRub || 0);
  }, 0);

  // Очищаем ТОЛЬКО автоматическую область I:R.
  let rowsToClear = Math.max(sh.getMaxRows(), 50);
  sh.getRange(
    1,
    CASH_SUMMARY_COL_,
    rowsToClear,
    CASH_SUMMARY_COLS_
  ).clearContent().clearFormat();

  let headers = [
    'Актив',
    'Текущий курс',
    'Кол-во',
    'На сумму',
    '% от наличных',
    'Средний курс покупки',
    'P/L ₽',
    'P/L %',
    'Реализ. P/L',
    'Проценты / доход'
  ];

  sh.getRange(1, CASH_SUMMARY_COL_, 1, CASH_SUMMARY_COLS_)
    .setValues([headers])
    .setBackground(C.DARK)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  if (active.length) {
    let rows = active.map(function(a) {
      return [
        a.asset,
        a.currentRate || '',
        a.qty,
        a.valueRub || 0,
        totalCashRub > 0 ? a.valueRub / totalCashRub : 0,
        a.avgBuyRate || '',
        a.unrealizedPL || 0,
        a.unrealizedPct || 0,
        a.realizedPL || 0,
        a.incomeRub || 0
      ];
    });

    sh.getRange(
      2,
      CASH_SUMMARY_COL_,
      rows.length,
      CASH_SUMMARY_COLS_
    ).setValues(rows);

    // J: курс
    sh.getRange(2, 10, rows.length, 1)
      .setNumberFormat('#,##0.0000');

    // K: количество
    sh.getRange(2, 11, rows.length, 1)
      .setNumberFormat('#,##0.####');

    // L: текущая стоимость
    sh.getRange(2, 12, rows.length, 1)
      .setNumberFormat('#,##0.00 [$₽-ru-RU]');

    // M: доля
    sh.getRange(2, 13, rows.length, 1)
      .setNumberFormat('0.0%');

    // N: средний курс
    sh.getRange(2, 14, rows.length, 1)
      .setNumberFormat('#,##0.0000');

    // O / Q / R: рубли
    [15, 17, 18].forEach(function(col) {
      sh.getRange(2, col, rows.length, 1)
        .setNumberFormat('+#,##0.00 [$₽-ru-RU];-#,##0.00 [$₽-ru-RU];0.00 [$₽-ru-RU]');
    });

    // P: P/L %
    sh.getRange(2, 16, rows.length, 1)
      .setNumberFormat('+0.0%;-0.0%;0.0%');

    // Выравнивание и зебра.
    for (let i = 0; i < rows.length; i++) {
      let bg = i % 2 === 0 ? C.EVEN : C.ODD;
      sh.getRange(i + 2, CASH_SUMMARY_COL_, 1, CASH_SUMMARY_COLS_)
        .setBackground(bg);
    }

    // В J пользователь может вручную менять текущий курс.
    // Помечаем эту колонку визуально как ввод.
    sh.getRange(2, 10, rows.length, 1)
      .setBackground(C.INPUT);
  }

  // Итоговая строка
  let totalRow = active.length + 3;
  sh.getRange(totalRow, CASH_SUMMARY_COL_, 1, 3)
    .merge()
    .setValue('ИТОГО НАЛИЧНЫЕ')
    .setBackground(C.MID)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sh.getRange(totalRow, 12)
    .setValue(totalCashRub)
    .setNumberFormat('#,##0.00 [$₽-ru-RU]')
    .setBackground(C.MID)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sh.getRange(totalRow, 13)
    .setValue(totalCashRub > 0 ? 1 : 0)
    .setNumberFormat('0.0%')
    .setBackground(C.MID)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Подсказка.
  let noteRow = totalRow + 2;
  let shortHint =
    '💡 A:G — ручной журнал · J — текущий курс · наведите для инструкции';

  let fullHint =
    'A:G — ручной журнал (скрипт его не изменяет).\n\n' +
    'Жёлтая колонка «Текущий курс» редактируется вручную ' +
    'и сохраняется при обновлении.\n\n' +
    'Поддерживаемые операции:\n' +
    '• Покупка\n' +
    '• Продажа\n' +
    '• Пополнение\n' +
    '• Снятие\n' +
    '• Проценты\n\n' +
    'Базовая валюта определяется по последнему слову в названии актива.\n' +
    'Например:\n' +
    '• USD → USD\n' +
    '• Вклад RUB → RUB\n' +
    '• Сейф RUB → RUB';

  let noteRange = sh.getRange(
    noteRow,
    CASH_SUMMARY_COL_,
    1,
    CASH_SUMMARY_COLS_
  );

  noteRange
    .merge()
    .setValue(shortHint)
    .setNote(fullHint)
    .setFontColor('#666666')
    .setFontStyle('italic')
    .setFontSize(9)
    .setWrap(false)
    .setHorizontalAlignment('left');

  sh.setRowHeight(noteRow, 22);

  // Ширины A:G не трогаем — это ручная часть пользователя.
  [
    150, 105, 90, 130, 105,
    145, 120, 90, 120, 125
  ].forEach(function(width, idx) {
    sh.setColumnWidth(CASH_SUMMARY_COL_ + idx, width);
  });

  SpreadsheetApp.flush();
}
