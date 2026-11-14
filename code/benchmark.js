/**
 * benchmark.js — Сравнение портфеля с IMOEX + единая точка входа «Аналитика доходности»
 *
 * ВАЖНО:
 * - Использует ту же дату начала, что и XIRR (readXirrStartDate_()).
 * - Учитывает и пополнения, и выводы.
 * - Физические наличные (source === 'Manual') не входят в фактическую
 *   стоимость инвестиционного портфеля для сравнения с IMOEX.
 */

const IMOEX_UID_PROPERTY = 'IMOEX_UID';
const BENCH_YEARS_BACK   = 5;

const ANALYTICS_SECTION_TITLE = '▌ АНАЛИТИКА ДОХОДНОСТИ';
const ANALYTICS_XIRR_PROP     = 'ANALYTICS_XIRR_VALUE';
const ANALYTICS_BENCH_PROP    = 'ANALYTICS_BENCH_VALUE';

function getImoexUid_() {
  let cached = PropertiesService.getScriptProperties().getProperty(IMOEX_UID_PROPERTY);
  if (cached) return cached;

  let resp = tiFetch_(
    '/tinkoff.public.invest.api.contract.v1.InstrumentsService/Indicatives',
    {}
  );

  let list = resp.instruments || [];
  let imoex = list.find(function(i) {
    return i.ticker === 'IMOEX';
  });

  if (!imoex) {
    throw new Error('Не удалось найти IMOEX в списке индикативов T-Invest API');
  }

  PropertiesService.getScriptProperties().setProperty(IMOEX_UID_PROPERTY, imoex.uid);
  return imoex.uid;
}

function getImoexPriceMap_(uid, fromDate, toDate) {
  let priceMap = {};
  let chunkStart = new Date(fromDate);

  while (chunkStart < toDate) {
    let chunkEnd = new Date(
      Math.min(
        chunkStart.getTime() + 365 * 24 * 3600 * 1000,
        toDate.getTime()
      )
    );

    let resp = tiFetch_(
      '/tinkoff.public.invest.api.contract.v1.MarketDataService/GetCandles',
      {
        instrumentId: uid,
        from: chunkStart.toISOString(),
        to: chunkEnd.toISOString(),
        interval: 'CANDLE_INTERVAL_DAY',
      }
    );

    (resp.candles || []).forEach(function(c) {
      priceMap[c.time.substring(0, 10)] = qToNumber_(c.close);
    });

    chunkStart = new Date(chunkEnd.getTime() + 24 * 3600 * 1000);
    pause_(60);
  }

  return priceMap;
}

function findNearestPrice_(priceMap, date) {
  for (let i = 0; i <= 7; i++) {
    let d = new Date(date.getTime() - i * 24 * 3600 * 1000);
    let key = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');

    if (priceMap[key]) {
      return priceMap[key];
    }
  }

  return null;
}

function getBenchmarkStartDate_() {
  if (typeof readXirrStartDate_ === 'function') {
    let configured = readXirrStartDate_();

    if (configured instanceof Date && !isNaN(configured.getTime())) {
      return configured;
    }
  }

  return new Date(
    new Date().getTime() -
    BENCH_YEARS_BACK * 365 * 24 * 3600 * 1000
  );
}

function calculateBenchmark() {
  let uid;

  try {
    uid = getImoexUid_();
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Ошибка получения IMOEX: ' + e.message
    );
    return;
  }

  let fromDate = getBenchmarkStartDate_();
  let toDate = new Date();

  if (fromDate >= toDate) {
    SpreadsheetApp.getUi().alert(
      'Дата начала аналитики должна быть раньше сегодняшней даты.'
    );
    return;
  }

  let cashFlows = [];

  getAccounts_().forEach(function(acc) {
    let ops = fetchOperations_(
      acc.id,
      acc.name,
      fromDate,
      toDate
    );

    ops.forEach(function(op) {
      if (op.type === 'OPERATION_TYPE_INPUT') {
        cashFlows.push({
          date: op.date,
          amount: Math.abs(op.amount),
          type: 'input',
        });
      }

      if (op.type === 'OPERATION_TYPE_OUTPUT') {
        cashFlows.push({
          date: op.date,
          amount: -Math.abs(op.amount),
          type: 'output',
        });
      }
    });
  });

  cashFlows.sort(function(a, b) {
    return a.date - b.date;
  });

  let inputCount = cashFlows.filter(function(f) {
    return f.type === 'input';
  }).length;

  if (inputCount === 0) {
    SpreadsheetApp.getUi().alert(
      'После выбранной даты нет ни одного пополнения — сравнивать с IMOEX не с чем.'
    );
    return;
  }

  let priceFromDate = new Date(
    fromDate.getTime() - 7 * 24 * 3600 * 1000
  );

  let priceMap = getImoexPriceMap_(
    uid,
    priceFromDate,
    toDate
  );

  let currentPrice = findNearestPrice_(priceMap, toDate);

  if (!currentPrice) {
    SpreadsheetApp.getUi().alert(
      'Не удалось получить текущую цену IMOEX через T-Invest API.'
    );
    return;
  }

  let totalUnits = 0;
  let skipped = 0;
  let processedInputs = 0;
  let processedOutputs = 0;

  cashFlows.forEach(function(flow) {
    let priceThen = findNearestPrice_(
      priceMap,
      flow.date
    );

    if (!priceThen || priceThen <= 0) {
      skipped++;
      return;
    }

    totalUnits += flow.amount / priceThen;

    if (flow.type === 'input') {
      processedInputs++;
    } else {
      processedOutputs++;
    }
  });

  let hypotheticalValue = Math.round(
    totalUnits * currentPrice
  );

  let config = readConfig_();

  let positions = readPositions_(config)
    .filter(function(p) {
      return p.source !== 'Manual';
    });

  let actualValue = Math.round(
    positions.reduce(function(s, p) {
      return s + p.valueRub;
    }, 0)
  );

  writeBenchmarkToDashboard_(
    actualValue,
    hypotheticalValue,
    skipped,
    fromDate,
    processedInputs,
    processedOutputs
  );
}

function writeBenchmarkToDashboard_(
  actualValue,
  hypotheticalValue,
  skipped,
  startDate,
  inputCount,
  outputCount
) {
  let tz = Session.getScriptTimeZone();

  let startDateStr =
    startDate instanceof Date && !isNaN(startDate.getTime())
      ? Utilities.formatDate(startDate, tz, 'dd.MM.yyyy')
      : '';

  PropertiesService.getScriptProperties().setProperty(
    ANALYTICS_BENCH_PROP,
    JSON.stringify({
      actual: actualValue,
      hypothetical: hypotheticalValue,
      skipped: skipped,
      startDate: startDateStr,
      inputCount: inputCount || 0,
      outputCount: outputCount || 0,
    })
  );
}

function calculateAnalytics() {
  withLock_('Рассчитать доходность', function() {
    calculateXIRR();
    calculateBenchmark();
    redrawAnalyticsSection_();
  });
}

function redrawAnalyticsSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;

  let props = PropertiesService.getScriptProperties();
  let xirrRaw = props.getProperty(ANALYTICS_XIRR_PROP);
  let benchRaw = props.getProperty(ANALYTICS_BENCH_PROP);

  if (!xirrRaw && !benchRaw) return;

  renderSection_(
    sh,
    ANALYTICS_SECTION_TITLE,
    function(sh, r, COLS, colStart) {
      if (xirrRaw) {
        let xirr = JSON.parse(xirrRaw);
        let pct = xirr.value;

        let bg = pct === null
          ? '#9e9e9e'
          : pct >= 0
            ? '#1b5e20'
            : '#b71c1c';

        let arrow = pct === null
          ? ''
          : pct >= 0
            ? '▲'
            : '▼';

        let label = pct === null
          ? 'н/д'
          : arrow + ' ' + Math.abs(pct).toFixed(1) + '% годовых';

        r = renderTile_(
          sh,
          r,
          COLS,
          '📈 Доходность портфеля (XIRR)',
          label,
          bg,
          colStart
        );

        if (pct !== null) {
          let params = readAdvancedParams_();

          let realPct =
            ((1 + pct / 100) /
             (1 + params.inflationPct / 100) - 1) * 100;

          let realBg = realPct >= 0
            ? '#1b5e20'
            : '#b71c1c';

          let realArrow = realPct >= 0
            ? '▲'
            : '▼';

          let realLabel =
            realArrow +
            ' ' +
            Math.abs(realPct).toFixed(1) +
            '% годовых' +
            '  (при инфляции ' +
            params.inflationPct.toFixed(1) +
            '%)';

          r = renderTile_(
            sh,
            r,
            COLS,
            '💰 Реальная доходность (с поправкой на инфляцию)',
            realLabel,
            realBg,
            colStart
          );
        }

        // Простой фактический результат периода:
        // не годовой темп, а сколько реально заработано/потеряно
        // относительно внесённых денег.
        if (
          xirr.simpleReturnPct !== null &&
          xirr.simpleReturnPct !== undefined &&
          xirr.profitRub !== null &&
          xirr.profitRub !== undefined
        ) {
          let simplePct = Number(xirr.simpleReturnPct);
          let profitRub = Number(xirr.profitRub);

          let simpleBg = profitRub >= 0
            ? '#1b5e20'
            : '#b71c1c';

          let simpleArrow = profitRub >= 0
            ? '▲'
            : '▼';

          let signedRub =
            (profitRub >= 0 ? '+' : '−') +
            rub_(Math.abs(profitRub));

          let simpleLabel =
            simpleArrow +
            ' ' +
            Math.abs(simplePct).toFixed(1) +
            '%  ·  ' +
            signedRub;

          let simpleTitle =
            '💵 Фактический результат периода' +
            (xirr.startDate
              ? ' с ' + xirr.startDate
              : '');

          r = renderTile_(
            sh,
            r,
            COLS,
            simpleTitle,
            simpleLabel,
            simpleBg,
            colStart
          );
        }
      }

      if (benchRaw) {
        let b = JSON.parse(benchRaw);

        let diff = b.actual - b.hypothetical;

        let bg = diff >= 0
          ? '#1b5e20'
          : '#b71c1c';

        let arrow = diff >= 0
          ? '▲'
          : '▼';

        let benchTitle =
          '📊 vs IMOEX' +
          (b.startDate ? ' с ' + b.startDate : '');

        r = renderTile_(
          sh,
          r,
          COLS,
          benchTitle,
          arrow + ' ' + rub_(Math.abs(diff)),
          bg,
          colStart
        );

        [
          ['Ваш портфель', rub_(b.actual)],
          ['Если бы покупали IMOEX', rub_(b.hypothetical)]
        ].forEach(function(row) {
          sh.getRange(r, colStart, 1, 3)
            .merge()
            .setValue(row[0]);

          sh.getRange(r, colStart + 3, 1, COLS - 3)
            .merge()
            .setValue(row[1])
            .setFontWeight('bold')
            .setHorizontalAlignment('right');

          sh.getRange(r, colStart, 1, COLS)
            .setBackground(C.EVEN);

          r++;
        });

        if (b.skipped > 0) {
          sh.getRange(r, colStart, 1, COLS)
            .merge()
            .setValue(
              '⚠️ ' +
              b.skipped +
              ' денежных потоков не удалось сопоставить с ценой IMOEX'
            )
            .setFontColor(C.WARN)
            .setFontStyle('italic')
            .setFontSize(9);

          r++;
        }
      }

      return r;
    },
    'right'
  );
}
