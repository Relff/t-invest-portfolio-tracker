/**
 * income_v4.js — Лист «Ожидаемый доход»
 *
 * v5: дивиденды теперь считаются автоматически из истории выплат (GetDividends),
 * если не заданы вручную в Config → Блок 4. Ручной ввод остаётся приоритетным.
 * v6: автооценка дивидендов теперь по CAGR-прогнозу (fetchDividendProjection_) —
 * учитывает темп роста выплат за последние годы, а не только факт за 12 мес.
 * v7: добавлен Payout Ratio по EPS (урезанный Dividend Safety Score) —
 * T-Invest API не даёт надёжный FCF, поэтому только EPS.
 * v8: Payout считается по ФАКТУ дивиденда за 12 мес. (не по CAGR-прогнозу —
 * сравнивать прогноз с прошлогодним EPS нечестно), вынесен в отдельную
 * колонку таблицы вместо текста в примечании.
 */

const BOND_TOP_N = 15;

// ════════════════════════════════════════════════════════════════════
// ГЛАВНАЯ ФУНКЦИЯ
// ════════════════════════════════════════════════════════════════════

function updateIncomeSheet() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.INCOME);
  if (!sh) sh = ss.insertSheet(DST.INCOME);
  sh.clearContents();
  sh.clearFormats();

  var config;
  try { config = readConfig_(); }
  catch(e) { sh.getRange(1,1).setValue('⚠️ ' + e.message); return; }

  var totalRub  = getTotalPortfolioValue_();
  var tz        = Session.getScriptTimeZone();
  var now       = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
  var divMap    = readDividendsFromConfig_();
  var figiMap   = buildFigiMap_();

  // Читаем акции и облигации напрямую
  var sharePositions = readSheetPositions_('_Дан_Акции');
  var bondPositions  = readSheetPositions_('_Дан_Облигации');

  // Считаем доход
  var shareRows = calcShareIncome_(sharePositions, divMap, figiMap, totalRub);
  var bondRows  = calcBondIncome_(bondPositions, figiMap, totalRub);

  // Сортировка
  shareRows.sort(function(a,b){ return a.name.localeCompare(b.name,'ru'); });
  bondRows.sort(function(a,b){ return b.incomeYear - a.incomeYear; });

  // Итоги
  var shareTotal = shareRows.reduce(function(s,r){ return s + r.incomeYear; }, 0);
  var bondTotal  = bondRows.reduce(function(s,r){ return s + r.incomeYear; }, 0);
  var grandTotal = shareTotal + bondTotal;
  var yieldPct   = totalRub > 0 ? (grandTotal / totalRub * 100) : 0;

  var COLS = 8;
  var r    = 1;

  // ── ШАПКА ─────────────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS,
    '💵  ОЖИДАЕМЫЙ ПАССИВНЫЙ ДОХОД — ближайшие 12 месяцев',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });
  r++;

  mergedCell_(sh, r, 1, 1, COLS,
    'Обновлено: ' + now + '   ·   Портфель: ' + rub_(totalRub),
    { bg: '#263238', fg: '#b0bec5', align: 'center' });
  r++;

  // ── 4 ПЛИТКИ ─────────────────────────────────────────────────────
  var half = Math.floor(COLS / 2);
  var tiles = [
    { label: '📅 Доход в год',   val: rub_(grandTotal), bg: '#1565c0' },
    { label: '📈 Доходность',    val: yieldPct.toFixed(1) + '%', bg: '#1565c0' },
    { label: '📊 Акции',        val: rub_(shareTotal), bg: C.MID },
    { label: '🏦 Облигации',    val: rub_(bondTotal),  bg: C.MID },
  ];
  for (var ti = 0; ti < tiles.length; ti++) {
    var col = (ti % 2 === 0) ? 1 : half + 1;
    if (ti % 2 === 0 && ti > 0) r++;
    sh.getRange(r, col, 1, half).merge()
      .setValue(tiles[ti].label + ':  ' + tiles[ti].val)
      .setBackground(tiles[ti].bg).setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  }
  r++;
  r++;

  // ── АКЦИИ ─────────────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '📈  АКЦИИ — ДИВИДЕНДЫ',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 12 });
  r++;
  hdrRow_(sh, r,
    ['Название','Тикер','Кол-во','Дивиденд/акц, ₽','Доход в год, ₽','% порт.','Примечание','Payout'],
    COLS);
  r++;

  if (shareRows.length === 0) {
    mergedCell_(sh, r, 1, 1, COLS,
      '⚠️  Нет акций в портфеле или лист Дан_Акции пуст',
      { fg: C.WARN, italic: true });
    r++;
  } else {
    shareRows.forEach(function(row, idx) {
      writeIncomeRow_(sh, r, row, idx, totalRub);
      r++;
    });
  }
  writeTotalRow_(sh, r, '  Итого по акциям', shareTotal, COLS, '#1b5e20');
  r += 2;

  // ── ОБЛИГАЦИИ ─────────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '🏦  ОБЛИГАЦИИ — КУПОНЫ (ближайшие 12 мес.)',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 12 });
  r++;
  hdrRow_(sh, r,
    ['Название','Тикер','Кол-во','Купон/год на шт, ₽','Доход в год, ₽','% порт.','Купонов',''],
    COLS);
  r++;

  var topBonds  = bondRows.slice(0, BOND_TOP_N);
  var restBonds = bondRows.slice(BOND_TOP_N);

  topBonds.forEach(function(row, idx) {
    writeIncomeRow_(sh, r, row, idx, totalRub);
    r++;
  });

  // Строка «Остальные» с суммой купонов
  if (restBonds.length > 0) {
    var restIncome = restBonds.reduce(function(s,x){ return s + x.incomeYear; }, 0);
    var restValue  = restBonds.reduce(function(s,x){ return s + x.valueRub;   }, 0);
    var restPct    = totalRub > 0 ? restValue / totalRub : 0;
    var restBg     = topBonds.length % 2 === 0 ? C.EVEN : C.ODD;

    sh.getRange(r, 1, 1, COLS).setValues([[
      '…ещё ' + restBonds.length + ' облигаций',
      '', '', '',
      restIncome, restPct,
      'сумма купонов: ' + rub_(restIncome), ''
    ]]).setBackground(restBg).setFontStyle('italic').setFontColor('#546e7a');
    sh.getRange(r, 5).setNumberFormat('#,##0 [$₽-ru-RU]')
      .setBackground('#e3f2fd').setFontColor('#0d47a1')
      .setFontStyle('normal').setFontWeight('bold');
    sh.getRange(r, 6).setNumberFormat('0.0%');
    r++;
  }

  writeTotalRow_(sh, r, '  Итого по облигациям', bondTotal, COLS, '#0d47a1');
  r += 2;

  // ── ВЫВОДЫ ────────────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '💡  ВЫВОДЫ',
    { bg: '#37474f', fg: '#ffffff', bold: true });
  r++;

  // Топ-3 облигации по доходу
  var top3 = bondRows.slice(0, 3);
  var top3txt = top3.map(function(b, i) {
    return (i+1) + '. ' + b.name + ' — ' + rub_(b.incomeYear);
  }).join('   |   ');

  writeConclusion_(sh, r, '🏆 Топ-3 облигации по доходу:', top3txt); r++;

  // Акции с дивидендами
  var divStocks = shareRows.filter(function(x){ return x.incomeYear > 0; });
  var divTxt = divStocks.length > 0
    ? divStocks.map(function(s){ return s.name + ' (' + rub_(s.incomeYear) + ')'; }).join(', ')
    : 'Пока нет дивидендных акций';
  writeConclusion_(sh, r, '💰 Дивиденды по акциям:', divTxt); r++;

  // Структура дохода
  var bondShare = grandTotal > 0 ? Math.round(bondTotal / grandTotal * 100) : 0;
  writeConclusion_(sh, r, '📊 Структура дохода:',
    'Облигации ' + bondShare + '%  ·  Акции ' + (100-bondShare) + '%'); r++;

  // Ежемесячный доход
  writeConclusion_(sh, r, '📅 Ежемесячно в среднем:',
    rub_(grandTotal / 12) + ' (~' + rub_(bondTotal / 12) + ' от купонов + ' +
    rub_(shareTotal / 12) + ' от дивидендов)'); r++;

  // Краткий вывод
  var conclusion = grandTotal === 0
    ? 'Нет данных для расчёта.'
    : shareTotal === 0
    ? 'Весь доход сейчас от облигаций (купоны). После покупки дивидендных акций (Сбер, Лукойл) картина изменится.'
    : bondShare > 80
    ? 'Основной доход от облигаций. Акции пока дают мало — продолжай наращивать дивидендные позиции.'
    : 'Хороший баланс между купонами и дивидендами.';
  writeConclusion_(sh, r, '💡 Вывод:', conclusion); r++;
  r++;

  // ── ОБЩИЙ ИТОГ ────────────────────────────────────────────────────
  sh.getRange(r, 1, 1, COLS).merge()
    .setValue('ИТОГО В ГОД:  ' + rub_(grandTotal) +
              '   ·   ' + yieldPct.toFixed(1) + '% годовых   ·   ' +
              rub_(grandTotal / 12) + ' в месяц')
    .setBackground(C.DARK).setFontColor('#ffd54f')
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');

  // Сохраняем итоги — переиспользуются в прогнозе денежного потока
  // на HTML-дашборде (htmldashboard.js), без повторных запросов к API.
  PropertiesService.getScriptProperties().setProperty('INCOME_TOTALS_DATA', JSON.stringify({
    shareTotal: Math.round(shareTotal),
    bondTotal:  Math.round(bondTotal),
    grandTotal: Math.round(grandTotal),
  }));

  // ── Ширина колонок ────────────────────────────────────────────────
  [255, 65, 75, 135, 155, 80, 210, 80].forEach(function(w,i){
    sh.setColumnWidth(i+1, w);
  });
  sh.setFrozenRows(6);
  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════════════
// ЧТЕНИЕ ПОЗИЦИЙ НАПРЯМУЮ ИЗ ЛИСТОВ
// ════════════════════════════════════════════════════════════════════

function readSheetPositions_(sheetName) {
  var ss   = SpreadsheetApp.getActive();
  var sh   = ss.getSheetByName(sheetName);
  if (!sh) return [];

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  var H = {};
  data[0].forEach(function(h,i){ H[String(h).trim()] = i; });

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row      = data[i];
    var name     = String(row[H['name']]   || '').trim();
    var ticker   = String(row[H['ticker']] || '').trim();
    var valueRub = Number(row[H['position_value_rub']]          || 0);
    var qty      = Number(row[H['quantity_pcs']]                || 0);
    var figi     = String(row[H['figi']]   || '').trim();

    if (!name || valueRub === 0) continue;
    result.push({ name: name, ticker: ticker, valueRub: valueRub, qty: qty, figi: figi });
  }
  return result;
}

function getTotalPortfolioValue_() {
  var ss  = SpreadsheetApp.getActive();
  var src = ['_Дан_Акции','_Дан_Облигации','_Дан_Фонды','_Дан_Деньги','_Дан_Валюта'];
  var total = 0;
  src.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var H = {};
    data[0].forEach(function(h,i){ H[String(h).trim()] = i; });
    for (var i = 1; i < data.length; i++) {
      total += Number(data[i][H['position_value_rub']] || 0);
    }
  });
  return total;
}


// ════════════════════════════════════════════════════════════════════
// РАСЧЁТ ДОХОДА ПО АКЦИЯМ И ОБЛИГАЦИЯМ
// ════════════════════════════════════════════════════════════════════

function calcShareIncome_(positions, divMap, figiMap, totalRub) {
  var assetUidMap = buildAssetUidMap_(); // один запрос на всю таблицу, не на каждую акцию

  return positions.map(function(p) {
    var div  = findDividend_(p.name, p.ticker, divMap);
    var figi = p.figi || (figiMap && (figiMap[p.name] || figiMap[p.ticker])) || '';

    // payoutBasis — именно ФАКТ дивиденда за 12 мес., для честного сравнения
    // с EPS (тоже TTM-факт). По умолчанию — то же, что и amount (например,
    // вручную заданное в Config), но если дивиденд считаем сами — берём
    // hist.trailing12, а НЕ hist.perUnit: perUnit может быть CAGR-прогнозом
    // на будущий год, а сравнивать прогноз с прошлогодним EPS — заведомо
    // нечестно (отсюда Payout за 200%+ у растущих компаний).
    var payoutBasis = div.amount;

    // Если в Config ничего не задано (0 или отсутствует) — считаем сами,
    // с приоритетом на CAGR-прогноз роста (fetchDividendProjection_)
    if (div.amount <= 0 && figi) {
      var hist = fetchDividendProjection_(figi);
      Utilities.sleep(50);
      if (hist.perUnit > 0) {
        div = { amount: hist.perUnit, note: hist.note };
        payoutBasis = hist.trailing12;
      }
    }

    var incomePerUnit = div.amount;
    var incomeYear    = incomePerUnit * p.qty;

    // Payout Ratio по EPS — насколько дивиденд (факт за 12 мес.) "съедает"
    // прибыль на акцию. Урезанная версия Dividend Safety Score: T-Invest
    // API не даёт надёжный FCF (часто нули), поэтому только EPS —
    // единственное реально заполняемое поле. GetAssetFundamentals ждёт
    // asset_uid, не figi — берём его из assetUidMap. См. dividendSafetyLabel_().
    var assetUid   = figi ? assetUidMap[figi] : null;
    var payoutInfo = null;
    if (payoutBasis > 0 && assetUid) {
      payoutInfo = dividendSafetyLabel_(payoutBasis, assetUid);
      Utilities.sleep(50);
    }

    return {
      name:          p.name,
      ticker:        p.ticker,
      qty:           p.qty,
      incomePerUnit: incomePerUnit,
      incomeYear:    incomeYear,
      valueRub:      p.valueRub,
      note:          div.note,
      payoutText:    (payoutInfo && payoutInfo.label !== 'н/д') ? payoutInfo.label : '',
      payoutColor:   (payoutInfo && payoutInfo.color) ? payoutInfo.color : null,
    };
  });
}

function calcBondIncome_(positions, figiMap, totalRub) {
  return positions.map(function(p) {
    var figi = p.figi || figiMap[p.name] || figiMap[p.ticker] || '';
    var coupon = { perUnit: 0, note: 'FIGI не найден' };
    if (figi) {
      coupon = fetchAnnualCoupon_(figi);
      Utilities.sleep(50);
    }
    return {
      name:          p.name,
      ticker:        p.ticker,
      qty:           p.qty,
      incomePerUnit: coupon.perUnit,
      incomeYear:    coupon.perUnit * p.qty,
      valueRub:      p.valueRub,
      note:          coupon.note,
    };
  });
}


// ════════════════════════════════════════════════════════════════════
// КУПОНЫ ЧЕРЕЗ T-INVEST API
// ════════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════════
// PAYOUT RATIO ПО EPS (урезанный Dividend Safety Score)
// ════════════════════════════════════════════════════════════════════

/**
 * buildAssetUidMap_ — карта figi → asset_uid, ОДНИМ запросом (GetAssets).
 *
 * У T-Invest API два РАЗНЫХ идентификатора: figi (используется почти
 * везде — GetDividends, GetBondCoupons и т.д.) и asset_uid — судя по
 * официальной proto-схеме (Asset.uid + AssetInstrument.figi), именно
 * его ждёт GetAssetFundamentals, а не figi. Строим карту заранее, а не
 * дёргаем по одному запросу на каждую акцию — иначе цена одной лишней
 * метрики была бы в 2 раза больше запросов, чем нужно.
 */
function buildAssetUidMap_() {
  var map = {};
  try {
    var resp = tiFetch_(
      '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetAssets', {}
    );
    var assets = resp.assets || [];
    assets.forEach(function(asset) {
      (asset.instruments || []).forEach(function(instr) {
        if (instr.figi) map[instr.figi] = asset.uid;
      });
    });
    Logger.log('buildAssetUidMap_: построена карта на ' + Object.keys(map).length + ' figi');
  } catch (e) {
    Logger.log('buildAssetUidMap_: ОШИБКА — ' + e.message);
  }
  return map;
}

/**
 * fetchAssetFundamentals_ — фундаментальные показатели по бумаге через
 * T-Invest API (GetAssetFundamentals), уже по правильному asset_uid.
 * Данные доступны НЕ по всем бумагам (недавние листинги, некоторые бумаги
 * без покрытия) — в таком случае тихо возвращает null, без ошибки.
 */
function fetchAssetFundamentals_(assetUid) {
  if (!assetUid) return null;
  try {
    var resp = tiFetch_(
      '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetAssetFundamentals',
      { assets: [assetUid] }
    );
    var list = resp.fundamentals || [];
    Logger.log('fetchAssetFundamentals_(' + assetUid + '): получено ' + list.length +
      ' записей. Сырой ответ: ' + JSON.stringify(resp).substring(0, 500));
    return list.length ? list[0] : null;
  } catch (e) {
    Logger.log('fetchAssetFundamentals_(' + assetUid + '): ОШИБКА — ' + e.message);
    return null;
  }
}

/**
 * Payout Ratio = дивиденд на акцию / EPS (прибыль на акцию, TTM).
 * Это НЕ полноценный Dividend Safety Score (для него нужен ещё тренд FCF
 * за несколько лет, а T-Invest API отдаёт только текущий срез, и часто
 * с нулевым FCF) — просто один честный сигнал на основе того, что реально
 * заполняется в API.
 *
 *   <40%  — 🟢 запас прочности большой
 *   40-70% — 🟡 платят большую часть прибыли
 *   >70%  — 🔴 почти вся прибыль уходит на дивиденды — риск урезания
 *   EPS ≤ 0 (убыток) или нет данных — «н/д», без ложной уверенности
 */
function dividendSafetyLabel_(dividendPerShare, assetUid) {
  var fund = fetchAssetFundamentals_(assetUid);
  if (!fund || !fund.epsTtm || fund.epsTtm <= 0) {
    Logger.log('dividendSafetyLabel_(' + assetUid + '): нет пригодного epsTtm (fund=' +
      (fund ? JSON.stringify(fund).substring(0, 200) : 'null') + ')');
    return { payout: null, label: 'н/д', color: null };
  }
  var payout = dividendPerShare / fund.epsTtm;
  var color  = payout < 0.4 ? '#1b5e20' : payout < 0.7 ? '#f57f17' : '#b71c1c';
  return { payout: payout, label: (payout * 100).toFixed(0) + '%', color: color };
}


function fetchAnnualCoupon_(figi) {
  if (!figi) return { perUnit: 0, note: 'Нет FIGI' };
  try {
    var now  = new Date();
    var in1y = new Date(now.getTime() + 365*24*3600*1000);
    var resp = tiFetch_(
      '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetBondCoupons',
      { figi: figi, from: now.toISOString(), to: in1y.toISOString() }
    );
    var coupons = resp.events || [];
    var total   = 0;
    coupons.forEach(function(c){
      total += moneyToNumber_(c.payOneBond || c.couponAmount || null);
    });
    return {
      perUnit: total,
      note:    coupons.length > 0 ? (coupons.length + ' купонов') : 'Нет данных'
    };
  } catch(e) {
    return { perUnit: 0, note: 'Ошибка: ' + e.message.substring(0,35) };
  }
}


// ════════════════════════════════════════════════════════════════════
// ДИВИДЕНДЫ ИЗ ИСТОРИИ ВЫПЛАТ (T-INVEST API) — автооценка
// ════════════════════════════════════════════════════════════════════

/**
 * fetchDividendProjection_ — оценка дивиденда на акцию на ближайшие 12 месяцев.
 *
 * В отличие от fetchAnnualDividendFromHistory_() (плоская сумма факта за
 * прошлые 12 месяцев), тянет историю выплат за ~4 года ОДНИМ запросом и:
 *
 *  - если есть минимум 2 полных календарных года истории — считает CAGR
 *    (среднегодовой темп роста дивиденда) и проецирует его поверх
 *    последнего полного года. Честнее для компаний со стабильно растущими
 *    выплатами (Сбер, Лукойл и т.п.) — плоская оценка систематически
 *    занижает будущий доход, если дивиденд из года в год растёт.
 *  - если истории недостаточно (недавний листинг, нерегулярные выплаты)
 *    — тихо откатывается на ту же плоскую сумму факта за 12 месяцев.
 *  - CAGR искусственно ограничен диапазоном ±50% в год — иначе разовый
 *    аномальный скачок (или пропуск выплаты в базовом году) экстраполируется
 *    буквально и даёт нереалистичный прогноз.
 */
function fetchDividendProjection_(figi) {
  if (!figi) return { perUnit: 0, trailing12: 0, note: 'Нет FIGI' };
  try {
    var now       = new Date();
    var yearsBack = 4;
    var from      = new Date(now.getTime() - yearsBack * 365 * 24 * 3600 * 1000);
    var resp = tiFetch_(
      '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetDividends',
      { figi: figi, from: from.toISOString(), to: now.toISOString() }
    );
    var events = resp.dividends || [];
    if (!events.length) return { perUnit: 0, trailing12: 0, note: 'Нет данных за ' + yearsBack + ' года' };

    function dividendDate_(d) { return new Date(d.paymentDate || d.recordDate || d.lastBuyDate || now); }
    function dividendAmount_(d) { return moneyToNumber_(d.dividendNet || d.dividendValue || null); }

    // Плоская сумма факта за последние 12 мес. — как в старом методе, она же fallback
    var ago12 = new Date(now.getTime() - 365 * 24 * 3600 * 1000);
    var trailing12 = 0;
    events.forEach(function(d) {
      var dt = dividendDate_(d);
      if (dt >= ago12 && dt <= now) trailing12 += dividendAmount_(d);
    });

    // Группировка по календарным годам для CAGR
    var byYear = {};
    events.forEach(function(d) {
      var y = dividendDate_(d).getFullYear();
      byYear[y] = (byYear[y] || 0) + dividendAmount_(d);
    });

    var currentYear   = now.getFullYear();
    var completeYears = Object.keys(byYear).map(Number)
      .filter(function(y) { return y < currentYear; })
      .sort(function(a, b) { return a - b; });

    if (completeYears.length >= 2) {
      var firstYear = completeYears[0];
      var lastYear  = completeYears[completeYears.length - 1];
      var firstVal  = byYear[firstYear];
      var lastVal   = byYear[lastYear];
      var n         = lastYear - firstYear;

      if (firstVal > 0 && lastVal > 0 && n > 0) {
        var cagr = Math.pow(lastVal / firstVal, 1 / n) - 1;
        cagr = Math.max(-0.5, Math.min(0.5, cagr)); // защита от аномальных скачков истории
        var projected = lastVal * (1 + cagr);
        if (projected > 0) {
          return {
            perUnit: projected,
            trailing12: trailing12, // факт за 12 мес. — для Payout Ratio, не путать с прогнозом
            note: 'CAGR-прогноз (' + completeYears.length + ' г., ' +
                  (cagr >= 0 ? '+' : '') + (cagr * 100).toFixed(1) + '%/г)',
          };
        }
      }
    }

    // Недостаточно истории для CAGR — плоская оценка по факту за 12 мес.
    return {
      perUnit: trailing12,
      trailing12: trailing12,
      note: trailing12 > 0 ? 'оценка (история, 12 мес.)' : 'Нет данных за 12 мес.',
    };
  } catch (e) {
    return { perUnit: 0, trailing12: 0, note: 'Ошибка: ' + e.message.substring(0, 35) };
  }
}


function fetchAnnualDividendFromHistory_(figi) {
  if (!figi) return { perUnit: 0, note: 'Нет FIGI' };
  try {
    var now   = new Date();
    var ago12 = new Date(now.getTime() - 365 * 24 * 3600 * 1000);
    var resp = tiFetch_(
      '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetDividends',
      { figi: figi, from: ago12.toISOString(), to: now.toISOString() }
    );
    var events = resp.dividends || [];
    var total  = 0;
    events.forEach(function(d){
      total += moneyToNumber_(d.dividendNet || d.dividendValue || null);
    });
    return {
      perUnit: total,
      note: events.length > 0
        ? 'оценка (история, ' + events.length + ' выпл.)'
        : 'Нет данных за 12 мес.'
    };
  } catch(e) {
    return { perUnit: 0, note: 'Ошибка: ' + e.message.substring(0,35) };
  }
}


// ════════════════════════════════════════════════════════════════════
// ДИВИДЕНДЫ ИЗ CONFIG (ручной приоритет)
// ════════════════════════════════════════════════════════════════════

function readDividendsFromConfig_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) return {};
  var v   = sh.getDataRange().getValues();
  var map = {};
  var startRow = -1;
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).includes('ДИВИДЕНДЫ')) { startRow = i + 2; break; }
  }
  if (startRow < 0) return {};
  for (var j = startRow; j < v.length; j++) {
    var name   = String(v[j][0]).trim();
    var amount = Number(v[j][1]) || 0;
    var ticker = String(v[j][2]).trim();
    if (!name) break;
    if (name)   map[name]   = { amount: amount, note: 'из Config' };
    if (ticker) map[ticker] = { amount: amount, note: 'из Config' };
  }
  return map;
}

function findDividend_(name, ticker, divMap) {
  if (divMap[name])   return divMap[name];
  if (divMap[ticker]) return divMap[ticker];
  var nl = name.toLowerCase();
  var keys = Object.keys(divMap);
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i].toLowerCase();
    if (nl.includes(k) || k.includes(nl)) return divMap[keys[i]];
  }
  return { amount: 0, note: '⚠️ Нет данных' };
}


// ════════════════════════════════════════════════════════════════════
// FIGI ИЗ ЛИСТА ПОЗИЦИИ
// ════════════════════════════════════════════════════════════════════

function buildFigiMap_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.POSITIONS);
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {};
  var H = {};
  data[0].forEach(function(h,i){ H[String(h).trim()] = i; });
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var name   = String(data[i][H['name']]   || '').trim();
    var ticker = String(data[i][H['ticker']] || '').trim();
    var figi   = String(data[i][H['figi']]   || '').trim();
    if (figi && name)   map[name]   = figi;
    if (figi && ticker) map[ticker] = figi;
  }
  return map;
}


// ════════════════════════════════════════════════════════════════════
// БЛОК 4 В CONFIG (теперь опциональный override, не обязательный)
// ════════════════════════════════════════════════════════════════════

function addDividendsBlock() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) { SpreadsheetApp.getUi().alert('⚠️ Сначала запустите initConfig()'); return; }

  var vals = sh.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).includes('ДИВИДЕНДЫ')) {
      SpreadsheetApp.getUi().alert('Блок 4 уже существует в Config.');
      return;
    }
  }

  var lastRow = sh.getLastRow() + 2;
  var block = [
    ['▌ ДИВИДЕНДЫ АКЦИЙ — заполняйте только если хотите переопределить автооценку', '', ''],
    ['Название (точно как в портфеле)', 'Дивиденд ₽/год', 'Тикер'],
  ];

  sh.getRange(lastRow, 1, block.length, 3).setValues(block);
  sh.getRange(lastRow, 1, 1, 3).merge()
    .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.getRange(lastRow+1, 1, 1, 3)
    .setBackground(C.MID).setFontColor('#ffffff').setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    '✅ Блок 4 добавлен!\n\n' +
    'Теперь можно оставить пустым — дивиденды считаются автоматически из истории.\n' +
    'Заполняй строку вручную только если хочешь переопределить автооценку конкретной бумаги.'
  );
}


// ════════════════════════════════════════════════════════════════════
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ════════════════════════════════════════════════════════════════════

function writeIncomeRow_(sh, row, data, idx, totalRub) {
  var bg     = idx % 2 === 0 ? C.EVEN : C.ODD;
  var pctPf  = (totalRub > 0 && data.valueRub > 0) ? data.valueRub / totalRub : 0;
  sh.getRange(row, 1, 1, 8).setValues([[
    data.name, data.ticker, data.qty,
    data.incomePerUnit, data.incomeYear,
    pctPf, data.note, data.payoutText || ''
  ]]).setBackground(bg);

  sh.getRange(row, 3).setNumberFormat('0');
  sh.getRange(row, 4).setNumberFormat('#,##0.00 [$₽-ru-RU]');
  sh.getRange(row, 6).setNumberFormat('0.0%');

  var incCell = sh.getRange(row, 5);
  incCell.setNumberFormat('#,##0 [$₽-ru-RU]');
  var inc = data.incomeYear;
  if      (inc <= 0)    incCell.setBackground('#ffcdd2').setFontColor('#b71c1c');
  else if (inc < 1000)  incCell.setBackground('#fff9c4').setFontColor('#f57f17');
  else if (inc < 5000)  incCell.setBackground('#c8e6c9').setFontColor('#1b5e20');
  else                  incCell.setBackground('#2e7d32').setFontColor('#ffffff');

  // Подсветка источника цифры: зелёный — подтверждено вручную, жёлтый — оценка по истории
  var noteCell = sh.getRange(row, 7);
  var noteStr  = String(data.note || '');
  if (noteStr.indexOf('оценка') === 0 || noteStr.indexOf('CAGR-прогноз') === 0) {
    noteCell.setBackground('#fff9c4').setFontColor('#f57f17').setFontStyle('italic');
  } else if (noteStr.indexOf('из Config') > -1) {
    noteCell.setBackground('#c8e6c9').setFontColor('#1b5e20');
  } else if (noteStr.indexOf('⚠️') > -1) {
    noteCell.setBackground('#ffcdd2').setFontColor('#b71c1c');
  }

  // Payout Ratio — своя колонка, свой светофор. Только у акций (data.payoutColor
  // задан); у облигаций поле просто отсутствует, ячейка остаётся пустой.
  if (data.payoutColor) {
    sh.getRange(row, 8)
      .setBackground(data.payoutColor).setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center');
  }
}

function writeTotalRow_(sh, row, label, amount, COLS, color) {
  sh.getRange(row, 1, 1, COLS).merge()
    .setValue(label + ':   ' + rub_(amount))
    .setBackground(color || C.MID).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('right');
}

function writeConclusion_(sh, row, label, text) {
  sh.getRange(row, 1, 1, 2).merge()
    .setValue(label).setBackground('#eceff1')
    .setFontWeight('bold').setFontColor('#37474f');
  sh.getRange(row, 3, 1, 5).merge()
    .setValue(text).setBackground('#eceff1').setFontColor('#546e7a')
    .setWrap(true);
}