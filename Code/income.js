/**
 * income_v4.js — Лист «Ожидаемый доход»
 *
 * Главное изменение vs v3:
 *   - Акции и облигации читаются НАПРЯМУЮ из Дан_Акции / Дан_Облигации
 *     (больше не зависит от категорийной логики readPositions_)
 *   - Исправлен баг: акции больше не попадают в «Прочие»
 *   - Улучшена строка «Остальные облигации» — показывает сумму купонов
 *   - Добавлен блок «Выводы» внизу
 */

const BOND_TOP_N = 15;

// ════════════════════════════════════════════════════════════════════
// ГЛАВНАЯ ФУНКЦИЯ
// ════════════════════════════════════════════════════════════════════

function updateIncomeSheet() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.INCOME);
  if (!sh) sh = ss.insertSheet(DST.INCOME);
  sh.clearContents();
  sh.clearFormats();

  let config;
  try { config = readConfig_(); }
  catch(e) { sh.getRange(1,1).setValue('⚠️ ' + e.message); return; }

  let totalRub  = getTotalPortfolioValue_();
  let tz        = Session.getScriptTimeZone();
  let now       = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
  let divMap    = readDividendsFromConfig_();
  let figiMap   = buildFigiMap_();

  // Читаем акции и облигации напрямую
  let sharePositions = readSheetPositions_('Дан_Акции');
  let bondPositions  = readSheetPositions_('Дан_Облигации');

  // Считаем доход
  let shareRows = calcShareIncome_(sharePositions, divMap, totalRub);
  let bondRows  = calcBondIncome_(bondPositions, figiMap, totalRub);

  // Сортировка
  shareRows.sort(function(a,b){ return a.name.localeCompare(b.name,'ru'); });
  bondRows.sort(function(a,b){ return b.incomeYear - a.incomeYear; });

  // Итоги
  let shareTotal = shareRows.reduce(function(s,r){ return s + r.incomeYear; }, 0);
  let bondTotal  = bondRows.reduce(function(s,r){ return s + r.incomeYear; }, 0);
  let grandTotal = shareTotal + bondTotal;
  let yieldPct   = totalRub > 0 ? (grandTotal / totalRub * 100) : 0;

  let COLS = 7;
  let r    = 1;

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
  let half = Math.floor(COLS / 2);
  let tiles = [
    { label: '📅 Доход в год',   val: rub_(grandTotal), bg: '#1565c0' },
    { label: '📈 Доходность',    val: yieldPct.toFixed(1) + '%', bg: '#1565c0' },
    { label: '📊 Акции',        val: rub_(shareTotal), bg: C.MID },
    { label: '🏦 Облигации',    val: rub_(bondTotal),  bg: C.MID },
  ];
  for (let ti = 0; ti < tiles.length; ti++) {
    let col = (ti % 2 === 0) ? 1 : half + 1;
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
    ['Название','Тикер','Кол-во','Дивиденд/акц, ₽','Доход в год, ₽','% порт.','Примечание'],
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
    ['Название','Тикер','Кол-во','Купон/год на шт, ₽','Доход в год, ₽','% порт.','Купонов'],
    COLS);
  r++;

  let topBonds  = bondRows.slice(0, BOND_TOP_N);
  let restBonds = bondRows.slice(BOND_TOP_N);

  topBonds.forEach(function(row, idx) {
    writeIncomeRow_(sh, r, row, idx, totalRub);
    r++;
  });

  // Строка «Остальные» с суммой купонов
  if (restBonds.length > 0) {
    let restIncome = restBonds.reduce(function(s,x){ return s + x.incomeYear; }, 0);
    let restValue  = restBonds.reduce(function(s,x){ return s + x.valueRub;   }, 0);
    let restPct    = totalRub > 0 ? restValue / totalRub : 0;
    let restBg     = topBonds.length % 2 === 0 ? C.EVEN : C.ODD;

    sh.getRange(r, 1, 1, COLS).setValues([[
      '…ещё ' + restBonds.length + ' облигаций',
      '', '', '',
      restIncome, restPct,
      'сумма купонов: ' + rub_(restIncome)
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
  let top3 = bondRows.slice(0, 3);
  let top3txt = top3.map(function(b, i) {
    return (i+1) + '. ' + b.name + ' — ' + rub_(b.incomeYear);
  }).join('   |   ');

  writeConclusion_(sh, r, '🏆 Топ-3 облигации по доходу:', top3txt); r++;

  // Акции с дивидендами
  let divStocks = shareRows.filter(function(x){ return x.incomeYear > 0; });
  let divTxt = divStocks.length > 0
    ? divStocks.map(function(s){ return s.name + ' (' + rub_(s.incomeYear) + ')'; }).join(', ')
    : 'Пока нет дивидендных акций';
  writeConclusion_(sh, r, '💰 Дивиденды по акциям:', divTxt); r++;

  // Структура дохода
  let bondShare = grandTotal > 0 ? Math.round(bondTotal / grandTotal * 100) : 0;
  writeConclusion_(sh, r, '📊 Структура дохода:',
    'Облигации ' + bondShare + '%  ·  Акции ' + (100-bondShare) + '%'); r++;

  // Ежемесячный доход
  writeConclusion_(sh, r, '📅 Ежемесячно в среднем:',
    rub_(grandTotal / 12) + ' (~' + rub_(bondTotal / 12) + ' от купонов + ' +
    rub_(shareTotal / 12) + ' от дивидендов)'); r++;

  // Краткий вывод
  let conclusion = grandTotal === 0
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

  // ── Ширина колонок ────────────────────────────────────────────────
  [255, 65, 75, 135, 155, 80, 200].forEach(function(w,i){
    sh.setColumnWidth(i+1, w);
  });
  sh.setFrozenRows(6);
  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════════════
// ЧТЕНИЕ ПОЗИЦИЙ НАПРЯМУЮ ИЗ ЛИСТОВ
// ════════════════════════════════════════════════════════════════════

function readSheetPositions_(sheetName) {
  let ss   = SpreadsheetApp.getActive();
  let sh   = ss.getSheetByName(sheetName);
  if (!sh) return [];

  let data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  let H = {};
  data[0].forEach(function(h,i){ H[String(h).trim()] = i; });

  let result = [];
  for (let i = 1; i < data.length; i++) {
    let row      = data[i];
    let name     = String(row[H['name']]   || '').trim();
    let ticker   = String(row[H['ticker']] || '').trim();
    let valueRub = Number(row[H['position_value_rub']]          || 0);
    let price    = Number(row[H['current_price_rub_per_piece']] || 0);
    let qty      = Number(row[H['quantity_pcs']]                || 0);
    let lot      = Number(row[H['lot']]    || 1);
    let figi     = String(row[H['figi']]   || '').trim();

    if (!name || valueRub === 0) continue;
    result.push({ name:name, ticker:ticker, valueRub:valueRub,
                  price:price, qty:qty, lot:lot, figi:figi });
  }
  return result;
}

function getTotalPortfolioValue_() {
  let ss  = SpreadsheetApp.getActive();
  let src = ['Дан_Акции','Дан_Облигации','Дан_Фонды','Дан_Деньги','Дан_Валюта'];
  let total = 0;
  src.forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) return;
    let data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    let H = {};
    data[0].forEach(function(h,i){ H[String(h).trim()] = i; });
    for (let i = 1; i < data.length; i++) {
      total += Number(data[i][H['position_value_rub']] || 0);
    }
  });
  return total;
}


// ════════════════════════════════════════════════════════════════════
// РАСЧЁТ ДОХОДА
// ════════════════════════════════════════════════════════════════════

function calcShareIncome_(positions, divMap, totalRub) {
  return positions.map(function(p) {
    let div           = findDividend_(p.name, p.ticker, divMap);
    let incomePerUnit = div.amount;
    let incomeYear    = incomePerUnit * p.qty;
    return {
      name:          p.name,
      ticker:        p.ticker,
      qty:           p.qty,
      incomePerUnit: incomePerUnit,
      incomeYear:    incomeYear,
      valueRub:      p.valueRub,
      note:          div.note,
    };
  });
}

function calcBondIncome_(positions, figiMap, totalRub) {
  return positions.map(function(p) {
    let figi = p.figi || figiMap[p.name] || figiMap[p.ticker] || '';
    let coupon = { perUnit: 0, note: 'FIGI не найден' };
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

function fetchAnnualCoupon_(figi) {
  if (!figi) return { perUnit: 0, note: 'Нет FIGI' };
  try {
    let now  = new Date();
    let in1y = new Date(now.getTime() + 365*24*3600*1000);
    let resp = tiFetch_(
      '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetBondCoupons',
      { figi: figi, from: now.toISOString(), to: in1y.toISOString() }
    );
    let coupons = resp.events || [];
    let total   = 0;
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
// ДИВИДЕНДЫ ИЗ CONFIG
// ════════════════════════════════════════════════════════════════════

function readDividendsFromConfig_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) return {};
  let v   = sh.getDataRange().getValues();
  let map = {};
  let startRow = -1;
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).includes('ДИВИДЕНДЫ')) { startRow = i + 2; break; }
  }
  if (startRow < 0) return {};
  for (let j = startRow; j < v.length; j++) {
    let name   = String(v[j][0]).trim();
    let amount = Number(v[j][1]) || 0;
    let ticker = String(v[j][2]).trim();
    if (!name) break;
    if (name)   map[name]   = { amount: amount, note: 'из Config' };
    if (ticker) map[ticker] = { amount: amount, note: 'из Config' };
  }
  return map;
}

function findDividend_(name, ticker, divMap) {
  if (divMap[name])   return divMap[name];
  if (divMap[ticker]) return divMap[ticker];
  let nl = name.toLowerCase();
  let keys = Object.keys(divMap);
  for (let i = 0; i < keys.length; i++) {
    let k = keys[i].toLowerCase();
    if (nl.includes(k) || k.includes(nl)) return divMap[keys[i]];
  }
  return { amount: 0, note: '⚠️ Укажите в Config → Блок 4' };
}


// ════════════════════════════════════════════════════════════════════
// FIGI ИЗ ЛИСТА ПОЗИЦИИ
// ════════════════════════════════════════════════════════════════════

function buildFigiMap_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.POSITIONS);
  if (!sh) return {};
  let data = sh.getDataRange().getValues();
  if (data.length < 2) return {};
  let H = {};
  data[0].forEach(function(h,i){ H[String(h).trim()] = i; });
  let map = {};
  for (let i = 1; i < data.length; i++) {
    let name   = String(data[i][H['name']]   || '').trim();
    let ticker = String(data[i][H['ticker']] || '').trim();
    let figi   = String(data[i][H['figi']]   || '').trim();
    if (figi && name)   map[name]   = figi;
    if (figi && ticker) map[ticker] = figi;
  }
  return map;
}


// ════════════════════════════════════════════════════════════════════
// БЛОК 4 В CONFIG
// ════════════════════════════════════════════════════════════════════

function addDividendsBlock() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CONFIG);
  if (!sh) { SpreadsheetApp.getUi().alert('⚠️ Сначала запустите initConfig()'); return; }

  let vals = sh.getDataRange().getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).includes('ДИВИДЕНДЫ')) {
      SpreadsheetApp.getUi().alert('Блок 4 уже существует в Config.');
      return;
    }
  }

  let lastRow = sh.getLastRow() + 2;
  let block = [
    ['▌ ДИВИДЕНДЫ АКЦИЙ — заполните вручную (₽ в год на 1 акцию)', '', ''],
    ['Название (точно как в портфеле)', 'Дивиденд ₽/год', 'Тикер'],
    ['Сбербанк',                 34,    'SBER'],
    ['Лукойл',                   1000,  'LKOH'],
    ['Т-Технологии',             0,     'T'],
    ['ЯНДЕКС ао01',              0,     'YDEX'],
    ['Корпоративный Центр Икс 5',0,     'X5'],
    ['ФосАгро',                  900,   'PHOR'],
    ['Positive Technologies',    0,     'POSI'],
    ['Озон',                     0,     'OZON'],
    ['Полюс',                    700,   'PLZL'],
  ];

  sh.getRange(lastRow, 1, block.length, 3).setValues(block);
  sh.getRange(lastRow, 1, 1, 3).merge()
    .setBackground(C.DARK).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.getRange(lastRow+1, 1, 1, 3)
    .setBackground(C.MID).setFontColor('#ffffff').setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    '✅ Блок 4 добавлен!\n\n' +
    'Проверьте суммы дивидендов — они примерные.\n' +
    'Т-Техно, Яндекс, X5, Ozon, Positive → дивидендов нет, оставьте 0.'
  );
}


// ════════════════════════════════════════════════════════════════════
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ════════════════════════════════════════════════════════════════════

function writeIncomeRow_(sh, row, data, idx, totalRub) {
  let bg     = idx % 2 === 0 ? C.EVEN : C.ODD;
  let pctPf  = (totalRub > 0 && data.valueRub > 0) ? data.valueRub / totalRub : 0;
  sh.getRange(row, 1, 1, 7).setValues([[
    data.name, data.ticker, data.qty,
    data.incomePerUnit, data.incomeYear,
    pctPf, data.note
  ]]).setBackground(bg);

  sh.getRange(row, 3).setNumberFormat('0');
  sh.getRange(row, 4).setNumberFormat('#,##0.00 [$₽-ru-RU]');
  sh.getRange(row, 6).setNumberFormat('0.0%');

  let incCell = sh.getRange(row, 5);
  incCell.setNumberFormat('#,##0 [$₽-ru-RU]');
  let inc = data.incomeYear;
  if      (inc <= 0)    incCell.setBackground('#ffcdd2').setFontColor('#b71c1c');
  else if (inc < 1000)  incCell.setBackground('#fff9c4').setFontColor('#f57f17');
  else if (inc < 5000)  incCell.setBackground('#c8e6c9').setFontColor('#1b5e20');
  else                  incCell.setBackground('#2e7d32').setFontColor('#ffffff');
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
    .setValue(text).setBackground('#eceff1').setFontColor('#546e7a');
}
