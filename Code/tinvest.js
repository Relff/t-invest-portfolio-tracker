/**
 * Tinkoff Invest → Google Sheets
 * Лоты, ₽, средняя цена/шт, разложения по типам, агрегация по всем счетам.
 *
 * Листы:
 * - Positions — сырые позиции (все счета как есть)
 * - Positions_Aggregated — агрегировано по AGGREGATE_BY (по всем счетам)
 * - Positions_Shares / Bonds / ETFs / Currencies / Futures / Other — типы, уже СУММИРОВАНО по всем счетам
 * - Positions_Money — деньги (RUB) по всем счетам (агрегировано)
 * - Positions_SummaryByType — свод сумм ₽ по категориям (включая Money)
 *
 * Перед запуском:
 *   Project Settings → Script properties → TINKOFF_TOKEN = t.xxxxx...
 */

///////////////////////
// Константы
///////////////////////
const TOKEN_PROPERTY_KEY = 'TINKOFF_TOKEN';
const BASES = [
  'https://invest-public-api.tbank.ru/rest',
  'https://invest-public-api.tinkoff.ru/rest', // fallback
];
const RAW_SHEET = 'Позиции'; // должно совпадать с DST.POSITIONS в dashboard.gs
const AGG_SHEET = 'Позиции_Сводка';
const SUMMARY_BY_TYPE_SHEET = 'Свод_по_типам';

const TYPE_SHEETS = {
  Shares:     'Дан_Акции',
  Bonds:      'Дан_Облигации',
  ETFs:       'Дан_Фонды',
  Currencies: 'Дан_Валюта',
  Futures:    'Дан_Фьючерсы',
  Other:      'Дан_Прочее',
  Money:      'Дан_Деньги',
};

const CACHE_TTL_SEC = 6 * 3600; // 6 часов кэш справочника

// Чем группировать сводку/листы типов: 'name' | 'ticker' | 'figi'
const AGGREGATE_BY = 'name';

///////////////////////
// HTTP и утилиты
///////////////////////
function tiFetch_(path, body) {
  const token = PropertiesService.getScriptProperties().getProperty(TOKEN_PROPERTY_KEY);
  if (!token) throw new Error('В Script properties нет ключа ' + TOKEN_PROPERTY_KEY);

  let lastErr = null;
  for (const base of BASES) {
    try {
      const resp = UrlFetchApp.fetch(base + path, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
        payload: JSON.stringify(body || {}),
        muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) return JSON.parse(resp.getContentText());
      if (code === 404) { lastErr = new Error('TINKOFF API 404 @ ' + base + path); continue; }
      throw new Error('TINKOFF API ' + code + ': ' + resp.getContentText());
    } catch (e) { lastErr = e; }
  }
  throw lastErr || new Error('Не удалось выполнить запрос к Invest API');
}

function qToNumber_(q) {
  if (!q) return 0;
  return Number(q.units || 0) + Number(q.nano || 0) / 1e9;
}

function moneyToNumber_(m) {
  if (!m) return 0;
  return Number(m.units || 0) + Number(m.nano || 0) / 1e9;
}

function pause_(ms) { Utilities.sleep(ms || 35); }

///////////////////////
// Справочник инструмента по FIGI
///////////////////////
function getInstrumentByFigi_(figi) {
  if (!figi) return null;
  const cache = CacheService.getScriptCache();
  const key = 'inst:' + figi;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const resp = tiFetch_(
    '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetInstrumentBy',
    { idType: 'INSTRUMENT_ID_TYPE_FIGI', id: figi }
  );

  const data = resp.instrument || resp.share || resp.bond || resp.etf || resp.future || null;
  const out = data ? {
    figi,
    lot: Number(data.lot || 1),
    ticker: data.ticker || '',
    name: data.name || data.issuerName || '',
    instrumentType: (data.instrumentType || data.type || '').toLowerCase(),
    currency: String(data.currency || data.nominalCurrency || '').toUpperCase(),
  } : null;

  if (out) cache.put(key, JSON.stringify(out), CACHE_TTL_SEC);
  return out;
}

///////////////////////
// Классификация типа → категория/лист
///////////////////////
function toCategory_(instrType) {
  const t = String(instrType || '').toLowerCase();
  if (t.includes('share')) return 'Shares';
  if (t === 'bond')        return 'Bonds';
  if (t === 'etf')         return 'ETFs';
  if (t === 'currency')    return 'Currencies';
  if (t === 'future' || t === 'futures') return 'Futures';
  if (t === 'money')       return 'Money'; // спец-метка
  return 'Other';
}

///////////////////////
// Запись таблицы (с лечением багов фильтра)
///////////////////////
function writeTable_(sheetName, rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // убираем старый фильтр, иначе "на листе можно создать только один фильтр"
  const existingFilter = sh.getFilter();
  if (existingFilter) existingFilter.remove();

  sh.clearContents();
  if (!rows.length) return;

  const range = sh.getRange(1, 1, rows.length, rows[0].length);
  range.setValues(rows);

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.autoResizeColumns(1, rows[0].length);

  if (rows.length > 1) {
    const header = rows[0];
    const n = rows.length - 1;
    const idx = Object.fromEntries(header.map((k, i) => [k, i + 1])); // 1-based

    // количества
    ['quantity_lots','quantity_pcs'].forEach(k => {
      if (idx[k]) sh.getRange(2, idx[k], n, 1).setNumberFormat('0.########');
    });
    // деньги
    ['current_price_rub_per_piece','avg_price_rub_per_piece','price_rub_per_lot','position_value_rub'].forEach(k => {
      if (idx[k]) sh.getRange(2, idx[k], n, 1).setNumberFormat('#,##0.00 [$₽-ru-RU]');
    });

    range.createFilter();
  }
}

///////////////////////
// Агрегация по ключу (name/ticker/figi)
///////////////////////
function aggregateByKey_(rows, keyName) {
  if (!rows || rows.length <= 1) return rows;
  const header = rows[0];
  const H = Object.fromEntries(header.map((k, i) => [k, i]));

  const map = new Map(); // key -> aggregator
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const key = String(r[H[keyName]] || '').trim();
    if (!key) continue;

    const qtyLots = Number(r[H['quantity_lots']] || 0);
    const qtyPcs  = Number(r[H['quantity_pcs']]  || 0);
    const curPrice = Number(r[H['current_price_rub_per_piece']] || 0);
    const avgPrice = Number(r[H['avg_price_rub_per_piece']] || 0);
    const posRub = Number(r[H['position_value_rub']] || 0);

    if (!map.has(key)) {
      map.set(key, {
        name: key,
        types: new Set([r[H['type']]]),
        lots: new Set([r[H['lot']]]),
        tickers: new Set([r[H['ticker']]]),
        currencies: new Set([r[H['instrument_currency']]]),
        qtyLotsSum: 0,
        qtyPcsSum: 0,
        posRubSum: 0,
        curWeightedSum: 0, curQtySum: 0,
        avgWeightedSum: 0, avgQtySum: 0,
      });
    }
    const ag = map.get(key);
    if (r[H['type']]) ag.types.add(r[H['type']]);
    if (r[H['lot']]) ag.lots.add(r[H['lot']]);
    if (r[H['ticker']]) ag.tickers.add(r[H['ticker']]);
    if (r[H['instrument_currency']]) ag.currencies.add(r[H['instrument_currency']]);

    ag.qtyLotsSum += qtyLots;
    ag.qtyPcsSum  += qtyPcs;
    ag.posRubSum  += posRub;

    if (curPrice > 0 && qtyPcs > 0) { ag.curWeightedSum += curPrice * qtyPcs; ag.curQtySum += qtyPcs; }
    if (avgPrice > 0 && qtyPcs > 0) { ag.avgWeightedSum += avgPrice * qtyPcs; ag.avgQtySum += qtyPcs; }
  }

  const out = [header.slice()];

  for (const [,ag] of map) {
    const lotConsensus = (ag.lots.size === 1) ? Number([...ag.lots][0]) : '';
    const instrType = (ag.types.size === 1) ? [...ag.types][0] : '';
    const tickerOne = (ag.tickers.size === 1) ? [...ag.tickers][0] : '';
    const currOne = (ag.currencies.size === 1) ? [...ag.currencies][0] : '';

    const curAvgPerPiece = ag.curQtySum > 0 ? (ag.curWeightedSum / ag.curQtySum) : 0;
    const avgAvgPerPiece = ag.avgQtySum > 0 ? (ag.avgWeightedSum / ag.avgQtySum) : 0;
    const pricePerLot = (lotConsensus && curAvgPerPiece) ? curAvgPerPiece * lotConsensus : '';

    const row = [];
    row[H['accountId']] = 'ALL';
    row[H['type']] = instrType || 'security';
    row[H['figi']] = '';
    row[H['ticker']] = tickerOne || '';
    row[H['name']] = ag.name;
    row[H['lot']] = lotConsensus || '';
    row[H['quantity_lots']] = ag.qtyLotsSum;
    row[H['quantity_pcs']]  = ag.qtyPcsSum;
    row[H['current_price_rub_per_piece']] = curAvgPerPiece;
    row[H['avg_price_rub_per_piece']] = avgAvgPerPiece || '';
    row[H['price_rub_per_lot']] = pricePerLot;
    row[H['position_value_rub']] = ag.posRubSum;
    row[H['instrument_currency']] = currOne || '';

    for (let i = 0; i < header.length; i++) if (row[i] === undefined) row[i] = '';
    out.push(row);
  }

  // сортировка по стоимости позиции
  const posIdx = header.indexOf('position_value_rub');
  out.splice(1, out.length - 1, ...out.slice(1).sort((a, b) => Number(b[posIdx]) - Number(a[posIdx])));
  return out;
}

///////////////////////
// Свод по типам (₽)
///////////////////////
function buildSummaryByType_(rows) {
  if (!rows || rows.length <= 1) return [['type','position_value_rub'],['EMPTY',0]];
  const header = rows[0];
  const H = Object.fromEntries(header.map((k, i) => [k, i]));
  const sums = new Map(); // type -> ₽

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const cat = toCategory_(r[H['type']]);
    const val = Number(r[H['position_value_rub']] || 0);
    sums.set(cat, (sums.get(cat) || 0) + val);
  }

  const out = [['type','position_value_rub']];
  for (const key of ['Shares','Bonds','ETFs','Currencies','Futures','Other']) {
    if (sums.has(key)) out.push([key, sums.get(key)]);
  }
  return out; // Money добавим отдельно (из листа Money)
}

///////////////////////
// Деньги (кэш) RUB
///////////////////////
function fetchMoneyRowsRUB_(accountId, header) {
  const H = Object.fromEntries(header.map((k, i) => [k, i]));
  const rows = [];
  const pos = tiFetch_('/tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions', { accountId });
  const money = pos.money || pos.securities?.money || [];

  const rubItem = (money || []).find(m => String(m.currency || (m.currencyIsoCode || '')).toUpperCase() === 'RUB');
  const rubAmount = rubItem ? moneyToNumber_(rubItem) : 0;

  if (rubAmount > 0) {
    const row = [];
    row[H['accountId']] = accountId;
    row[H['type']] = 'money';
    row[H['figi']] = '';
    row[H['ticker']] = 'RUB';
    row[H['name']] = 'Деньги (RUB)';
    row[H['lot']] = '';
    row[H['quantity_lots']] = '';
    row[H['quantity_pcs']]  = '';
    row[H['current_price_rub_per_piece']] = '';
    row[H['avg_price_rub_per_piece']] = '';
    row[H['price_rub_per_lot']] = '';
    row[H['position_value_rub']] = rubAmount; // ₽
    row[H['instrument_currency']] = 'RUB';
    for (let i = 0; i < header.length; i++) if (row[i] === undefined) row[i] = '';
    rows.push(row);
  }
  return rows;
}

///////////////////////
// Главная синхронизация
///////////////////////
function syncTinkoffPositions() {
  const header = [
    'accountId',
    'type',
    'figi',
    'ticker',
    'name',
    'lot',
    'quantity_lots',
    'quantity_pcs',
    'current_price_rub_per_piece',
    'avg_price_rub_per_piece',
    'price_rub_per_lot',
    'position_value_rub',
    'instrument_currency',
  ];
  const raw = [header];

  // 1) Аккаунты
  const accResp = tiFetch_('/tinkoff.public.invest.api.contract.v1.UsersService/GetAccounts', {});
  const accounts = (accResp.accounts || []).filter(a => a.status === 'ACCOUNT_STATUS_OPEN');

  // 2) Собираем все позиции по всем аккаунтам
  for (const acc of accounts) {
    const accountId = acc.id;

    // Портфель в ₽ (все бумаги)
    const pf = tiFetch_('/tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio',
                        { accountId, currency: 'RUB' });
    const positions = pf.positions || [];

    for (const p of positions) {
      const figi = p.figi || '';
      const info = figi ? getInstrumentByFigi_(figi) : null;

      const lot = Math.max(1, Number(info?.lot || 1));
      const qtyPcs = qToNumber_(p.quantity);

      const currentPriceMoney =
        p.currentPrice           ||
        p.current_price          || null;
      const averagePriceMoney =
        p.averagePositionPrice   ||
        p.average_position_price || null;

      const currentPriceRubPerPiece = moneyToNumber_(currentPriceMoney);
      const avgPriceRubPerPiece     = moneyToNumber_(averagePriceMoney);

      const qtyLots = lot ? (qtyPcs / lot) : qtyPcs;
      const priceRubPerLot = currentPriceRubPerPiece * lot;
      const positionRub = currentPriceRubPerPiece * qtyPcs;

      raw.push([
        accountId,
        info?.instrumentType || p.instrumentType || 'security',
        figi,
        info?.ticker || '',
        info?.name || '',
        lot,
        qtyLots,
        qtyPcs,
        currentPriceRubPerPiece,
        avgPriceRubPerPiece || '',
        priceRubPerLot,
        positionRub,
        info?.currency || ''
      ]);

      pause_(30);
    }

    // Деньги (RUB)
try {
  const moneyRows = fetchMoneyRowsRUB_(accountId, header);
  for (const r of moneyRows) raw.push(r);
} catch(e) {
  console.warn('Пропущен кэш для счёта ' + accountId + ': ' + e.message);
}
  }

  // 3) RAW
  writeTable_(RAW_SHEET, raw);

  // 4) Общая агрегированная сводка
  if (['name','ticker','figi'].includes(AGGREGATE_BY)) {
    const aggAll = aggregateByKey_(raw, AGGREGATE_BY);
    writeTable_(AGG_SHEET, aggAll);
  }

  // 5) Разложение по типам: сначала раскладываем, потом АГРЕГИРУЕМ внутри типа
  const perType = {
    Shares: [header.slice()],
    Bonds: [header.slice()],
    ETFs: [header.slice()],
    Currencies: [header.slice()],
    Futures: [header.slice()],
    Other: [header.slice()],
    Money: [header.slice()],
  };
  const H = Object.fromEntries(header.map((k, i) => [k, i]));

  for (let i = 1; i < raw.length; i++) {
    const r = raw[i];
    const isMoney = String(r[H['type']]).toLowerCase() === 'money';
    const cat = isMoney ? 'Money' : toCategory_(r[H['type']]);
    perType[cat].push(r);
  }

  const aggShares     = aggregateByKey_(perType.Shares,     AGGREGATE_BY);
  const aggBonds      = aggregateByKey_(perType.Bonds,      AGGREGATE_BY);
  const aggETFs       = aggregateByKey_(perType.ETFs,       AGGREGATE_BY);
  const aggCurrencies = aggregateByKey_(perType.Currencies, AGGREGATE_BY);
  const aggFutures    = aggregateByKey_(perType.Futures,    AGGREGATE_BY);
  const aggOther      = aggregateByKey_(perType.Other,      AGGREGATE_BY);
  const aggMoney      = aggregateByKey_(perType.Money,      AGGREGATE_BY); // сольётся в один RUB

  writeTable_(TYPE_SHEETS.Shares,     aggShares);
  writeTable_(TYPE_SHEETS.Bonds,      aggBonds);
  writeTable_(TYPE_SHEETS.ETFs,       aggETFs);
  writeTable_(TYPE_SHEETS.Currencies, aggCurrencies);
  writeTable_(TYPE_SHEETS.Futures,    aggFutures);
  writeTable_(TYPE_SHEETS.Other,      aggOther);
  writeTable_(TYPE_SHEETS.Money,      aggMoney);

  // 6) Свод по типам
  const summary = buildSummaryByType_(raw);
  const moneySum = aggMoney.slice(1).reduce((s, r) => s + Number(r[H['position_value_rub']] || 0), 0);
  if (moneySum > 0) summary.push(['Money', moneySum]);

  writeTable_(SUMMARY_BY_TYPE_SHEET, summary);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SUMMARY_BY_TYPE_SHEET);
  if (sh && summary.length > 1) {
    sh.getRange(2, 2, summary.length - 1, 1).setNumberFormat('#,##0.00 [$₽-ru-RU]');
    sh.autoResizeColumns(1, 2);
  }
}

///////////////////////
// Меню
///////////////////////


///////////////////////
// Утилита (разово): снести фильтры везде, если что-то залипло
///////////////////////
function dropAllFiltersOnce() {
  const ss = SpreadsheetApp.getActive();
  const sheetNames = [
    RAW_SHEET, AGG_SHEET, SUMMARY_BY_TYPE_SHEET,
    TYPE_SHEETS.Shares, TYPE_SHEETS.Bonds, TYPE_SHEETS.ETFs,
    TYPE_SHEETS.Currencies, TYPE_SHEETS.Futures, TYPE_SHEETS.Other, TYPE_SHEETS.Money
  ];
  for (const name of sheetNames) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    const f = sh.getFilter();
    if (f) f.remove();
  }
}

///////////////////////
// (опц) Самотест авторизации
///////////////////////
function testInfo() {
  const r = tiFetch_('/tinkoff.public.invest.api.contract.v1.UsersService/GetInfo', {});
  console.log(JSON.stringify(r, null, 2));
}
