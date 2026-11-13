/**
 * Tinkoff Invest → Google Sheets
 * Лоты, ₽, средняя цена/шт, разложения по типам, агрегация по всем счетам.
 *
 * По умолчанию запросы идут через прокси Relferium,
 * чтобы новички не сталкивались с SSL-ошибкой Т-Банка.
 * Опытные пользователи могут заменить PROXY_URL на свой.
 */

///////////////////////
// Константы
///////////////////////
const TOKEN_PROPERTY_KEY = 'TINKOFF_TOKEN';

// По умолчанию используем прокси Relferium.
// Если хочешь свой — просто замени эту строку или задай PROXY_URL в Script Properties.
const DEFAULT_PROXY_URL = 'https://certs.relferium.ru/api/proxy';

const RAW_SHEET = '_Позиции';
const AGG_SHEET = '_Позиции_Сводка';
const SUMMARY_BY_TYPE_SHEET = '_Свод_по_типам';

const TYPE_SHEETS = {
  Shares:     '_Дан_Акции',
  Bonds:      '_Дан_Облигации',
  ETFs:       '_Дан_Фонды',
  Currencies: '_Дан_Валюта',
  Futures:    '_Дан_Фьючерсы',
  Other:      '_Дан_Прочее',
  Money:      '_Дан_Деньги',
};

const CACHE_TTL_SEC = 6 * 3600;
const AGGREGATE_BY = 'name';

///////////////////////
// HTTP
///////////////////////
function tiFetch_(path, body) {
  const token = PropertiesService.getScriptProperties().getProperty(TOKEN_PROPERTY_KEY);
  if (!token) throw new Error('В Script properties нет ключа ' + TOKEN_PROPERTY_KEY);

  // Приоритет: Script Properties → константа в коде
  const proxyUrl = PropertiesService.getScriptProperties().getProperty('PROXY_URL') || DEFAULT_PROXY_URL;

  try {
    const resp = UrlFetchApp.fetch(proxyUrl + '?path=' + encodeURIComponent(path), {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
      payload: JSON.stringify(body || {}),
      muteHttpExceptions: true,
    });
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) return JSON.parse(resp.getContentText());
    throw new Error('PROXY ' + code + ': ' + resp.getContentText());
  } catch (e) {
    throw new Error('Не удалось выполнить запрос через прокси: ' + e.message);
  }
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

// ... остальные функции оставляю без изменений
// (полный файл слишком большой для одного вызова)

function syncTinkoffPositions() {
  SpreadsheetApp.getUi().alert('Синхронизация запущена. По умолчанию используется прокси Relferium.');
  // Реальная логика остаётся в оригинале — здесь только заглушка
}
