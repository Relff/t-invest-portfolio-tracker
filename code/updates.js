/* Produced by Relferium — t-invest-portfolio-tracker */

///////////////////////
// Проверка обновлений
///////////////////////

// Текущая версия трекера. Читается как дефолт для ячейки версии на листе
// Dashboard (см. dashboard.js) и как база для сравнения с GitHub-релизами.
const TRACKER_VERSION_DEFAULT = 'v1.0';

const GITHUB_REPO_ = 'Relff/t-invest-portfolio-tracker';

/**
 * Разбирает строку версии вида "v1.2.3" / "1.2" в массив чисел [1, 2, 3].
 * Нечисловые/отсутствующие сегменты трактуются как 0.
 */
function parseVersion_(v) {
  return String(v || '')
    .replace(/^v/i, '')
    .split('.')
    .map((part) => parseInt(part, 10) || 0);
}

/**
 * true, если версия remote строго новее версии local (сравнение по
 * сегментам числами, а не строками — чтобы "v1.10" считалась новее "v1.9").
 */
function isNewerVersion_(remote, local) {
  const r = parseVersion_(remote);
  const l = parseVersion_(local);
  const len = Math.max(r.length, l.length);

  for (let i = 0; i < len; i++) {
    const rv = r[i] || 0;
    const lv = l[i] || 0;
    if (rv > lv) return true;
    if (rv < lv) return false;
  }
  return false;
}

/**
 * Проверяет на GitHub, есть ли релиз новее TRACKER_VERSION_DEFAULT.
 *
 * Запрашивает https://api.github.com/repos/Relff/t-invest-portfolio-tracker/releases/latest,
 * сравнивает `tag_name` из ответа с TRACKER_VERSION_DEFAULT.
 *
 * Это вспомогательная, необязательная для работы трекера проверка — если
 * GitHub недоступен, лимит анонимных запросов API исчерпан (60/час на IP),
 * репозиторий ещё без единого релиза (404) или ответ пришёл в неожиданном
 * виде, функция тихо возвращает { hasUpdate: false } и не бросает
 * исключение, чтобы ежедневная синхронизация (syncAndRefresh) никогда не
 * падала из-за стороннего сервиса.
 *
 * @returns {{hasUpdate: boolean, latestVersion?: string, url?: string}}
 */
function checkForUpdates_() {
  try {
    const resp = UrlFetchApp.fetch(
      'https://api.github.com/repos/' + GITHUB_REPO_ + '/releases/latest',
      {
        method: 'get',
        headers: { Accept: 'application/vnd.github+json' },
        muteHttpExceptions: true,
      },
    );

    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      return { hasUpdate: false };
    }

    const data = JSON.parse(resp.getContentText());
    const latestVersion = data.tag_name;
    if (!latestVersion) {
      return { hasUpdate: false };
    }

    if (isNewerVersion_(latestVersion, TRACKER_VERSION_DEFAULT)) {
      return {
        hasUpdate: true,
        latestVersion,
        url: data.html_url || ('https://github.com/' + GITHUB_REPO_ + '/releases/latest'),
      };
    }

    return { hasUpdate: false };
  } catch (e) {
    return { hasUpdate: false };
  }
}
