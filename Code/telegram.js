/**
 * telegram.js — Уведомления в Telegram
 *
 * Децентрализованная архитектура: каждый пользователь заводит своего
 * собственного бота через @BotFather и хостит его сам — общего сервера нет,
 * сообщения идут напрямую из Google Apps Script пользователя в Telegram API.
 *
 * Настройка (Свойства скрипта):
 *   TELEGRAM_BOT_TOKEN — токен бота от @BotFather
 *   TELEGRAM_CHAT_ID   — chat_id, куда присылать сообщения (см. INSTALLATION.md)
 *
 * Защита от двойного нажатия/повторной отправки — см. isDebouncedAction_()
 * ниже: одно и то же действие подряд в течение 10 сек обрабатывается только
 * один раз, повторные нажатия тихо игнорируются.
 *
 * Чтобы диалог не захламлялся, сообщения бота редактируются на месте
 * (sendTg_), а «служебные» сообщения человека — введённая сумма, список
 * тикеров для пропуска — удаляются сразу после обработки (deleteTgMessage_).
 * При возврате в меню весь «рабочий блок» предыдущей сессии (например
 * все шаги ребалансировки) удаляется одним махом — см. flushBlockMessages_().
 */

function sendTelegramMessage_(text) {
  let props   = PropertiesService.getScriptProperties();
  let token   = props.getProperty('TELEGRAM_BOT_TOKEN');
  let chatId  = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token || !chatId) {
    Logger.log('Telegram не настроен — заполните TELEGRAM_BOT_TOKEN и TELEGRAM_CHAT_ID в Свойствах скрипта.');
    return false;
  }

  let url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  let payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'HTML',
  };

  try {
    let resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    let code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log('Telegram API вернул код ' + code + ': ' + resp.getContentText());
      return false;
    }
    return true;
  } catch (e) {
    Logger.log('Ошибка отправки в Telegram: ' + e.message);
    return false;
  }
}

function testTelegramConnection() {
  let ok = sendTelegramMessage_(
    '✅ <b>Бот подключён и работает.</b>\nЭто тестовое сообщение от t-invest-portfolio-tracker.'
  );
  let ui = SpreadsheetApp.getUi();
  if (ok) {
    ui.alert('✅ Сообщение отправлено — проверьте Telegram.');
  } else {
    ui.alert('⚠️ Не удалось отправить сообщение. Проверьте TELEGRAM_BOT_TOKEN и TELEGRAM_CHAT_ID в Свойствах скрипта, и что вы нажали Start в чате с ботом.');
  }
}

/**
 * Уведомление об изменении статуса по правилу 5/25 (только при изменении).
 * Сравнивает текущий статус каждой категории (ОК/Умеренно/Требует внимания)
 * с тем, что было при прошлой проверке — присылает сообщение, только если
 * что-то реально поменялось, а не каждый раз, пока отклонение сохраняется.
 * Использует ту же deviationStatus_() из dashboard.js, что и сам дашборд —
 * статусы совпадают один в один.
 */
const DEVIATION_STATUS_PROP = 'DEVIATION_STATUS_LAST';

function checkAndNotifyDeviations_() {
  let config;
  try { config = readConfig_(); }
  catch (e) { return; }

  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);
  if (totalRub <= 0) return;

  let cats = Object.keys(config.classTargets);
  let currentStatus = {};
  cats.forEach(function(cat) {
    let actual = positions.filter(function(p) { return p.category === cat; })
                          .reduce(function(s, p) { return s + p.valueRub; }, 0);
    let actPct = actual / totalRub;
    let tgtPct = config.classTargets[cat] || 0;
    currentStatus[cat] = deviationStatus_(actPct, tgtPct).txt;
  });

  let props    = PropertiesService.getScriptProperties();
  let prevRaw  = props.getProperty(DEVIATION_STATUS_PROP);
  let prevStatus = prevRaw ? JSON.parse(prevRaw) : {};

  let changes = [];
  cats.forEach(function(cat) {
    if (prevStatus[cat] !== currentStatus[cat]) {
      changes.push(cat + ': ' + (prevStatus[cat] || 'нет данных') + ' → ' + currentStatus[cat]);
    }
  });

  props.setProperty(DEVIATION_STATUS_PROP, JSON.stringify(currentStatus));

  if (changes.length === 0) return;

  sendTelegramMessage_(
    '📊 <b>Статус по правилу 5/25 изменился</b>\n\n' + changes.map(function(c) { return '• ' + c; }).join('\n')
  );
}

/**
 * Уведомление о завершении ежедневной синхронизации.
 * Вызывается из syncAndRefresh() в dashboard.js. Если Telegram не настроен —
 * sendTelegramMessage_ тихо ничего не делает, ошибок не будет.
 */
/**
 * Подсказка про вывод дивидендов на карту — от остатка лимита вычета ИИС-3.
 * Две ситуации:
 *  1. Лимит ещё не выбран, а до конца года осталось мало времени — подсказка
 *     «пора включить вывод дивидендов на карту и занести их обратно как
 *     пополнение», чтобы успеть добить лимит до дедлайна.
 *  2. Лимит уже полностью выбран — подсказка «можно выключить», раз смысла
 *     больше нет (дивиденды снова спокойно капают внутрь ИИС).
 *
 * Каждое из двух сообщений шлётся МАКСИМУМ ОДИН РАЗ за календарный год —
 * иначе при ежедневном триггере бот присылал бы одно и то же каждый день.
 * Тихо ничего не делает, если ИИС-3 не настроен в Config — это опциональная
 * фича, а не обязательная часть синхронизации.
 */
const IIS_DIVIDEND_HINT_DAYS_THRESHOLD = 60; // за сколько дней до 31 декабря начинать подсказывать

function checkIisDividendHint_() {
  let data = computeIisDeductionData_();
  if (!data) return; // ИИС-3 не настроен в Config — тихо ничего не делаем

  let props = PropertiesService.getScriptProperties();
  let year  = data.year;

  if (data.remaining > 0 && data.daysLeft <= IIS_DIVIDEND_HINT_DAYS_THRESHOLD) {
    let flagKey = 'IIS_DIV_HINT_ENABLE_' + year;
    if (!props.getProperty(flagKey)) {
      sendTelegramMessage_(
        '💡 <b>Подсказка по ИИС-3</b>\n\n' +
        'До конца года осталось ' + data.daysLeft + ' дн., а лимит вычета ещё не выбран: ' +
        rub_(data.remaining) + ' из ' + rub_(IIS_DEDUCTION_LIMIT) + '.\n\n' +
        'Если пользуешься схемой «вывод дивидендов на карту → занести обратно как пополнение» — ' +
        'самое время её включить, чтобы успеть добить лимит.'
      );
      props.setProperty(flagKey, 'sent');
    }
  }

  if (data.remaining <= 0) {
    let flagKey = 'IIS_DIV_HINT_DISABLE_' + year;
    if (!props.getProperty(flagKey)) {
      sendTelegramMessage_(
        '✅ <b>Лимит вычета ИИС-3 выбран</b>\n\n' +
        'За ' + year + ' год довнесено ' + rub_(data.contributed) + ' из ' + rub_(IIS_DEDUCTION_LIMIT) + '.\n\n' +
        'Если включал вывод дивидендов на карту под эту схему — можно выключать, дальше смысла в нём нет.'
      );
      props.setProperty(flagKey, 'sent');
    }
  }
}

function notifySyncComplete_() {
  let config;
  try { config = readConfig_(); }
  catch (e) { return; }

  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);

  let tz  = Session.getScriptTimeZone();
  let now = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');

  sendTelegramMessage_(
    '🔄 <b>Синхронизация завершена</b>\n' +
    now + '\n\n' +
    '💼 Портфель: ' + rub_(totalRub)
  );
}

/**
 * Еженедельное уведомление о выплатах на ближайшие 7 дней.
 * Читает уже готовую таблицу «Все выплаты» из листа «Календарь выплат»
 * (она обновляется ежедневно через syncAndRefresh → updateCalendarSheet) —
 * заново дёргать T-Invest API здесь не нужно.
 *
 * Требует ОТДЕЛЬНЫЙ еженедельный триггер (Триггеры → Добавить триггер →
 * функция notifyWeeklyPayments → По времени → По неделям, в любой удобный
 * день недели), а не привязку к ежедневной синхронизации — частота другая.
 */
function notifyWeeklyPayments() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CALENDAR);
  if (!sh) return;

  let data = sh.getDataRange().getValues();
  let headerRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).indexOf('ВСЕ ВЫПЛАТЫ') >= 0) { headerRow = i; break; }
  }
  if (headerRow < 0) return;

  let now       = new Date();
  let weekLater = new Date(now.getTime() + 7 * 24 * 3600 * 1000);

  let couponSum = 0, dividendSum = 0;
  for (let i = headerRow + 2; i < data.length; i++) {
    let row  = data[i];
    let date = row[0];
    if (!(date instanceof Date)) continue;
    if (date < now || date > weekLater) continue;

    let type  = String(row[2]);
    let total = Number(row[5]) || 0;
    if (type.indexOf('Купон') === 0) couponSum += total; else dividendSum += total;
  }

  let totalSum = couponSum + dividendSum;
  sendTelegramMessage_(
    '📅 <b>Выплаты на неделю вперёд</b>\n\n' +
    '🪙 Купоны: ' + rub_(couponSum) + '\n' +
    '💵 Дивиденды: ' + rub_(dividendSum) + '\n\n' +
    '<b>Итого: ' + rub_(totalSum) + '</b>'
  );
}


// ════════════════════════════════════════════════════════════════════
// ДВУСТОРОННИЙ БОТ — меню, расчёт ребалансировки по команде
// ════════════════════════════════════════════════════════════════════
//
// СЕЙЧАС ИСПОЛЬЗУЕТСЯ POLLING (см. checkTelegramUpdates / pollOnce_ ниже) —
// это доставляет команды боту через триггер «раз в минуту», без деплоя
// и без вебхука. doPost() и setTelegramWebhook() оставлены в резерве —
// доставка мгновенная, но требует деплоя веб-приложения и аккуратной
// синхронизации URL (см. INSTALLATION.md, там же — как переключиться
// обратно на вебхук, если задержка polling когда-нибудь станет мешать).
//
// Из соображений приватности: бот отвечает ТОЛЬКО в чат с chat_id,
// совпадающим с TELEGRAM_CHAT_ID из Свойств скрипта — даже если кто-то
// узнает URL веб-приложения, писать боту от чужого лица не получится.

function setTelegramWebhook() {
  let token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) { SpreadsheetApp.getUi().alert('⚠️ Сначала заполните TELEGRAM_BOT_TOKEN.'); return; }

  let deployUrl = ScriptApp.getService().getUrl();
  if (!deployUrl) {
    SpreadsheetApp.getUi().alert('⚠️ Проект ещё не развёрнут как веб-приложение. Deploy → New deployment → Web app, затем повтори.');
    return;
  }

  let resp = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + token + '/setWebhook',
    { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ url: deployUrl }), muteHttpExceptions: true }
  );
  let text = resp.getContentText();
  Logger.log('setTelegramWebhook: ' + text);
  try { SpreadsheetApp.getUi().alert('Ответ Telegram: ' + text); }
  catch (e) { /* запущено не из таблицы — смотри Журнал выполнения */ }
}

function doPost(e) {
  let chatIdForError = null;
  try {
    let update = JSON.parse(e.postData.contents);

    if (update.update_id && isDuplicateUpdate_(update.update_id)) {
      return ContentService.createTextOutput('ok'); // уже обрабатывали это событие — Telegram прислал повторно
    }

    if (update.callback_query) {
      let cq     = update.callback_query;
      let chatId = cq.message.chat.id;
      chatIdForError = chatId;
      tgApiCall_('answerCallbackQuery', { callback_query_id: cq.id });
      if (isAuthorizedChat_(chatId) && !isDebouncedAction_(chatId, 'CB:' + cq.data)) {
        routeCallback_(chatId, cq.data);
      }

    } else if (update.message) {
      let msg    = update.message;
      let chatId = msg.chat.id;
      chatIdForError = chatId;
      let text = (msg.text || '').trim();
      if (isAuthorizedChat_(chatId) && !isDebouncedAction_(chatId, 'MSG:' + text)) {
        routeMessage_(chatId, text, msg.message_id);
      }
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    if (chatIdForError && isAuthorizedChat_(chatIdForError)) {
      try {
        sendTg_(chatIdForError, '⚠️ Что-то пошло не так:\n' + err.message, mainMenuKeyboard_(), { topicKey: 'ERROR' });
      } catch (e2) { /* даже это не сработало — сдаёмся молча */ }
    }
  }
  return ContentService.createTextOutput('ok');
}

/**
 * Защита от повторной обработки: Telegram может доставить одно и то же
 * обновление (update_id) больше одного раза — например, если веб-хук
 * не ответил быстро в первый раз. Запоминаем последний обработанный
 * update_id и пропускаем всё, что не новее него.
 */
function isDuplicateUpdate_(updateId) {
  let props = PropertiesService.getScriptProperties();
  let last  = Number(props.getProperty('TG_LAST_UPDATE_ID') || 0);
  if (updateId <= last) return true;
  props.setProperty('TG_LAST_UPDATE_ID', String(updateId));
  return false;
}

function isAuthorizedChat_(chatId) {
  let allowed = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
  return String(chatId) === String(allowed);
}

/**
 * Защита от двойного нажатия / повторной отправки одного и того же
 * действия подряд. Отличается от isDuplicateUpdate_() выше: там —
 * защита от того, что Telegram сам повторно доставил один и тот же
 * update_id; здесь — защита от того, что человек два раза ткнул одну
 * и ту же кнопку (или дважды отправил одно и то же сообщение), пока
 * не увидел ответа. Окно намеренно широкое (10 сек), потому что при
 * polling ответ и так может идти до 10 сек — тут-то обычно и тянет
 * нажать ещё раз.
 *
 * key — произвольная строка, уникальная для конкретного действия
 * (например 'CB:REBALANCE' для кнопки или 'MSG:45000' для сообщения).
 */
const TG_DEBOUNCE_MS = 10000;

function isDebouncedAction_(chatId, key) {
  let props   = PropertiesService.getScriptProperties();
  let propKey = 'TG_LAST_ACTION_' + chatId;
  let now     = new Date().getTime();

  let raw = props.getProperty(propKey);
  if (raw) {
    let last = JSON.parse(raw);
    if (last.key === key && (now - last.ts) < TG_DEBOUNCE_MS) return true;
  }
  props.setProperty(propKey, JSON.stringify({ key: key, ts: now }));
  return false;
}

// ════════════════════════════════════════════════════════════════════
// POLLING — альтернатива веб-хуку, без деплоя и без URL вообще
// ════════════════════════════════════════════════════════════════════
//
// Проще в настройке: не нужен Deploy → Web app, не нужен setWebhook,
// не нужно бояться, что URL развёртывания «уедет». Плата — задержка
// ответа до 1 минуты (периодичность триггера), вместо мгновенной.
//
// Порядок подключения:
//   1. Разово запусти removeTelegramWebhook() из редактора — снимает
//      старый вебхук (Telegram не даёт опрашивать, пока вебхук активен)
//   2. Добавь триггер: функция checkTelegramUpdates, По времени,
//      По минутам, каждую 1 минуту
//   3. doPost() и всё, что связано с веб-хуком, можно больше не трогать —
//      оставлены в коде на случай, если решишь вернуться к нему

function removeTelegramWebhook() {
  let resp = tgApiCall_('deleteWebhook', {});
  let text = resp ? resp.getContentText() : 'нет ответа';
  Logger.log('removeTelegramWebhook: ' + text);
  try { SpreadsheetApp.getUi().alert('Ответ Telegram: ' + text); }
  catch (e) { /* запущено не из таблицы — alert недоступен, смотри Журнал выполнения (Logger.log выше) */ }
}

/**
 * Диагностика: показывает, есть ли сейчас активный вебхук у бота и куда
 * он указывает. Если webhook активен — Telegram НЕ отдаёт обновления через
 * getUpdates (которым пользуется polling/checkTelegramUpdates), и бот
 * молчит на любые команды, даже без единой ошибки в логах. Всегда смотри
 * результат через Журнал выполнения (Logger.log) — ui.alert() не работает
 * при запуске напрямую из редактора Apps Script, только из меню таблицы.
 */
function getTelegramWebhookInfo() {
  let resp = tgApiCall_('getWebhookInfo', {});
  let text = resp ? resp.getContentText() : 'нет ответа';
  Logger.log('getTelegramWebhookInfo: ' + text);
  try { SpreadsheetApp.getUi().alert('Ответ Telegram: ' + text); }
  catch (e) { /* запущено не из таблицы — смотри Журнал выполнения */ }
}

/**
 * Триггер срабатывает раз в минуту (жёсткий лимит Apps Script), но внутри
 * одного срабатывания опрашиваем Telegram каждые 5 секунд в течение ~50
 * секунд — так задержка ответа на деле около 5-10 сек, а не до минуты.
 */
function checkTelegramUpdates() {
  let deadline = new Date().getTime() + 50 * 1000; // запас 10 сек до следующего триггера
  while (new Date().getTime() < deadline) {
    try {
      pollOnce_();
    } catch (e) {
      // Точечный сетевой сбой (например "Address unavailable") — тихо
      // логируем и ждём следующего цикла опроса, вместо того чтобы
      // ронять весь триггер целиком и заваливать почту письмами от
      // Google. Раз опрос всё равно идёт раз в минуту — один пропущенный
      // цикл не критичен, бот просто ответит на следующей итерации.
      Logger.log('checkTelegramUpdates: pollOnce_ упал, пропускаю цикл — ' + e.message);
    }
    Utilities.sleep(5000);
  }
}

function pollOnce_() {
  let token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) return;

  let props  = PropertiesService.getScriptProperties();
  let offset = Number(props.getProperty('TG_POLL_OFFSET') || 0);

  let resp;
  try {
    resp = UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + token + '/getUpdates?offset=' + offset + '&timeout=0',
      { method: 'get', muteHttpExceptions: true }
    );
  } catch (e) {
    // muteHttpExceptions защищает только от HTTP-кодов ошибок (404, 500 и
    // т.п.), но не от сетевых сбоев уровня "сервер не ответил вообще"
    // (DNS, обрыв соединения) — такие исключения долетают досюда напрямую.
    Logger.log('pollOnce_: сетевая ошибка при getUpdates — ' + e.message);
    return;
  }

  if (resp.getResponseCode() !== 200) {
    Logger.log('pollOnce_: getUpdates вернул ' + resp.getResponseCode() + ': ' + resp.getContentText());
    return;
  }

  let data = JSON.parse(resp.getContentText());
  Logger.log('pollOnce_: получено апдейтов — ' + (data.result ? data.result.length : 0));
  if (!data.ok || !data.result || !data.result.length) return;

  data.result.forEach(function(update) {
    try {
      if (update.callback_query) {
        let cq     = update.callback_query;
        let chatId = cq.message.chat.id;
        Logger.log('pollOnce_: callback_query от chatId=' + chatId + ', data=' + cq.data);
        tgApiCall_('answerCallbackQuery', { callback_query_id: cq.id });
        if (!isAuthorizedChat_(chatId)) {
          Logger.log('pollOnce_: chatId ' + chatId + ' НЕ авторизован (ожидался ' + PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID') + ')');
        } else if (isDebouncedAction_(chatId, 'CB:' + cq.data)) {
          Logger.log('pollOnce_: debounce заблокировал CB:' + cq.data);
        } else {
          routeCallback_(chatId, cq.data);
        }

      } else if (update.message) {
        let msg    = update.message;
        let chatId = msg.chat.id;
        let text = (msg.text || '').trim();
        Logger.log('pollOnce_: message от chatId=' + chatId + ', text="' + text + '"');
        if (!isAuthorizedChat_(chatId)) {
          Logger.log('pollOnce_: chatId ' + chatId + ' НЕ авторизован (ожидался ' + PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID') + ')');
        } else if (isDebouncedAction_(chatId, 'MSG:' + text)) {
          Logger.log('pollOnce_: debounce заблокировал MSG:' + text);
        } else {
          routeMessage_(chatId, text, msg.message_id);
        }
      }
    } catch (err) {
      Logger.log('pollOnce_ update error: ' + err.message + '\n' + err.stack);
      let chatId = update.callback_query ? update.callback_query.message.chat.id
                 : update.message ? update.message.chat.id : null;
      if (chatId && isAuthorizedChat_(chatId)) {
        try { sendTg_(chatId, '⚠️ Что-то пошло не так:\n' + err.message, mainMenuKeyboard_(), { topicKey: 'ERROR' }); } catch (e2) {}
      }
    }
    // offset = следующий update_id, чтобы Telegram больше не присылал уже обработанные
    props.setProperty('TG_POLL_OFFSET', String(update.update_id + 1));
  });
}

function tgApiCall_(method, payload) {
  let token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) return null;
  try {
    let resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true,
    });
    let code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log('tgApiCall_ (' + method + ') вернул код ' + code + ': ' + resp.getContentText());
    }
    return resp;
  } catch (e) {
    Logger.log('tgApiCall_ error: ' + e.message);
    return null;
  }
}

/**
 * Удаляет сообщение (в т.ч. входящее, от человека) — в приватном чате
 * Telegram Bot API это разрешает, в отличие от групп, где бот может
 * удалять чужие сообщения только с правами администратора. Используется,
 * чтобы «ответы-данные» (введённая сумма, список тикеров) не оседали
 * в чате мусором после того, как бот их уже обработал.
 */
function deleteTgMessage_(chatId, messageId) {
  if (!messageId) return;
  tgApiCall_('deleteMessage', { chat_id: chatId, message_id: messageId });
}

/**
 * «Семья меню» — Меню, Статус, ИИС-3, Концентрация. Все четыре открываются
 * прямо с одной и той же 2×2 клавиатуры и по смыслу — разные «лица» одного
 * и того же меню, а не отдельные шаги диалога. Поэтому переход МЕЖДУ ними
 * (а не только «Меню» → что-то ещё) тоже подчищает предыдущую карточку —
 * иначе Статус → ИИС-3 → Концентрация просто копится одно под другим.
 */
const MENU_FAMILY_TOPICS_ = ['MENU', 'STATUS', 'IIS', 'HEALTH', 'PL', 'ABOUT'];

/**
 * Удаляет все сообщения текущего «рабочего блока» (всё, что бот отправил
 * с момента последнего показа меню — например шаги ребалансировки).
 * Вызывается автоматически из sendTg_() при возврате в семью меню, так что
 * прошлая сессия не остаётся мусором ниже свежего меню.
 */
function flushBlockMessages_(chatId) {
  let props     = PropertiesService.getScriptProperties();
  let blockProp = 'TG_BLOCK_MSGS_' + chatId;
  let raw = props.getProperty(blockProp);
  if (!raw) return;
  let ids = JSON.parse(raw);
  ids.forEach(function(id) { deleteTgMessage_(chatId, id); });
  props.deleteProperty(blockProp);
}

/**
 * Отправка интерактивного сообщения (меню, статус, шаги ребалансировки).
 *
 * ВАЖНО: редактирует предыдущее сообщение бота на месте ТОЛЬКО если это
 * буквально повтор той же самой «темы» (opts.topicKey совпадает с темой
 * последнего отправленного сообщения) — например, два раза подряд нажали
 * «Статус». Любой другой переход (Меню → Ребаланс → сумма → результат)
 * — это НОВЫЙ шаг диалога, и он всегда шлётся новым сообщением.
 *
 * Это принципиально: если редактировать вообще любое сообщение подряд,
 * ответ бота остаётся там, где было старое сообщение, а не появляется
 * внизу рядом с вводом человека — после пары шагов кажется, что бот
 * не отвечает вообще, хотя на деле ответ просто «уехал» вверх по экрану.
 *
 * Все сообщения одного «рабочего блока» (всё, кроме самого меню) сами
 * запоминаются — и одним махом удаляются, как только человек возвращается
 * в меню (topicKey === 'MENU'). И наоборот: при уходе ИЗ меню В блок само
 * меню тоже удаляется — на экране в любой момент времени только одно
 * актуальное «место» (меню ИЛИ рабочий блок), никогда оба сразу.
 *
 * opts:
 *   topicKey — строка-тема сообщения ('MENU', 'STATUS', 'REBALANCE_RESULT'
 *              и т.п.). Без неё сообщение никогда не редактируется задним
 *              числом и не становится целью для редактирования следующего.
 *   forceNew — принудительно шлёт новое сообщение, даже если topicKey
 *              совпадает с предыдущим (нужно, когда предыдущую карточку
 *              важно оставить на экране, например список тикеров).
 */
function sendTg_(chatId, text, keyboard, opts) {
  opts = opts || {};
  let topicKey = opts.topicKey || null;
  let forceNew = !!opts.forceNew;

  let props        = PropertiesService.getScriptProperties();
  let msgIdKey     = 'TG_LAST_MSG_ID_' + chatId;
  let topicKeyProp = 'TG_LAST_TOPIC_' + chatId;
  let blockProp    = 'TG_BLOCK_MSGS_' + chatId;

  // Возврат в семью меню (Меню/Статус/ИИС-3/Концентрация) = предыдущий
  // рабочий блок (например шаги ребалансировки) завершён — подчищаем целиком.
  let isMenuFamily = MENU_FAMILY_TOPICS_.indexOf(topicKey) >= 0;
  if (isMenuFamily) {
    flushBlockMessages_(chatId);
  }

  let lastMsgId = props.getProperty(msgIdKey);
  let lastTopic = props.getProperty(topicKeyProp);
  let lastIsMenuFamily = MENU_FAMILY_TOPICS_.indexOf(lastTopic) >= 0;

  // Переход МЕЖДУ разными «лицами» семьи меню (Статус → ИИС-3 и т.п.),
  // а также уход ИЗ семьи меню В блок (например «Ребаланс») — предыдущую
  // карточку меню подчищаем: на экране в любой момент только одно
  // актуальное «место», никогда несколько одновременно.
  if (topicKey && lastMsgId && lastTopic && lastTopic !== topicKey && lastIsMenuFamily) {
    deleteTgMessage_(chatId, Number(lastMsgId));
    props.deleteProperty(msgIdKey);
    props.deleteProperty(topicKeyProp);
    lastMsgId = null;
    lastTopic = null;
  }

  let canEdit = !forceNew && lastMsgId && topicKey && topicKey === lastTopic;

  let replyMarkup = keyboard ? { inline_keyboard: keyboard } : null;

  if (canEdit) {
    let editPayload = { chat_id: chatId, message_id: Number(lastMsgId), text: text, parse_mode: 'HTML' };
    if (replyMarkup) editPayload.reply_markup = JSON.stringify(replyMarkup);

    let editResp = tgApiCall_('editMessageText', editPayload);
    if (editResp) {
      try {
        let body = JSON.parse(editResp.getContentText());
        Logger.log('sendTg_: editMessageText (topicKey=' + topicKey + ') → ' + editResp.getContentText());
        if (body.ok) {
          return; // успешно отредактировали на месте — новое сообщение не нужно
        }
        if (body.description && body.description.indexOf('message is not modified') >= 0) {
          // Содержимое не изменилось, но раз действие запросили явно (нажали
          // кнопку / прислали команду) — человек хочет видеть карточку СЕЙЧАС,
          // у себя перед глазами, а не там, где она случайно осталась выше
          // по чату. Убираем старую и шлём такую же новую — уже внизу.
          deleteTgMessage_(chatId, Number(lastMsgId));
          // не return — специально падаем ниже, на отправку нового сообщения
        }
      } catch (e) { Logger.log('sendTg_: не смог распарсить ответ editMessageText: ' + e.message); }
    } else {
      Logger.log('sendTg_: editMessageText вернул null (tgApiCall_ упал) — падаю на sendMessage');
    }
    // Редактирование не помогло (сообщение недоступно ИЛИ содержимое не изменилось) — обычная отправка
  }

  let payload = { chat_id: chatId, text: text, parse_mode: 'HTML' };
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);

  let sendResp = tgApiCall_('sendMessage', payload);
  if (!sendResp) { Logger.log('sendTg_: sendMessage вернул null (tgApiCall_ упал, см. лог выше)'); return; }
  Logger.log('sendTg_: sendMessage (topicKey=' + topicKey + ') → ' + sendResp.getContentText());
  try {
    let body = JSON.parse(sendResp.getContentText());
    if (body.ok && body.result && body.result.message_id) {
      let newId = body.result.message_id;
      props.setProperty(msgIdKey, String(newId));
      if (topicKey) props.setProperty(topicKeyProp, topicKey);
      else props.deleteProperty(topicKeyProp); // без темы — точно не должно "слипнуться" со следующим

      // Часть текущего рабочего блока (всё, кроме семьи меню) —
      // запоминаем id, чтобы удалить одним махом при возврате в меню.
      if (topicKey && MENU_FAMILY_TOPICS_.indexOf(topicKey) === -1) {
        let raw = props.getProperty(blockProp);
        let ids = raw ? JSON.parse(raw) : [];
        ids.push(newId);
        props.setProperty(blockProp, JSON.stringify(ids));
      }
    }
  } catch (e) { Logger.log('sendTg_: не смог распарсить ответ sendMessage: ' + e.message); }
}

// ── Клавиатуры ──────────────────────────────────────────────────────
//
// Меню — 2×2 сетка, сгруппированная по смыслу: «действия» (Ребаланс) отдельно
// от «показать цифру» (Статус/ИИС-3/Концентрация). «Статус» уже работает
// (sendStatusCard_) и отвечает и на кнопку, и на команду /status. «ИИС-3»
// и «Концентрация» пока заглушка — см. список дел (пункты 20, 22).

function mainMenuKeyboard_()   {
  return [
    [{ text: '💰 Ребаланс', callback_data: 'REBALANCE' }, { text: '📊 Статус', callback_data: 'STATUS' }],
    [{ text: '🏦 ИИС-3', callback_data: 'IIS' }, { text: '⚠️ Концентрация', callback_data: 'HEALTH' }],
  ];
}
function backKeyboard_()       { return [[{ text: '⬅️ Назад', callback_data: 'MENU' }]]; }
function resultKeyboard_(stockRows) {
  let rows = [];
  if (stockRows && stockRows.length) {
    let tickerButtons = stockRows
      .filter(function(r) { return r.ticker; })
      .map(function(r) {
        return {
          text: (r.isSkip ? '⚪ ' : '🟢 ') + r.ticker,
          callback_data: 'SKIP_TOGGLE:' + r.ticker,
        };
      });
    // По 3 кнопки в ряд — компактно даже при 9-10 акциях в портфеле
    for (let i = 0; i < tickerButtons.length; i += 3) {
      rows.push(tickerButtons.slice(i, i + 3));
    }
  }
  rows.push([{ text: '✏️ Изменить пропуски (текстом)', callback_data: 'EDIT_SKIP' }]);
  rows.push([{ text: '⬅️ Назад в меню', callback_data: 'MENU' }]);
  return rows;
}

// ── Состояние диалога (по chat_id) ───────────────────────────────────

function getChatState_(chatId) {
  let raw = PropertiesService.getScriptProperties().getProperty('TG_STATE_' + chatId);
  return raw ? JSON.parse(raw) : { step: 'MENU' };
}
function setChatState_(chatId, state) {
  PropertiesService.getScriptProperties().setProperty('TG_STATE_' + chatId, JSON.stringify(state));
}

// ── Маршрутизация ────────────────────────────────────────────────────

function routeCallback_(chatId, data) {
  if (data === 'MENU') {
    setChatState_(chatId, { step: 'MENU' });
    sendTg_(chatId, '📋 <b>Меню</b>', mainMenuKeyboard_(), { topicKey: 'MENU' });

  } else if (data === 'REBALANCE') {
    setChatState_(chatId, { step: 'AWAITING_AMOUNT' });
    sendTg_(chatId, '💰 Введите сумму пополнения, ₽:', backKeyboard_(), { topicKey: 'REBALANCE_AMOUNT' });

  } else if (data.indexOf('SKIP_TOGGLE:') === 0) {
    let ticker = data.substring('SKIP_TOGGLE:'.length);
    let state  = getChatState_(chatId);
    if (!state.amount) {
      sendTg_(chatId, '⚠️ Сначала запусти ребалансировку.', mainMenuKeyboard_(), { topicKey: 'ERROR' });
      return;
    }

    let plan = computeRebalancePlan_(state.amount, state.skipped || null);
    let row  = plan.stockRows.find(function(r) { return r.ticker === ticker; });
    if (!row) return; // тикер не нашёлся в текущем плане — игнорируем нажатие

    // Берём ПОЛНЫЙ текущий набор пропусков (категории + акции) из уже
    // посчитанного плана, а не только то, что лежит в state.skipped —
    // при первом расчёте пропуски могли прийти с чекбоксов листа
    // «Ребалансировка», а не из явного состояния чата.
    let currentSkipped = []
      .concat(plan.classRows.filter(function(r) { return r.isSkip; }).map(function(r) { return r.cat; }))
      .concat(plan.stockRows.filter(function(r) { return r.isSkip; }).map(function(r) { return r.name; }))
      .concat(plan.otherRows.filter(function(r) { return r.isSkip; }).map(function(r) { return r.cat; }));

    let idx = currentSkipped.indexOf(row.name);
    if (idx >= 0) currentSkipped.splice(idx, 1); else currentSkipped.push(row.name);

    setChatState_(chatId, { step: 'RESULT', amount: state.amount, skipped: currentSkipped });
    sendRebalanceResult_(chatId, state.amount, currentSkipped);

  } else if (data === 'EDIT_SKIP') {
    let state = getChatState_(chatId);
    setChatState_(chatId, { step: 'AWAITING_SKIP', amount: state.amount });
    sendTg_(chatId,
      '✏️ <b>Пропуск бумаг</b>\n\n' +
      'Список выше — впиши тикеры через запятую, которые нужно пропустить:\n' +
      '<code>SBER, LKOH</code>\n\n' +
      'Или просто «нет», если пропускать никого не нужно.',
      backKeyboard_(), { forceNew: true, topicKey: 'REBALANCE_SKIP_PROMPT' });

  } else if (data === 'STATUS') {
    sendStatusCard_(chatId);

  } else if (data === 'IIS') {
    sendIisCard_(chatId);

  } else if (data === 'HEALTH') {
    sendHealthCard_(chatId);
  }
}

function routeMessage_(chatId, text, messageId) {
  if (text === '/menu' || text === '/start') {
    setChatState_(chatId, { step: 'MENU' });
    sendTg_(chatId, '📋 <b>Меню</b>', mainMenuKeyboard_(), { topicKey: 'MENU' });
    return;
  }

  if (text === '/status') {
    sendStatusCard_(chatId);
    return;
  }

  if (text === '/iis') {
    sendIisCard_(chatId);
    return;
  }

  if (text === '/health') {
    sendHealthCard_(chatId);
    return;
  }

  if (text === '/pl' || text.indexOf('/pl ') === 0) {
    let arg = text.substring(3).trim();
    sendPlCard_(chatId, arg);
    return;
  }

  if (text === '/about') {
    sendTg_(chatId, aboutTrackerText_(), mainMenuKeyboard_(), { topicKey: 'ABOUT' });
    return;
  }

  let state = getChatState_(chatId);

  if (state.step === 'AWAITING_AMOUNT') {
    let amount = parseAmountInput_(text);
    if (amount === null) {
      sendTg_(chatId, '⚠️ Не понял сумму. Введите число, например 45000.', backKeyboard_(), { topicKey: 'REBALANCE_AMOUNT' });
      deleteTgMessage_(chatId, messageId);
      return;
    }
    setChatState_(chatId, { step: 'RESULT', amount: amount, skipped: null });
    sendRebalanceResult_(chatId, amount, null);
    deleteTgMessage_(chatId, messageId);
    return;
  }

  if (state.step === 'AWAITING_SKIP' || state.step === 'RESULT') {
    let plan      = computeRebalancePlan_(state.amount, null); // берём позиции для сопоставления тикеров
    let skipNames = parseSkipTickers_(text, plan.stockRows);
    setChatState_(chatId, { step: 'RESULT', amount: state.amount, skipped: skipNames });
    sendRebalanceResult_(chatId, state.amount, skipNames);
    deleteTgMessage_(chatId, messageId);
    return;
  }

  sendTg_(chatId, 'Напишите /menu, чтобы открыть меню.', null, { topicKey: 'FALLBACK' });
}

/**
 * /status — моментальный текстовый снимок портфеля, без захода в таблицу.
 * Полностью на уже посчитанных данных (Script Properties) — ни одного
 * нового запроса к T-Invest API, поэтому отвечает мгновенно. Если какая-то
 * метрика ни разу не считалась — честно пишет «н/д» вместо тишины.
 */
function sendStatusCard_(chatId) {
  let config;
  try { config = readConfig_(); }
  catch (e) { sendTg_(chatId, '⚠️ Не получилось прочитать Config: ' + e.message, mainMenuKeyboard_(), { topicKey: 'ERROR' }); return; }

  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);

  let props    = PropertiesService.getScriptProperties();
  let xirrRaw  = props.getProperty('ANALYTICS_XIRR_VALUE');
  let benchRaw = props.getProperty('ANALYTICS_BENCH_VALUE');
  let devRaw   = props.getProperty(DEVIATION_STATUS_PROP);

  let lines = ['📊 <b>Портфель</b>', '', rub_(Math.round(totalRub)), ''];

  if (xirrRaw) {
    let xirr = JSON.parse(xirrRaw).value;
    lines.push('XIRR       ' + (xirr !== null ? (xirr >= 0 ? '+' : '') + xirr.toFixed(1) + '%' : 'н/д'));
  } else {
    lines.push('XIRR       н/д — запусти расчёт на дашборде');
  }

  if (benchRaw) {
    let b    = JSON.parse(benchRaw);
    let diff = b.actual - b.hypothetical;
    lines.push('vs IMOEX   ' + (diff >= 0 ? '+' : '−') + rub_(Math.abs(diff)));
  } else {
    lines.push('vs IMOEX   н/д — запусти расчёт на дашборде');
  }

  lines.push('');
  if (devRaw) {
    let dev  = JSON.parse(devRaw);
    let cats = Object.keys(dev);
    if (cats.length) {
      lines.push('<b>5/25:</b>');
      cats.forEach(function(cat) { lines.push('• ' + cat + ': ' + dev[cat]); });
    }
  } else {
    lines.push('5/25       н/д — ещё ни разу не проверялось');
  }

  sendTg_(chatId, lines.join('\n'), mainMenuKeyboard_(), { topicKey: 'STATUS' });
}

/**
 * Карточка ИИС-3 — переиспользует IIS_DEDUCTION_DATA, посчитанный на
 * дашборде (Tinkoff → 🎯 Аналитика → «ИИС-3 — вычет за год»). Ни одного
 * нового запроса к API — если данных ещё нет, честно просит сначала
 * посчитать на дашборде.
 */
function sendIisCard_(chatId) {
  let raw = PropertiesService.getScriptProperties().getProperty('IIS_DEDUCTION_DATA');
  if (!raw) {
    sendTg_(chatId,
      '🏦 <b>ИИС-3</b>\n\nЕщё не считалось. Запусти на дашборде:\nTinkoff → 🎯 Аналитика → «ИИС-3 — вычет за год»',
      mainMenuKeyboard_(), { topicKey: 'IIS' });
    return;
  }

  let d = JSON.parse(raw);
  let lines = [
    '🏦 <b>ИИС-3 — вычет за ' + d.year + ' год</b>',
    '',
    'Счёт: ' + d.accountName,
    'Использовано: ' + (d.pct * 100).toFixed(1) + '%',
    rub_(d.contributed) + ' из ' + rub_(IIS_DEDUCTION_LIMIT),
    '',
  ];
  lines.push(d.remaining > 0
    ? 'Осталось довнести до конца года: ' + rub_(d.remaining)
    : '✅ Лимит вычета за этот год полностью выбран');
  if (d.remaining > 0 && d.daysLeft !== undefined) {
    lines.push('Осталось дней до 31 декабря: ' + d.daysLeft);
  }

  if (d.holdInfo) {
    let h = d.holdInfo;
    let tz = Session.getScriptTimeZone();
    let unlockStr = Utilities.formatDate(new Date(h.unlockDate), tz, 'dd.MM.yyyy');
    lines.push('');
    if (h.daysUntilUnlock > 0) {
      let years = Math.floor(h.daysUntilUnlock / 365);
      let restDays = h.daysUntilUnlock % 365;
      lines.push('⏳ Мин. срок владения: ' + h.minYears + ' лет');
      lines.push('До ' + unlockStr + ' — ещё ~' + years + ' г. ' + restDays + ' дн.');
      lines.push('(закрыть раньше — потерять все льготы)');
    } else {
      lines.push('✅ Минимальный срок владения (' + h.minYears + ' лет) уже прошёл — ' + unlockStr);
    }
  }

  sendTg_(chatId, lines.join('\n'), mainMenuKeyboard_(), { topicKey: 'IIS' });
}

/**
 * Карточка концентрации — переиспользует HEALTH_CONCENTRATION_DATA,
 * посчитанный на дашборде (Tinkoff → 🎯 Аналитика → «Health check:
 * концентрация»). Ни одного нового запроса к API.
 */
function sendHealthCard_(chatId) {
  let raw = PropertiesService.getScriptProperties().getProperty('HEALTH_CONCENTRATION_DATA');
  if (!raw) {
    sendTg_(chatId,
      '⚠️ <b>Концентрация</b>\n\nЕщё не считалось. Запусти на дашборде:\nTinkoff → 🎯 Аналитика → «Health check: концентрация»',
      mainMenuKeyboard_(), { topicKey: 'HEALTH' });
    return;
  }

  let d = JSON.parse(raw);
  let lines = ['⚠️ <b>Концентрация портфеля</b>', ''];

  if (d.overall && d.overall.top1Name) {
    lines.push('Крупнейшая позиция (весь портфель):');
    lines.push(d.overall.top1Name + ' — ' + (d.overall.top1Pct * 100).toFixed(1) + '%');
    lines.push('Топ-3: ' + (d.overall.top3Pct * 100).toFixed(1) + '%');
    if (d.overall.effectiveN !== undefined) {
      lines.push('Эффективно позиций: ' + d.overall.effectiveN.toFixed(1) + ' из ' + d.overall.realN);
    }
    lines.push('');
  }

  if (d.shares && d.shares.top1Name) {
    lines.push('Крупнейшая акция (доля от акций):');
    lines.push(d.shares.top1Name + ' — ' + (d.shares.top1Pct * 100).toFixed(1) + '%');
    lines.push('Топ-3 акции: ' + (d.shares.top3Pct * 100).toFixed(1) + '%');
    if (d.shares.effectiveN !== undefined) {
      lines.push('Эффективно акций: ' + d.shares.effectiveN.toFixed(1) + ' из ' + d.shares.realN);
    }
  }

  sendTg_(chatId, lines.join('\n'), mainMenuKeyboard_(), { topicKey: 'HEALTH' });
}

/**
 * /pl <тикер> — P/L по конкретной акции. Переиспользует уже посчитанные
 * данные из avgprice.js (AVGPRICE_DATA), сопоставляя тикер с позицией
 * через readPositions_(). Ни одного нового запроса к API.
 */
function sendPlCard_(chatId, tickerArg) {
  if (!tickerArg) {
    sendTg_(chatId, '✏️ Укажи тикер: например /pl SBER', mainMenuKeyboard_(), { topicKey: 'PL' });
    return;
  }

  let avgRaw = PropertiesService.getScriptProperties().getProperty(AVGPRICE_PROP);
  if (!avgRaw) {
    sendTg_(chatId,
      '📊 <b>P/L по акции</b>\n\nЕщё не считалось. Запусти на дашборде:\nTinkoff → 🎯 Аналитика → «Средняя цена и P/L»',
      mainMenuKeyboard_(), { topicKey: 'PL' });
    return;
  }

  let config;
  try { config = readConfig_(); }
  catch (e) { sendTg_(chatId, '⚠️ Не получилось прочитать Config: ' + e.message, mainMenuKeyboard_(), { topicKey: 'PL' }); return; }

  let positions = readPositions_(config);
  let ticker = tickerArg.trim().toUpperCase();
  let pos = positions.find(function(p) {
    return p.category === 'Акции' && (p.ticker || '').toUpperCase() === ticker;
  });

  if (!pos) {
    sendTg_(chatId, '⚠️ Не нашёл акцию с тикером «' + ticker + '» в портфеле.', mainMenuKeyboard_(), { topicKey: 'PL' });
    return;
  }

  let avgData = JSON.parse(avgRaw);
  let res = avgData.find(function(r) { return r.name === pos.name; });

  if (!res || !(res.avgPrice > 0)) {
    sendTg_(chatId,
      '📊 <b>' + pos.name + ' (' + ticker + ')</b>\n\nНет истории сделок за выбранный период — средняя цена не посчитана.',
      mainMenuKeyboard_(), { topicKey: 'PL' });
    return;
  }

  let arrow = res.pl >= 0 ? '📈' : '📉';
  let lines = [
    '📊 <b>' + res.name + ' (' + ticker + ')</b>',
    '',
    'Кол-во: ' + res.qty,
    'Ср. цена: ' + rub_(res.avgPrice),
    'Тек. цена: ' + rub_(res.curPrice),
    '',
    arrow + ' P/L: ' + (res.pl >= 0 ? '+' : '') + rub_(Math.round(res.pl)) +
      '  (' + (res.plPct * 100 >= 0 ? '+' : '') + (res.plPct * 100).toFixed(1) + '%)',
  ];

  if (res.mismatch) {
    lines.push('');
    lines.push('⚠️ FIFO-количество (' + res.fifoQty + ') не совпадает с фактическим (' + res.qty + ')');
  }

  sendTg_(chatId, lines.join('\n'), mainMenuKeyboard_(), { topicKey: 'PL' });
}

function parseAmountInput_(text) {
  let cleaned = text.replace(/\s/g, '').replace(',', '.');
  let n = Number(cleaned);
  return (!isNaN(n) && n > 0) ? n : null;
}

function parseSkipTickers_(text, stockRows) {
  if (/^нет$/i.test(text.trim())) return [];
  let tickers = text.split(',').map(function(s) { return s.trim().toUpperCase(); }).filter(Boolean);
  let names = [];
  tickers.forEach(function(t) {
    let found = stockRows.find(function(r) { return (r.ticker || '').toUpperCase() === t; });
    if (found) names.push(found.name); // .name здесь — это точная строка из Config, не из сырых позиций
  });
  return names;
}

function sendRebalanceResult_(chatId, amount, skippedOverride) {
  try {
    let plan = computeRebalancePlan_(amount, skippedOverride);
    let text = formatRebalanceText_(plan);
    sendTg_(chatId, text, resultKeyboard_(plan.stockRows), { topicKey: 'REBALANCE_RESULT' });
  } catch (e) {
    Logger.log('sendRebalanceResult_ error: ' + e.message + '\n' + e.stack);
    sendTg_(chatId, '⚠️ Не получилось посчитать: ' + e.message, mainMenuKeyboard_(), { topicKey: 'ERROR' });
  }
}

/**
 * Расчёт ребалансировки для бота — независимая копия математики из
 * calculateRebalance() (dashboard.js), но НЕ трогает лист «Ребалансировка»
 * и не изменяет его. Если skippedOverride === null — берёт список
 * пропущенных бумаг/категорий с чекбоксов на самом листе (если он есть);
 * если передан массив — использует его вместо чекбоксов листа.
 */
function computeRebalancePlan_(amount, skippedOverride) {
  let config    = readConfig_();
  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);
  let newTotal  = totalRub + amount;

  let skipped = skippedOverride;
  if (skipped === null || skipped === undefined) {
    let ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(DST.REBALANCE);
    skipped = sh ? readSkipped_(sh) : [];
  }

  let cats      = Object.keys(config.classTargets);
  let classNeed = {};
  cats.forEach(function(cat) {
    let actual = positions.filter(function(p) { return p.category === cat; })
                          .reduce(function(s, p) { return s + p.valueRub; }, 0);
    let need   = Math.max(0, newTotal * (config.classTargets[cat] || 0) - actual);
    classNeed[cat] = { need: need, actual: actual };
  });

  let activeCats   = cats.filter(function(c) { return skipped.indexOf(c) === -1; });
  let totalNeedCls = activeCats.reduce(function(s, c) { return s + classNeed[c].need; }, 0);
  let classAlloc   = {};
  cats.forEach(function(cat) {
    classAlloc[cat] = (skipped.indexOf(cat) !== -1 || totalNeedCls === 0)
      ? 0 : (classNeed[cat].need / totalNeedCls) * amount;
  });

  let classRows = cats.map(function(cat) {
    let actual = classNeed[cat].actual;
    let actPct = totalRub > 0 ? actual / totalRub : 0;
    let tgtPct = config.classTargets[cat] || 0;
    return { cat: cat, actPct: actPct, tgtPct: tgtPct, alloc: classAlloc[cat], isSkip: skipped.indexOf(cat) !== -1 };
  });

  let stockBudget = classAlloc['Акции'] || 0;
  let sharePos    = positions.filter(function(p) { return p.category === 'Акции'; });
  let stockNeed   = {};
  Object.keys(config.stockTargets).forEach(function(name) {
    let tgt    = config.stockTargets[name];
    let pos    = matchPos_(sharePos, name);
    let actual = pos ? pos.valueRub : 0;
    let need   = Math.max(0, newTotal * tgt - actual);
    stockNeed[name] = { need: need, pos: pos, actual: actual, tgt: tgt, price: pos ? pos.price : 0, lot: pos ? pos.lot : 1 };
  });
  let lotResult  = allocateWithLots_(stockBudget, stockNeed, skipped);
  let stockOrder = Object.keys(stockNeed).sort(function(a, b) { return stockNeed[b].need - stockNeed[a].need; });
  let stockRows  = stockOrder.map(function(name) {
    let info   = stockNeed[name];
    let isSkip = skipped.indexOf(name) !== -1;
    let res    = lotResult.results[name] || { lots: 0, actualAlloc: 0 };
    return {
      name: name, ticker: info.pos ? info.pos.ticker : '',
      alloc: isSkip ? 0 : res.actualAlloc, lots: isSkip ? 0 : res.lots,
      isSkip: isSkip, unknown: res.unknown,
    };
  });

  let otherCats  = ['Золото', 'Замещайки', 'Денежный рынок'];
  let otherRows  = otherCats.map(function(cat) {
    let alloc  = classAlloc[cat] || 0;
    let isSkip = skipped.indexOf(cat) !== -1;
    let catPos = positions.filter(function(p) { return p.category === cat; })
                          .sort(function(a, b) { return (b.price || 0) - (a.price || 0); });
    return { cat: cat, alloc: isSkip ? 0 : alloc, isSkip: isSkip, instrName: catPos.length ? catPos[0].name : '—' };
  });

  return {
    amount: amount, totalRub: totalRub, classRows: classRows,
    stockBudget: stockBudget, stockRows: stockRows,
    otherRows: otherRows, sharePositions: sharePos,
  };
}

function formatRebalanceText_(plan) {
  let lines = ['💰 <b>Ребалансировка на ' + rub_(plan.amount) + '</b>', ''];

  lines.push('<b>По классам:</b>');
  plan.classRows.forEach(function(r) {
    if (r.isSkip || r.alloc <= 0) return;
    lines.push('• ' + r.cat + ': ' + rub_(Math.round(r.alloc)));
  });

  lines.push('', '<b>Акции</b> (бюджет ' + rub_(Math.round(plan.stockBudget)) + '):');
  let anyStock = false;
  plan.stockRows.forEach(function(r) {
    if (r.isSkip) {
      lines.push('• ' + r.name + ' (' + r.ticker + '): пропущено');
      return;
    }
    if (r.lots <= 0) return;
    anyStock = true;
    lines.push('• ' + r.name + ' (' + r.ticker + '): ' + r.lots + ' лот(ов)');
  });
  if (!anyStock) lines.push('— нечего докупать по акциям в этот раз');

  lines.push('', '<b>Другие категории:</b>');
  let anyOther = false;
  plan.otherRows.forEach(function(r) {
    if (r.isSkip) {
      lines.push('• ' + r.cat + ' (' + r.instrName + '): пропущено');
      return;
    }
    if (r.alloc <= 0) return;
    anyOther = true;
    lines.push('• ' + r.cat + ' (' + r.instrName + '): ' + rub_(Math.round(r.alloc)));
  });
  if (!anyOther) lines.push('— ничего не требуется');

  return lines.join('\n');
}
