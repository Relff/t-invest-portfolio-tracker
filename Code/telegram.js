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
    '📊 <b>Изменение статуса по правилу 5/25</b>\n' + changes.join('\n')
  );
}

/**
 * Уведомление о завершении ежедневной синхронизации.
 * Вызывается из syncAndRefresh() в dashboard.js. Если Telegram не настроен —
 * sendTelegramMessage_ тихо ничего не делает, ошибок не будет.
 */
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
    now + '\n' +
    'Портфель: ' + rub_(totalRub)
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
    '📅 <b>Выплаты на ближайшую неделю</b>\n' +
    'Купоны: ' + rub_(couponSum) + '\n' +
    'Дивиденды: ' + rub_(dividendSum) + '\n' +
    'Итого: ' + rub_(totalSum)
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
  SpreadsheetApp.getUi().alert('Ответ Telegram: ' + resp.getContentText());
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
      if (isAuthorizedChat_(chatId)) routeCallback_(chatId, cq.data);

    } else if (update.message) {
      let msg    = update.message;
      let chatId = msg.chat.id;
      chatIdForError = chatId;
      if (isAuthorizedChat_(chatId)) routeMessage_(chatId, (msg.text || '').trim());
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    if (chatIdForError && isAuthorizedChat_(chatIdForError)) {
      try {
        sendTg_(chatIdForError, '⚠️ Внутренняя ошибка бота:\n' + err.message, mainMenuKeyboard_());
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
  SpreadsheetApp.getUi().alert('Ответ Telegram: ' + (resp ? resp.getContentText() : 'нет ответа'));
}

/**
 * Триггер срабатывает раз в минуту (жёсткий лимит Apps Script), но внутри
 * одного срабатывания опрашиваем Telegram каждые 5 секунд в течение ~50
 * секунд — так задержка ответа на деле около 5-10 сек, а не до минуты.
 */
function checkTelegramUpdates() {
  let deadline = new Date().getTime() + 50 * 1000; // запас 10 сек до следующего триггера
  while (new Date().getTime() < deadline) {
    pollOnce_();
    Utilities.sleep(5000);
  }
}

function pollOnce_() {
  let token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) return;

  let props  = PropertiesService.getScriptProperties();
  let offset = Number(props.getProperty('TG_POLL_OFFSET') || 0);

  let resp = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + token + '/getUpdates?offset=' + offset + '&timeout=0',
    { method: 'get', muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) {
    Logger.log('pollOnce_: getUpdates вернул ' + resp.getResponseCode() + ': ' + resp.getContentText());
    return;
  }

  let data = JSON.parse(resp.getContentText());
  if (!data.ok || !data.result || !data.result.length) return;

  data.result.forEach(function(update) {
    try {
      if (update.callback_query) {
        let cq     = update.callback_query;
        let chatId = cq.message.chat.id;
        tgApiCall_('answerCallbackQuery', { callback_query_id: cq.id });
        if (isAuthorizedChat_(chatId)) routeCallback_(chatId, cq.data);

      } else if (update.message) {
        let msg    = update.message;
        let chatId = msg.chat.id;
        if (isAuthorizedChat_(chatId)) routeMessage_(chatId, (msg.text || '').trim());
      }
    } catch (err) {
      Logger.log('pollOnce_ update error: ' + err.message + '\n' + err.stack);
      let chatId = update.callback_query ? update.callback_query.message.chat.id
                 : update.message ? update.message.chat.id : null;
      if (chatId && isAuthorizedChat_(chatId)) {
        try { sendTg_(chatId, '⚠️ Внутренняя ошибка бота:\n' + err.message, mainMenuKeyboard_()); } catch (e2) {}
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

function sendTg_(chatId, text, keyboard) {
  let payload = { chat_id: chatId, text: text, parse_mode: 'HTML' };
  if (keyboard) payload.reply_markup = JSON.stringify({ inline_keyboard: keyboard });
  tgApiCall_('sendMessage', payload);
}

// ── Клавиатуры ──────────────────────────────────────────────────────

function mainMenuKeyboard_()   { return [[{ text: '📊 Ребаланс', callback_data: 'REBALANCE' }]]; }
function backKeyboard_()       { return [[{ text: '⬅️ Назад', callback_data: 'MENU' }]]; }
function resultKeyboard_()     {
  return [
    [{ text: '✏️ Изменить пропуски', callback_data: 'EDIT_SKIP' }],
    [{ text: '⬅️ Назад в меню', callback_data: 'MENU' }],
  ];
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
    sendTg_(chatId, '📋 <b>Меню</b>', mainMenuKeyboard_());

  } else if (data === 'REBALANCE') {
    setChatState_(chatId, { step: 'AWAITING_AMOUNT' });
    sendTg_(chatId, '💰 Введите сумму пополнения, ₽:', backKeyboard_());

  } else if (data === 'EDIT_SKIP') {
    let state = getChatState_(chatId);
    setChatState_(chatId, { step: 'AWAITING_SKIP', amount: state.amount });
    sendTg_(chatId,
      '✏️ Напишите тикеры через запятую, которые нужно пропустить (например SBER, LKOH), или «нет», чтобы никого не пропускать:',
      backKeyboard_());
  }
}

function routeMessage_(chatId, text) {
  if (text === '/menu' || text === '/start') {
    setChatState_(chatId, { step: 'MENU' });
    sendTg_(chatId, '📋 <b>Меню</b>', mainMenuKeyboard_());
    return;
  }

  let state = getChatState_(chatId);

  if (state.step === 'AWAITING_AMOUNT') {
    let amount = parseAmountInput_(text);
    if (amount === null) {
      sendTg_(chatId, '⚠️ Не понял сумму. Введите число, например 45000.', backKeyboard_());
      return;
    }
    setChatState_(chatId, { step: 'RESULT', amount: amount, skipped: null });
    sendRebalanceResult_(chatId, amount, null);
    return;
  }

  if (state.step === 'AWAITING_SKIP') {
    let plan      = computeRebalancePlan_(state.amount, null); // берём позиции для сопоставления тикеров
    let skipNames = parseSkipTickers_(text, plan.stockRows);
    setChatState_(chatId, { step: 'RESULT', amount: state.amount, skipped: skipNames });
    sendRebalanceResult_(chatId, state.amount, skipNames);
    return;
  }

  sendTg_(chatId, 'Напишите /menu, чтобы открыть меню.');
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
    sendTg_(chatId, text, resultKeyboard_());
  } catch (e) {
    Logger.log('sendRebalanceResult_ error: ' + e.message + '\n' + e.stack);
    sendTg_(chatId, '⚠️ Ошибка при расчёте: ' + e.message, mainMenuKeyboard_());
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
