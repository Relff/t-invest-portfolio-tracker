/**
 * calendar.js — Календарь выплат
 *
 * Показывает ожидаемые купоны и дивиденды по месяцам.
 *
 * Листы:
 *   «Календарь выплат» — месячная сводка + список ближайших выплат
 *
 * Установка:
 *   1. Добавить файл calendar в Apps Script
 *   2. В onOpen() добавить:
 *      .addItem('📅  Обновить Календарь выплат', 'updateCalendarSheet')
 *
 * Зависимости (из других файлов проекта):
 *   tiFetch_(), moneyToNumber_()         — tinvest.gs
 *   C, DST, rub_(), mergedCell_(), hdrRow_() — dashboard.gs
 */

// Название листа берётся из DST.CALENDAR (dashboard.gs)

// ════════════════════════════════════════════════════════════════════
// ГЛАВНАЯ ФУНКЦИЯ
// ════════════════════════════════════════════════════════════════════

function updateCalendarSheet() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.CALENDAR);
  if (!sh) sh = ss.insertSheet(DST.CALENDAR);
  sh.clearContents();
  sh.clearFormats();

  let tz  = Session.getScriptTimeZone();
  let now = new Date();
  let nowStr = Utilities.formatDate(now, tz, 'dd.MM.yyyy HH:mm');

  // Читаем позиции
  let bondPositions  = readSheetPositions_('Дан_Облигации');
  let sharePositions = readSheetPositions_('Дан_Акции');
  let figiMap        = buildFigiMap_();
  let divMap         = readDividendsFromConfig_();

  // Собираем расписание выплат
  let couponPayments   = fetchCouponCalendar_(bondPositions, figiMap);
  let dividendPayments = fetchDividendCalendar_(sharePositions, divMap, figiMap);
  let allPayments      = couponPayments.concat(dividendPayments);

  // Строим месячные итоги
  let monthlyGrid = buildMonthlyGrid_(allPayments, now);

  // Рендерим лист
  let r    = 1;
  let COLS = 7;

  // ── Шапка ─────────────────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS,
    '📅  КАЛЕНДАРЬ ВЫПЛАТ — купоны и дивиденды',
    { bg: C.DARK, fg: '#ffffff', bold: true, size: 14, align: 'center' });
  r++;

  mergedCell_(sh, r, 1, 1, COLS,
    'Обновлено: ' + nowStr + '   ·   Данные за следующие 12 месяцев',
    { bg: '#263238', fg: '#b0bec5', align: 'center' });
  r += 2;

  // ── Месячная сводка ───────────────────────────────────────────────
  mergedCell_(sh, r, 1, 1, COLS, '▌ МЕСЯЧНАЯ СВОДКА',
    { bg: C.MID, fg: '#ffffff', bold: true });
  r++;

  hdrRow_(sh, r,
    ['Месяц', 'Купоны, ₽', 'Дивиденды, ₽', 'Итого, ₽', 'Прогресс', '', ''],
    COLS);
  r++;

  // Находим максимум для прогресс-баров
  let maxMonthly = 0;
  Object.keys(monthlyGrid).forEach(function(m) {
    let tot = (monthlyGrid[m].coupons || 0) + (monthlyGrid[m].dividends || 0);
    if (tot > maxMonthly) maxMonthly = tot;
  });

  let MONTH_NAMES = [
    'Январь','Февраль','Март','Апрель','Май','Июнь',
    'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'
  ];

  let yearCoupons   = 0;
  let yearDividends = 0;

  for (let mi = 0; mi < 12; mi++) {
    let monthIdx  = ((now.getMonth() + mi) % 12) + 1; // 1-12
    let yearOffset = (now.getMonth() + mi) >= 12 ? 1 : 0;
    let dispYear  = now.getFullYear() + yearOffset;
    let grid      = monthlyGrid[monthIdx + '_' + dispYear] || { coupons: 0, dividends: 0 };
    let couponAmt = grid.coupons   || 0;
    let divAmt    = grid.dividends || 0;
    let total     = couponAmt + divAmt;

    yearCoupons   += couponAmt;
    yearDividends += divAmt;

    // Прогресс-бар
    let barLen  = maxMonthly > 0 ? Math.round((total / maxMonthly) * 20) : 0;
    let barStr  = '█'.repeat(barLen) + '░'.repeat(20 - barLen);
    let barColor = total > 5000 ? '#1b5e20' : total > 2000 ? '#f57f17' : '#9e9e9e';

    // Подсветка текущего месяца
    let isCurrentMonth = (monthIdx === now.getMonth() + 1 && dispYear === now.getFullYear());
    let bg = isCurrentMonth ? '#e8f5e9' : (mi % 2 === 0 ? C.EVEN : C.ODD);

    sh.getRange(r, 1).setValue(MONTH_NAMES[monthIdx - 1] + ' ' + dispYear)
      .setBackground(bg).setFontWeight(isCurrentMonth ? 'bold' : 'normal');
    sh.getRange(r, 2).setValue(Math.round(couponAmt)).setNumberFormat('#,##0 [$₽-ru-RU]').setBackground(bg);
    sh.getRange(r, 3).setValue(Math.round(divAmt)).setNumberFormat('#,##0 [$₽-ru-RU]').setBackground(bg);
    sh.getRange(r, 4).setValue(Math.round(total)).setNumberFormat('#,##0 [$₽-ru-RU]').setBackground(bg)
      .setFontWeight('bold');
    sh.getRange(r, 5, 1, 3).merge()
      .setValue(barStr + '  ' + rub_(total))
      .setBackground(bg).setFontColor(barColor)
      .setFontFamily('Courier New').setFontSize(9);
    r++;
  }

  // Итоговая строка
  let yearTotal = yearCoupons + yearDividends;
  sh.getRange(r, 1, 1, COLS).merge()
    .setValue('ИТОГО ЗА 12 МЕСЯЦЕВ:   купоны ' + rub_(yearCoupons) +
              '  +  дивиденды ' + rub_(yearDividends) +
              '  =  ' + rub_(yearTotal))
    .setBackground(C.DARK).setFontColor('#ffd54f')
    .setFontWeight('bold').setHorizontalAlignment('center');
  r += 2;

  // ── Все будущие выплаты (12 месяцев, без дублирования) ──────────
  // 🟡 Строки с выплатами в ближайшие 7 дней подсвечены жёлтым
  let allFuture = allPayments
    .filter(function(p) { return p.date && p.date >= now; })
    .sort(function(a, b) { return a.date - b.date; });

  mergedCell_(sh, r, 1, 1, COLS,
    '▌ ВСЕ ВЫПЛАТЫ — следующие 12 месяцев (' + allFuture.length + ' выплат)' +
    '   |   🟡 жёлтым = ближайшие 7 дней',
    { bg: '#37474f', fg: '#ffffff', bold: true });
  r++;

  let paymentsHeaderRow = r; // запоминаем строку заголовка для фильтра
  hdrRow_(sh, r,
    ['Дата', 'Инструмент', 'Тип', '₽ / шт', 'Кол-во', 'Итого, ₽', 'Месяц'],
    COLS);
  r++;

  allFuture.forEach(function(p, idx) {
    let monthStr  = MONTH_NAMES[p.date.getMonth()] + ' ' + p.date.getFullYear();
    let bg        = idx % 2 === 0 ? C.EVEN : C.ODD;
    let daysLeft  = Math.ceil((p.date - now) / (24 * 3600 * 1000));
    let urgent    = daysLeft <= 7;

    sh.getRange(r, 1).setValue(p.date).setNumberFormat('dd.mm.yyyy').setBackground(urgent ? '#fff9c4' : bg);
    sh.getRange(r, 2, 1, COLS - 1).setValues([[
      p.name, p.type,
      Math.round(p.perUnit * 100) / 100,
      p.qty,
      Math.round(p.total),
      monthStr
    ]]).setBackground(urgent ? '#fff9c4' : bg);

    sh.getRange(r, 4).setNumberFormat('#,##0.00 [$₽-ru-RU]');
    sh.getRange(r, 6).setNumberFormat('#,##0 [$₽-ru-RU]');
    if (urgent) sh.getRange(r, 1).setFontColor('#e65100').setFontWeight('bold');
    r++;
  });

  // ── Ширина колонок ────────────────────────────────────────────────
  [100, 230, 90, 130, 80, 140, 130].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(6);

  // Добавляем фильтр на таблицу выплат (чтобы можно было фильтровать по месяцу/инструменту)
  try {
    let existingFilter = sh.getFilter();
    if (existingFilter) existingFilter.remove();
    // Фильтр на всю таблицу (от строки заголовка до конца данных)
    sh.getRange(paymentsHeaderRow, 1, allFuture.length + 1, COLS).createFilter();
  } catch(e) {
    console.warn('Календарь: не удалось создать фильтр: ' + e.message);
  }

  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════════════
// КУПОННЫЙ КАЛЕНДАРЬ — точные даты из T-Invest API
// ════════════════════════════════════════════════════════════════════

/**
 * Для каждой облигации в портфеле получаем расписание купонов
 * через GetBondCoupons. Возвращаем массив платежей с точными датами.
 */
function fetchCouponCalendar_(bondPositions, figiMap) {
  let now  = new Date();
  let in1y = new Date(now.getTime() + 366 * 24 * 3600 * 1000);
  let payments = [];

  bondPositions.forEach(function(p) {
    let figi = p.figi || figiMap[p.name] || figiMap[p.ticker] || '';
    if (!figi) return;

    try {
      let resp = tiFetch_(
        '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetBondCoupons',
        { figi: figi, from: now.toISOString(), to: in1y.toISOString() }
      );

      let events = resp.events || [];
      events.forEach(function(c) {
        // Дата купона
        let dateRaw = c.couponDate;
        if (!dateRaw) return;

        let payDate;
        if (typeof dateRaw === 'string') {
          payDate = new Date(dateRaw);
        } else if (dateRaw.seconds) {
          payDate = new Date(Number(dateRaw.seconds) * 1000);
        } else {
          return;
        }

        if (isNaN(payDate.getTime())) return;
        if (payDate < now) return; // пропускаем прошедшие

        let perUnit = moneyToNumber_(c.payOneBond || c.couponAmount || null);
        if (perUnit <= 0) return;

        payments.push({
          name:    p.name,
          ticker:  p.ticker,
          type:    'Купон',
          date:    payDate,
          perUnit: perUnit,
          qty:     p.qty,
          total:   perUnit * p.qty,
        });
      });

      Utilities.sleep(50); // throttle API calls
    } catch(e) {
      // Пропускаем облигацию если ошибка API
      console.warn('Календарь: ошибка купонов для ' + p.name + ': ' + e.message);
    }
  });

  return payments;
}


// ════════════════════════════════════════════════════════════════════
// ДИВИДЕНДНЫЙ КАЛЕНДАРЬ — исторические даты + Config суммы
// ════════════════════════════════════════════════════════════════════

/**
 * Для акций пытаемся определить типичный месяц выплаты дивидендов
 * через GetDividends (исторические данные).
 * Сумму берём из Config (Блок 4).
 */
function fetchDividendCalendar_(sharePositions, divMap, figiMap) {
  let payments = [];
  let now      = new Date();

  sharePositions.forEach(function(p) {
    let divData = findDividend_(p.name, p.ticker, divMap);
    if (!divData || divData.amount <= 0) return; // нет дивидендов

    let figi = p.figi || figiMap[p.name] || figiMap[p.ticker] || '';
    let payMonth = null;
    let payYear  = now.getFullYear();

    // Пытаемся получить типичный месяц из истории
    if (figi) {
      try {
        let from2y = new Date(now.getTime() - 2 * 365 * 24 * 3600 * 1000);
        let resp   = tiFetch_(
          '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetDividends',
          { figi: figi, from: from2y.toISOString(), to: now.toISOString() }
        );
        let divs = resp.dividends || [];
        if (divs.length > 0) {
          // Берём последнюю выплату, смотрим месяц
          let lastDiv = divs[divs.length - 1];
          let dateRaw = lastDiv.paymentDate || lastDiv.recordDate;
          if (dateRaw) {
            let histDate;
            if (typeof dateRaw === 'string') {
              histDate = new Date(dateRaw);
            } else if (dateRaw.seconds) {
              histDate = new Date(Number(dateRaw.seconds) * 1000);
            }
            if (histDate && !isNaN(histDate.getTime())) {
              payMonth = histDate.getMonth(); // 0-11
              // Если исторический месяц уже прошёл в этом году → следующий год
              let estimDate = new Date(payYear, payMonth, 15);
              if (estimDate < now) {
                payYear = payYear + 1;
              }
            }
          }
        }
        Utilities.sleep(50);
      } catch(e) {
        console.warn('Календарь: дивидендная история для ' + p.name + ': ' + e.message);
      }
    }

    if (payMonth === null) {
      // Месяц неизвестен — добавляем без конкретной даты (не войдёт в список по датам)
      return;
    }

    let estimatedDate = new Date(payYear, payMonth, 15); // примерно 15-е число
    payments.push({
      name:    p.name,
      ticker:  p.ticker,
      type:    'Дивиденд ~',  // ~ означает приблизительно
      date:    estimatedDate,
      perUnit: divData.amount,
      qty:     p.qty,
      total:   divData.amount * p.qty,
    });
  });

  return payments;
}


// ════════════════════════════════════════════════════════════════════
// ПОСТРОЕНИЕ МЕСЯЧНОЙ СЕТКИ
// ════════════════════════════════════════════════════════════════════

/**
 * Группирует платежи по ключу "месяц_год".
 * Возвращает объект: { "1_2026": {coupons: X, dividends: Y}, ... }
 */
function buildMonthlyGrid_(payments, fromDate) {
  let grid = {};

  payments.forEach(function(p) {
    if (!p.date || isNaN(p.date.getTime())) return;
    if (p.date < fromDate) return;

    let month  = p.date.getMonth() + 1; // 1-12
    let year   = p.date.getFullYear();
    let key    = month + '_' + year;

    if (!grid[key]) {
      grid[key] = { coupons: 0, dividends: 0 };
    }

    if (p.type === 'Купон') {
      grid[key].coupons += p.total;
    } else {
      grid[key].dividends += p.total;
    }
  });

  return grid;
}
