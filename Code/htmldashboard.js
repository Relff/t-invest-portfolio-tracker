/**
 * htmldashboard.js — Кастомный HTML-дашборд с карточками (Chart.js)
 *
 * Открывается модальным окном, использует уже посчитанные данные
 * из Script Properties и листа «История баланса» — ничего заново
 * не запрашивает у T-Invest API, поэтому открывается мгновенно.
 *
 * Зависимости: readConfig_(), readPositions_() — dashboard.js
 *              ANALYTICS_XIRR_PROP, ANALYTICS_BENCH_PROP — benchmark.js
 *              SNAPSHOT_SHEET — snapshot.js
 *              AVGPRICE_PROP — avgprice.js
 *              INCOME_TOTALS_DATA — income.js
 *              DISCIPLINE_DATA — discipline.js
 *
 * Радар «Факт vs Цель» строится по ВСЕМ категориям из classTargets (даже
 * с нулевым фактом) — иначе форма радара нечестно скрывала бы недостающие
 * классы вместо того, чтобы показать явный провал по оси.
 */

function showHtmlDashboard() {
  let payload = getDashboardPayload_();
  let template = HtmlService.createTemplate(HTML_DASHBOARD_TEMPLATE_);
  template.dataJson = JSON.stringify(payload);
  let html = template.evaluate().setWidth(760).setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Дашборд портфеля');
}

function getDashboardPayload_() {
  let config;
  try { config = readConfig_(); }
  catch (e) { return { error: e.message }; }

  let positions = readPositions_(config);
  let totalRub  = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);

  let cats = Object.keys(config.classTargets);
  let allocation = cats.map(function(cat) {
    let sum = positions.filter(function(p) { return p.category === cat; })
                       .reduce(function(s, p) { return s + p.valueRub; }, 0);
    return { name: cat, value: Math.round(sum) };
  }).filter(function(a) { return a.value > 0; });

  // Для радара берём ВСЕ категории из целевой структуры (даже с нулевым фактом) —
  // иначе форма радара будет «дырявой» и нечестно скроет недостающие классы.
  let radarAllocation = cats.map(function(cat) {
    let sum = positions.filter(function(p) { return p.category === cat; })
                       .reduce(function(s, p) { return s + p.valueRub; }, 0);
    return {
      name: cat,
      actualPct: totalRub > 0 ? Math.round((sum / totalRub) * 1000) / 10 : 0,
      targetPct: Math.round((config.classTargets[cat] || 0) * 1000) / 10,
    };
  });

  let props = PropertiesService.getScriptProperties();
  let xirrRaw  = props.getProperty('ANALYTICS_XIRR_VALUE');
  let benchRaw = props.getProperty('ANALYTICS_BENCH_VALUE');

  let xirr = xirrRaw ? JSON.parse(xirrRaw).value : null;
  let bench = benchRaw ? JSON.parse(benchRaw) : null;

  // P/L по акциям — переиспользуем уже посчитанные данные из avgprice.js,
  // ничего заново не считаем и не ходим в API.
  let avgRaw = props.getProperty(AVGPRICE_PROP);
  let plStocks = [];
  if (avgRaw) {
    plStocks = JSON.parse(avgRaw)
      .filter(function(r) { return r.avgPrice > 0; })
      .map(function(r) {
        return {
          name: r.name.length > 22 ? r.name.substring(0, 22) + '…' : r.name,
          pl: Math.round(r.pl),
          plPct: Math.round(r.plPct * 1000) / 10,
        };
      })
      .sort(function(a, b) { return b.plPct - a.plPct; });
  }

  // Прогноз денежного потока на месяц вперёд — переиспользуем уже
  // посчитанные средние (доход/год из income.js, дисциплина пополнений
  // из discipline.js), новых запросов к API не делаем.
  let incomeTotalsRaw = props.getProperty('INCOME_TOTALS_DATA');
  let disciplineRaw   = props.getProperty('DISCIPLINE_DATA');
  let cashForecast = null;
  if (incomeTotalsRaw) {
    let inc = JSON.parse(incomeTotalsRaw);
    let incomeMonthly = Math.round(inc.grandTotal / 12);
    let contribMonthly = 0;
    if (disciplineRaw) {
      let disc = JSON.parse(disciplineRaw);
      contribMonthly = Math.max(0, Math.round(disc.avgNet));
    }
    cashForecast = {
      incomeMonthly: incomeMonthly,
      contribMonthly: contribMonthly,
      total: incomeMonthly + contribMonthly,
    };
  }

  // История из снапшот-листа
  let ss = SpreadsheetApp.getActive();
  let snapSh = ss.getSheetByName('История баланса');
  let history = [];
  if (snapSh && snapSh.getLastRow() > 1) {
    let lastCol = Math.max(snapSh.getLastColumn(), 9);
    let rows = snapSh.getRange(2, 1, snapSh.getLastRow() - 1, lastCol).getValues();
    let tz = Session.getScriptTimeZone();
    history = rows.map(function(r) {
      let dateVal = r[0] instanceof Date ? r[0] : new Date(r[0]);
      let xirrRaw = r[8];
      return {
        iso: Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd'),
        date: Utilities.formatDate(dateVal, tz, 'dd.MM'),
        total: Math.round(r[1]),
        contrib: Math.round(r[2]),
        income: Math.round(r[3]),
        xirr: (xirrRaw === '' || xirrRaw === undefined || xirrRaw === null) ? null : Number(xirrRaw),
      };
    });
  }

  // Max Drawdown — максимальная просадка от локального пика по общей
  // стоимости портфеля (по снапшотам). Считается по факту стоимости, поэтому
  // отражает и рыночное падение, и выводы средств вместе — это ожидаемое
  // упрощение для дневных данных, не «чистая» рыночная просадка.
  let maxDrawdown = null;
  if (history.length > 1) {
    let peak = -Infinity, peakDate = null, maxDD = 0, ddPeakDate = null, ddTroughDate = null;
    history.forEach(function(h) {
      if (h.total > peak) { peak = h.total; peakDate = h.date; }
      if (peak > 0) {
        let dd = (peak - h.total) / peak;
        if (dd > maxDD) { maxDD = dd; ddPeakDate = peakDate; ddTroughDate = h.date; }
      }
    });
    if (maxDD > 0) {
      maxDrawdown = {
        pct: Math.round(maxDD * 1000) / 10,
        peakDate: ddPeakDate,
        troughDate: ddTroughDate,
      };
    }
  }

  return {
    totalRub: Math.round(totalRub),
    xirr: xirr,
    bench: bench,
    allocation: allocation,
    radarAllocation: radarAllocation,
    maxDrawdown: maxDrawdown,
    plStocks: plStocks,
    cashForecast: cashForecast,
    history: history,
  };
}

const HTML_DASHBOARD_TEMPLATE_ = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
  <style>
    body { font-family: 'Google Sans', Roboto, Arial, sans-serif; background: #f1f3f4; margin: 0; padding: 20px; }
    .cards { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
    .card { background: #fff; border-radius: 14px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.12); flex: 1; min-width: 140px; }
    .card .label { font-size: 12px; color: #666; margin-bottom: 6px; }
    .card .value { font-size: 22px; font-weight: 700; color: #1a237e; }
    .value.pos { color: #1b5e20; }
    .value.neg { color: #b71c1c; }
    .dd-caption { font-size: 11px; color: #9aa0a6; margin: -12px 0 16px; text-align: right; }
    .forecast-box { background: #fff; border-radius: 14px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.12); margin-bottom: 16px; }
    .forecast-title { font-size: 14px; color: #1a237e; font-weight: 700; margin-bottom: 10px; }
    .forecast-row { display: flex; justify-content: space-between; font-size: 13px; color: #555; padding: 4px 0; }
    .forecast-row b { color: #1b5e20; font-weight: 700; }
    .forecast-row.total { border-top: 1px solid #eee; margin-top: 6px; padding-top: 8px; font-weight: 700; color: #1a237e; }
    .forecast-row.total b { font-size: 16px; }
    .forecast-note { font-size: 10px; color: #9aa0a6; margin-top: 8px; font-style: italic; }
    .chart-box { margin-bottom: 16px; }
    .pie-box { max-width: 320px; }
    .chart-box { background: #fff; border-radius: 14px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.12); flex: 1; }
    .chart-box h3 { margin: 0 0 10px; font-size: 14px; color: #1a237e; display: flex; justify-content: space-between; align-items: center; }
    .charts-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
    .charts-row .chart-box { margin-bottom: 0; min-width: 260px; }
    .empty { color: #999; text-align: center; padding: 40px; font-size: 13px; }
    .range-buttons { display: flex; gap: 4px; }
    .rangeBtn { border: 1px solid #dadce0; background: #fff; border-radius: 8px; padding: 3px 10px; font-size: 11px; cursor: pointer; color: #444; }
    .rangeBtn:hover { background: #f1f3f4; }
    .rangeBtn.active { background: #1565c0; color: #fff; border-color: #1565c0; }
    .watermark { text-align: center; font-size: 10px; color: #9aa0a6; margin-top: 14px; }
    .watermark a { color: #9aa0a6; }
  </style>
</head>
<body>
  <div id="root"></div>

  <script>
    const data = <?!= dataJson ?>;
    const root = document.getElementById('root');
    let growthChartInstance = null;
    let currentRangeDays = 180; // по умолчанию — 6 месяцев

    if (data.error) {
      root.innerHTML = '<div class="empty">⚠️ ' + data.error + '</div>';
    } else {
      let cardsHtml = '<div class="cards">';
      cardsHtml += card_('Общая стоимость', fmt_(data.totalRub) + ' ₽', '');
      if (data.xirr !== null) {
        let cls = data.xirr >= 0 ? 'pos' : 'neg';
        cardsHtml += card_('XIRR', (data.xirr >= 0 ? '+' : '') + data.xirr.toFixed(1) + '%', cls);
      } else {
        cardsHtml += card_('XIRR', 'н/д', '');
      }
      if (data.bench) {
        let diff = data.bench.actual - data.bench.hypothetical;
        let cls = diff >= 0 ? 'pos' : 'neg';
        cardsHtml += card_('vs IMOEX', (diff >= 0 ? '+' : '') + fmt_(diff) + ' ₽', cls);
      }
      if (data.maxDrawdown !== null) {
        cardsHtml += card_('Max просадка', '−' + data.maxDrawdown.pct.toFixed(1) + '%', 'neg');
      } else {
        cardsHtml += card_('Max просадка', 'н/д', '');
      }
      cardsHtml += '</div>';

      let ddCaptionHtml = data.maxDrawdown
        ? '<div class="dd-caption">Просадка: пик ' + data.maxDrawdown.peakDate + ' → дно ' + data.maxDrawdown.troughDate + '</div>'
        : '';

      let forecastHtml = '';
      if (data.cashForecast) {
        let f = data.cashForecast;
        forecastHtml =
          '<div class="forecast-box">' +
          '<div class="forecast-title">💰 Прогноз денежного потока на следующий месяц</div>' +
          '<div class="forecast-row"><span>Ожидаемый доход (купоны + дивиденды)</span><b>+' + fmt_(f.incomeMonthly) + ' ₽</b></div>' +
          '<div class="forecast-row"><span>Обычное пополнение (среднее за 12 мес)</span><b>+' + fmt_(f.contribMonthly) + ' ₽</b></div>' +
          '<div class="forecast-row total"><span>Итого ожидается</span><b>+' + fmt_(f.total) + ' ₽</b></div>' +
          '<div class="forecast-note">Оценка по средним — не гарантия: точные даты купонов/дивидендов смотрите в «Календаре выплат», пополнение — по вашей обычной дисциплине за год.</div>' +
          '</div>';
      }

      let rangeButtonsHtml =
        '<div class="range-buttons">' +
        '<button data-days="30" class="rangeBtn">1 мес</button>' +
        '<button data-days="180" class="rangeBtn active">6 мес</button>' +
        '<button data-days="365" class="rangeBtn">12 мес</button>' +
        '<button data-days="0" class="rangeBtn">Всё время</button>' +
        '</div>';

      root.innerHTML = cardsHtml + ddCaptionHtml + forecastHtml +
        '<div class="charts-row">' +
        '<div class="chart-box pie-box"><h3>Распределение по классам</h3><canvas id="pieChart"></canvas></div>' +
        '<div class="chart-box radar-box"><h3>Факт vs Цель</h3><canvas id="radarChart"></canvas></div>' +
        '</div>' +
        '<div class="chart-box"><h3>Рост портфеля во времени' + rangeButtonsHtml + '</h3><canvas id="growthChart"></canvas></div>' +
        '<div class="chart-box"><h3>История доходности (XIRR)</h3><canvas id="xirrChart"></canvas></div>' +
        '<div class="chart-box"><h3>P/L по акциям</h3><canvas id="plChart"></canvas></div>' +
        '<div class="watermark">t-invest-portfolio-tracker · <a href="https://github.com/Relff/t-invest-portfolio-tracker" target="_blank">github.com/Relff/t-invest-portfolio-tracker</a></div>';

      if (data.allocation.length) {
        new Chart(document.getElementById('pieChart'), {
          type: 'doughnut',
          data: {
            labels: data.allocation.map(a => a.name),
            datasets: [{
              data: data.allocation.map(a => a.value),
              backgroundColor: ['#1565c0', '#0d47a1', '#ffd54f', '#43a047', '#ef6c00'],
            }]
          },
          options: { plugins: { legend: { position: 'bottom', labels: { font: { size: 10 } } } } }
        });
      }

      if (data.radarAllocation && data.radarAllocation.length >= 3) {
        new Chart(document.getElementById('radarChart'), {
          type: 'radar',
          data: {
            labels: data.radarAllocation.map(a => a.name),
            datasets: [
              {
                label: 'Факт, %',
                data: data.radarAllocation.map(a => a.actualPct),
                backgroundColor: 'rgba(21, 101, 192, 0.25)',
                borderColor: '#1565c0',
                pointBackgroundColor: '#1565c0',
                borderWidth: 2,
              },
              {
                label: 'Цель, %',
                data: data.radarAllocation.map(a => a.targetPct),
                backgroundColor: 'rgba(239, 108, 0, 0.12)',
                borderColor: '#ef6c00',
                pointBackgroundColor: '#ef6c00',
                borderWidth: 2,
                borderDash: [5, 4],
              },
            ]
          },
          options: {
            plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, boxWidth: 10 } } },
            scales: {
              r: {
                beginAtZero: true,
                ticks: { font: { size: 9 }, backdropColor: 'transparent' },
                pointLabels: { font: { size: 10 } },
                grid: { color: '#e8eaed' },
              }
            }
          }
        });
      } else if (data.radarAllocation) {
        document.getElementById('radarChart').parentElement.innerHTML =
          '<h3>Факт vs Цель</h3><div class="empty">Нужно минимум 3 категории в целевой структуре для радара.</div>';
      }

      if (data.history.length > 1) {
        renderGrowthChart_(currentRangeDays);

        document.querySelectorAll('.rangeBtn').forEach(function(btn) {
          btn.addEventListener('click', function() {
            document.querySelectorAll('.rangeBtn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderGrowthChart_(Number(btn.dataset.days));
          });
        });
      } else {
        document.getElementById('growthChart').parentElement.innerHTML =
          '<h3>Рост портфеля во времени</h3><div class="empty">Пока мало данных — снапшот-лист копится по одной точке в день, график появится через пару недель.</div>';
      }

      let xirrPoints = data.history.filter(function(h) { return h.xirr !== null; });
      if (xirrPoints.length > 1) {
        new Chart(document.getElementById('xirrChart'), {
          type: 'line',
          data: {
            labels: xirrPoints.map(h => h.date),
            datasets: [{
              label: 'XIRR, %',
              data: xirrPoints.map(h => h.xirr),
              borderColor: '#1565c0',
              backgroundColor: 'rgba(21, 101, 192, 0.08)',
              fill: true,
              tension: 0.25,
              pointRadius: 2,
              pointBackgroundColor: '#1565c0',
            }]
          },
          options: {
            plugins: { legend: { display: false } },
            scales: {
              y: { ticks: { font: { size: 9 }, callback: v => v + '%' } },
              x: { ticks: { maxRotation: 60, minRotation: 60, font: { size: 8 } } },
            }
          }
        });
      } else {
        document.getElementById('xirrChart').parentElement.innerHTML =
          '<h3>История доходности (XIRR)</h3><div class="empty">Копится по одной точке в день начиная с сегодняшнего снапшота — график появится через пару недель.</div>';
      }

      if (data.plStocks && data.plStocks.length) {
        let plEl = document.getElementById('plChart');
        plEl.parentElement.style.height = (36 * data.plStocks.length + 40) + 'px';
        new Chart(plEl, {
          type: 'bar',
          data: {
            labels: data.plStocks.map(s => s.name),
            datasets: [{
              label: 'P/L, %',
              data: data.plStocks.map(s => s.plPct),
              backgroundColor: data.plStocks.map(s => s.plPct >= 0 ? '#1b5e20' : '#b71c1c'),
              borderRadius: 4,
              maxBarThickness: 22,
            }]
          },
          options: {
            indexAxis: 'y',
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: {
                callbacks: {
                  label: function(ctx) {
                    let s = data.plStocks[ctx.dataIndex];
                    return (s.plPct >= 0 ? '+' : '') + s.plPct.toFixed(1) + '%  (' +
                      (s.pl >= 0 ? '+' : '') + fmt_(s.pl) + ' ₽)';
                  }
                }
              }
            },
            scales: {
              x: { ticks: { font: { size: 9 }, callback: v => v + '%' } },
              y: { ticks: { font: { size: 10 } } },
            }
          }
        });
      } else {
        document.getElementById('plChart').parentElement.innerHTML =
          '<h3>P/L по акциям</h3><div class="empty">Сначала запусти «Средняя цена и P/L» в меню Tinkoff → 🎯 Аналитика.</div>';
      }
    }

    function renderGrowthChart_(days) {
      let filtered = data.history;
      if (days > 0) {
        let cutoff = new Date();
        cutoff.setDate(cutoff.getDate() - days);
        filtered = data.history.filter(function(h) {
          return new Date(h.iso) >= cutoff;
        });
      }
      if (!filtered.length) filtered = data.history.slice(-1);

      if (growthChartInstance) growthChartInstance.destroy();

      growthChartInstance = new Chart(document.getElementById('growthChart'), {
        type: 'bar',
        data: {
          labels: filtered.map(h => h.date),
          datasets: [
            { label: 'Вложено', data: filtered.map(h => h.contrib), backgroundColor: '#1565c0', stack: 's', maxBarThickness: 60 },
            { label: 'Доход/рост', data: filtered.map(h => h.total - h.contrib), backgroundColor: '#43a047', stack: 's', maxBarThickness: 60 },
          ]
        },
        options: {
          plugins: { legend: { position: 'top' } },
          scales: { x: { stacked: true, ticks: { maxRotation: 60, minRotation: 60, font: { size: 8 } } }, y: { stacked: true } }
        }
      });
    }

    function card_(label, value, cls) {
      return '<div class="card"><div class="label">' + label + '</div><div class="value ' + cls + '">' + value + '</div></div>';
    }
    function fmt_(n) {
      return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '\u00a0');
    }
  </script>
</body>
</html>
`;