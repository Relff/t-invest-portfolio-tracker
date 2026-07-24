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
 */

function showHtmlDashboard() {
  let payload = getDashboardPayload_();
  let template = HtmlService.createTemplate(HTML_DASHBOARD_TEMPLATE_);
  template.dataJson = JSON.stringify(payload);
  let html = template.evaluate().setWidth(720).setHeight(760);
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

  let props = PropertiesService.getScriptProperties();
  let xirrRaw  = props.getProperty('ANALYTICS_XIRR_VALUE');
  let benchRaw = props.getProperty('ANALYTICS_BENCH_VALUE');

  let xirr = xirrRaw ? JSON.parse(xirrRaw).value : null;
  let bench = benchRaw ? JSON.parse(benchRaw) : null;

  // История из снапшот-листа
  let ss = SpreadsheetApp.getActive();
  let snapSh = ss.getSheetByName('История баланса');
  let history = [];
  if (snapSh && snapSh.getLastRow() > 1) {
    let rows = snapSh.getRange(2, 1, snapSh.getLastRow() - 1, 4).getValues();
    let tz = Session.getScriptTimeZone();
    history = rows.map(function(r) {
      let dateVal = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return {
        iso: Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd'),
        date: Utilities.formatDate(dateVal, tz, 'dd.MM'),
        total: Math.round(r[1]),
        contrib: Math.round(r[2]),
        income: Math.round(r[3]),
      };
    });
  }

  return {
    totalRub: Math.round(totalRub),
    xirr: xirr,
    bench: bench,
    allocation: allocation,
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
    .chart-box { margin-bottom: 16px; }
    .pie-box { max-width: 320px; }
    .chart-box { background: #fff; border-radius: 14px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.12); flex: 1; }
    .chart-box h3 { margin: 0 0 10px; font-size: 14px; color: #1a237e; display: flex; justify-content: space-between; align-items: center; }
    .empty { color: #999; text-align: center; padding: 40px; font-size: 13px; }
    .range-buttons { display: flex; gap: 4px; }
    .rangeBtn { border: 1px solid #dadce0; background: #fff; border-radius: 8px; padding: 3px 10px; font-size: 11px; cursor: pointer; color: #444; }
    .rangeBtn:hover { background: #f1f3f4; }
    .rangeBtn.active { background: #1565c0; color: #fff; border-color: #1565c0; }
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
      cardsHtml += '</div>';

      let rangeButtonsHtml =
        '<div class="range-buttons">' +
        '<button data-days="30" class="rangeBtn">1 мес</button>' +
        '<button data-days="180" class="rangeBtn active">6 мес</button>' +
        '<button data-days="365" class="rangeBtn">12 мес</button>' +
        '<button data-days="0" class="rangeBtn">Всё время</button>' +
        '</div>';

      root.innerHTML = cardsHtml +
        '<div class="chart-box pie-box"><h3>Распределение по классам</h3><canvas id="pieChart"></canvas></div>' +
        '<div class="chart-box"><h3>Рост портфеля во времени' + rangeButtonsHtml + '</h3><canvas id="growthChart"></canvas></div>';

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