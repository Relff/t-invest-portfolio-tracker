const HEALTH_PROP          = 'HEALTH_CONCENTRATION_DATA';
const HEALTH_SECTION_TITLE = '▌ HEALTH CHECK: КОНЦЕНТРАЦИЯ';

function calculateConcentrationHealth() {
  let config;
  try { config = readConfig_(); }
  catch (e) { SpreadsheetApp.getUi().alert('⚠️ ' + e.message); return; }

  let positions = readPositions_(config);
  if (!positions.length) { SpreadsheetApp.getUi().alert('Нет данных о позициях.'); return; }

  let totalRub = positions.reduce(function(s, p) { return s + p.valueRub; }, 0);
  let overall  = concentrationStats_(positions, totalRub);

  let shares      = positions.filter(function(p) { return p.category === 'Акции'; });
  let sharesTotal = shares.reduce(function(s, p) { return s + p.valueRub; }, 0);
  let sharesConc  = concentrationStats_(shares, sharesTotal);

  PropertiesService.getScriptProperties().setProperty(HEALTH_PROP, JSON.stringify({
    overall: overall, shares: sharesConc,
  }));
  redrawHealthSection_();
}

function concentrationStats_(positions, total) {
  if (!positions.length || total <= 0) return { top1Name: null, top1Pct: 0, top3Names: [], top3Pct: 0 };
  let sorted = positions.slice().sort(function(a, b) { return b.valueRub - a.valueRub; });
  let top3 = sorted.slice(0, 3);
  return {
    top1Name: sorted[0].name,
    top1Pct:  sorted[0].valueRub / total,
    top3Names: top3.map(function(p) { return p.name; }),
    top3Pct:   top3.reduce(function(s, p) { return s + p.valueRub; }, 0) / total,
  };
}

function healthColor_(pct, warnThr, critThr) {
  if (pct >= critThr) return C.CRIT;
  if (pct >= warnThr) return C.WARN;
  return C.OK;
}

function redrawHealthSection_() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DST.DASHBOARD);
  if (!sh) return;
  let raw = PropertiesService.getScriptProperties().getProperty(HEALTH_PROP);
  if (!raw) return;
  let d = JSON.parse(raw);
  let params = readAdvancedParams_();

  renderSection_(sh, HEALTH_SECTION_TITLE, function(sh, r, COLS) {
    function writeRow_(label, valueText, clr) {
      sh.getRange(r, 1, 1, 3).merge().setValue(label).setFontWeight('bold');
      sh.getRange(r, 4, 1, COLS - 3).merge().setValue(valueText)
        .setFontColor(clr).setFontWeight('bold').setHorizontalAlignment('right');
      sh.getRange(r, 1, 1, COLS).setBackground(C.EVEN);
      r++;
    }

    let ov = d.overall;
    if (ov.top1Name) {
      writeRow_('Крупнейшая позиция (весь портфель)', ov.top1Name + ' — ' + (ov.top1Pct * 100).toFixed(1) + '%',
        healthColor_(ov.top1Pct, params.healthTop1Warn / 100, params.healthTop1Crit / 100));
      writeRow_('Топ-3 позиции (весь портфель)', (ov.top3Pct * 100).toFixed(1) + '%',
        healthColor_(ov.top3Pct, params.healthTop3Warn / 100, params.healthTop3Crit / 100));
    }

    let sc = d.shares;
    if (sc.top1Name) {
      writeRow_('Крупнейшая акция (доля от акционной части)', sc.top1Name + ' — ' + (sc.top1Pct * 100).toFixed(1) + '%',
        healthColor_(sc.top1Pct, params.healthTop1Warn / 100, params.healthTop1Crit / 100));
      writeRow_('Топ-3 акции (доля от акционной части)', (sc.top3Pct * 100).toFixed(1) + '%',
        healthColor_(sc.top3Pct, params.healthTop3Warn / 100, params.healthTop3Crit / 100));
    }
  });
}