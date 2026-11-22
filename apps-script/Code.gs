// CTG Dash App – manager scoring app
// Reads tasks from the DASH tab (Spanish layout) and writes scores straight into DASH (week columns G:P).

var SPREADSHEET_ID = '1ZCSHtPryT8Y13g7PL3pa-k62X2OC5j13YkYMh0R5maA';
var SHEET_NAME = 'DASH';
var MANAGERS = ['Dago', 'Jessica', 'Andrej'];

var NEW_APP_URL = 'https://cascadetropicals.github.io/ctg-dash-app/';

// The app now lives on GitHub Pages. This old address only forwards there and keeps the reminder-mail link parameters (?manager=Name&lang=es).
function doGet(e) {
  var p = (e && e.parameter) || {};
  // Design preview (test only): ?preview=1 shows the app page straight from Index.html
  if (p.preview === '1') return HtmlService.createHtmlOutputFromFile('Index').setTitle('CTG App').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  var q = [];
  if (p.manager) q.push('manager=' + encodeURIComponent(p.manager));
  if (p.lang) q.push('lang=' + encodeURIComponent(p.lang));
  var url = NEW_APP_URL + (q.length ? '?' + q.join('&') : '');
  var html = '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8"><title>CTG App</title>' +
    '<style>body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#f3f6f1;font-family:Arial,sans-serif;color:#17241b;text-align:center}' +
    '.c{background:#fff;border-radius:18px;padding:28px 24px;max-width:380px;box-shadow:0 8px 30px rgba(0,0,0,.12)}' +
    'a{display:inline-block;margin-top:16px;padding:12px 22px;border-radius:12px;background:#2f6b3a;color:#fff;font-weight:700;text-decoration:none}</style></head><body><div class="c">' +
    '<div style="font-size:22px;font-weight:800">CTG App</div>' +
    '<p>La app se mud\u00f3 a una nueva direcci\u00f3n.<br>The app has moved to a new address.</p>' +
    '<a href="' + url.replace(/&/g, '&amp;') + '" target="_top">Abrir / Open</a></div>' +
    '<script>try { window.top.location.replace(' + JSON.stringify(url) + '); } catch (err) {}<\/script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('CTG App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Managers, current week (DASH!B4) and week labels (from start date DASH!B2)
function getInit() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  var tz = ss.getSpreadsheetTimeZone();

  // Find "Start date" and "Current week" by their labels in column A (rows can move)
  var top = sh.getRange(1, 1, 30, 2).getValues();
  var start = '';
  var currentWeek = 1;
  for (var i = 0; i < top.length; i++) {
    var label = String(top[i][0]).trim().toLowerCase();
    if (label === 'fecha de inicio' || label === 'start date') start = top[i][1];
    if (label === 'semana actual' || label === 'current week') currentWeek = Number(top[i][1]) || 1;
  }
  if (currentWeek < 1) currentWeek = 1;
  if (currentWeek > 10) currentWeek = 10;

  var weeks = [];
  for (var w = 1; w <= 10; w++) {
    var weekDate = '';
    if (start instanceof Date) {
      var sp = Utilities.formatDate(new Date(start.getTime() + 43200000), tz, 'yyyy-MM-dd').split('-'); // start date in the sheet's time zone (+12 h absorbs any midnight offset)
      var d = new Date(Number(sp[0]), Number(sp[1]) - 1, Number(sp[2]) + (w - 1) * 7);
      weekDate = (d.getMonth() + 1) + '/' + d.getDate();
    }
    weeks.push({ n: w, date: weekDate });
  }

  return { managers: MANAGERS, currentWeek: currentWeek, weeks: weeks };
}

// Tasks for one manager and one week, grouped by employee. lang = 'es' (default) or 'en'.
function getTasks(manager, week, lang) {
  week = Number(week);
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var values = sh.getRange(1, 1, sh.getLastRow(), 16).getValues();

  var groups = [];
  var current = null;

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var a = String(row[0]).trim();
    var b = String(row[1]).trim();
    var f = String(row[5]).trim();

    // Block header row: col B says "Miles"
    if (a && /^(millas|miles)$/i.test(b)) {
      current = { name: a.replace(/^Tareas\s+/i, '').replace(/'s Tasks$/i, '').replace(/ Tasks$/i, '').trim(), tasks: [], otherSum: 0 };
      groups.push(current);
      continue;
    }
    if (!current) continue;

    // End of block
    if (/^Total/i.test(f)) {
      current = null;
      continue;
    }
    if (!a) continue;

    var reviewer = String(row[4]).trim();
    var isBonus = /^Bono|^Bonus/i.test(a);
    var mine = isBonus || reviewer.toLowerCase() === String(manager).toLowerCase();
    var existing = row[5 + week]; // G = index 6 = week 1

    // Tasks not shown to this manager (System, other reviewers) still count in the week total
    if (!mine) {
      current.otherSum += Number(existing) || 0;
      continue;
    }
    current.tasks.push({
      row: r + 1,
      task: a,
      miles: b,
      how: String(row[2]).trim(),
      when: f,
      isBonus: isBonus,
      value: existing === '' ? '' : existing
    });
  }

  // English view: task text comes from the English columns of the DD tab
  // (J = task, K = how we measure, L = when), matched by the Spanish task name in column A.
  if (String(lang).toLowerCase() === 'en') {
    var dd = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('DD').getDataRange().getValues();
    var en = {};
    var started = false;
    for (var d = 0; d < dd.length; d++) {
      var key = String(dd[d][0]).trim().toLowerCase();
      if (!started) { if (key === 'tareas') started = true; continue; }
      if (key && String(dd[d][9]).trim()) {
        en[key] = { task: String(dd[d][9]).trim(), how: String(dd[d][10]).trim(), when: String(dd[d][11]).trim() };
      }
    }
    groups.forEach(function (g) {
      g.tasks.forEach(function (t) {
        var e = en[String(t.task).trim().toLowerCase()];
        if (!e) return; // no English text yet: keep the Spanish
        t.task = e.task;
        t.how = e.how || t.how;
        t.when = e.when || t.when;
      });
    });
  }

  // Dago/Jessica: only employees they review (bonus alone is not enough). Andrej: also all bonus rows.
  var isAndrej = String(manager).toLowerCase() === 'andrej';
  return groups.filter(function (g) {
    var regular = g.tasks.some(function (t) { return !t.isBonus; });
    return regular || (isAndrej && g.tasks.length > 0);
  });
}

// Save scores: entries = [{row, value}] ; value '' clears the cell
function saveScores(manager, week, entries) {
  week = Number(week);
  if (!(week >= 1 && week <= 10)) throw new Error('Semana inválida / Invalid week');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var values = sh.getRange(1, 1, sh.getLastRow(), 16).getValues();

    // Rows this manager is allowed to write
    var allowed = {};
    var inBlock = false;
    for (var r = 0; r < values.length; r++) {
      var a = String(values[r][0]).trim();
      var b = String(values[r][1]).trim();
      var f = String(values[r][5]).trim();
      if (a && /^(millas|miles)$/i.test(b)) { inBlock = true; continue; }
      if (!inBlock) continue;
      if (/^Total/i.test(f)) { inBlock = false; continue; }
      if (!a) continue;
      var reviewer = String(values[r][4]).trim();
      if (/^Bono|^Bonus/i.test(a) || reviewer.toLowerCase() === String(manager).toLowerCase()) {
        allowed[r + 1] = true;
      }
    }

    var col = 6 + week; // G = column 7 = week 1
    var saved = 0;
    (entries || []).forEach(function (e) {
      var rowNum = Number(e.row);
      if (!allowed[rowNum]) return;
      var v = e.value;
      if (v === '' || v === null || v === undefined) {
        sh.getRange(rowNum, col).clearContent();
      } else {
        var n = Number(v);
        if (isNaN(n) || n < 0) return;
        sh.getRange(rowNum, col).setValue(n);
      }
      saved++;
    });
    return { saved: saved };
  } finally {
    lock.releaseLock();
  }
}

// Home screen overview: program week, dates, and each person's millas (all from DASH)
function getOverview() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  var tz = ss.getSpreadsheetTimeZone();
  var values = sh.getRange(1, 1, sh.getLastRow(), 16).getValues();

  // Start date + current week by label in column A
  var start = '';
  var currentWeek = 1;
  for (var i = 0; i < Math.min(30, values.length); i++) {
    var label = String(values[i][0]).trim().toLowerCase();
    if (label === 'fecha de inicio' || label === 'start date') start = values[i][1];
    if (label === 'semana actual' || label === 'current week') currentWeek = Number(values[i][1]) || 0;
  }
  var weekForData = Math.min(10, Math.max(1, currentWeek));

  // Goals from the summary table (header cell "Goal", names in the column to its left)
  var goals = {};
  for (var r = 0; r < Math.min(20, values.length); r++) {
    for (var c = 1; c < 16; c++) {
      if (String(values[r][c]).trim().toLowerCase() === 'meta' || String(values[r][c]).trim().toLowerCase() === 'goal') {
        for (var k = r + 1; k < Math.min(r + 30, values.length); k++) {
          var nm = String(values[k][c - 1]).trim();
          if (!nm) break;
          goals[nm.toLowerCase()] = Number(values[k][c]) || 0;
        }
        r = 99; break;
      }
    }
  }

  // Sum every task row per person (blocks: header row has "Miles" in col B, ends at "Total" in col F)
  var people = [];
  var cur = null;
  for (var j = 0; j < values.length; j++) {
    var row = values[j];
    var a = String(row[0]).trim();
    var b = String(row[1]).trim();
    var f = String(row[5]).trim();
    if (a && /^(millas|miles)$/i.test(b)) {
      var name = a.replace(/^Tareas\s+/i, '').replace(/'s Tasks$/i, '').replace(/ Tasks$/i, '').trim();
      cur = { name: name, total: 0, week: 0, bonus: 0, goal: goals[name.toLowerCase()] || 1000 };
      people.push(cur);
      continue;
    }
    if (!cur) continue;
    if (/^Total/i.test(f)) { cur = null; continue; }
    if (!a) continue;
    var isBonusRow = /^Bono|^Bonus/i.test(a);
    for (var w = 1; w <= 10; w++) {
      var v = Number(row[5 + w]) || 0;
      cur.total += v;
      if (isBonusRow) cur.bonus += v; // bonus is a real extra: kept in the total, left out of the pace comparison
      if (w === weekForData) cur.week += v;
    }
  }

  var out = {
    currentWeek: currentWeek,
    weekStart: '', weekEnd: '', programEnd: '', daysLeft: '',
    people: people,
    readAt: { ms: Date.now(), txt: Utilities.formatDate(new Date(), tz, 'M/d h:mm a z') }
  };

  if (start instanceof Date) {
    var sp = Utilities.formatDate(new Date(start.getTime() + 43200000), tz, 'yyyy-MM-dd').split('-'); // start date in the sheet's time zone (+12 h absorbs any midnight offset)
    var ws = new Date(Number(sp[0]), Number(sp[1]) - 1, Number(sp[2]) + (weekForData - 1) * 7);
    var we = new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() + 6);
    var pe = new Date(Number(sp[0]), Number(sp[1]) - 1, Number(sp[2]) + 69);
    var now = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    out.weekStart = (ws.getMonth() + 1) + '/' + ws.getDate();
    out.weekEnd = (we.getMonth() + 1) + '/' + we.getDate();
    out.programEnd = (pe.getMonth() + 1) + '/' + pe.getDate();
    out.daysLeft = Math.max(0, Math.round((pe - today) / 86400000) + 1);
  }
  return out;
}

// Performance overview: order seconds, finished planting and labels, last 12 weeks, all people.
// Order seconds and planting come from the two IMPORT tabs of this spreadsheet; labels are read from the CG Inventory App.
function getPerformance() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tz = ss.getSpreadsheetTimeZone();

  // Order seconds goal: first "Segundos de orden" row on the DASH tab (column D)
  var goal = 14;
  var dash = ss.getSheetByName(SHEET_NAME).getDataRange().getValues();
  for (var g = 0; g < dash.length; g++) {
    if (/segundos de orden|order seconds/i.test(String(dash[g][0])) && Number(dash[g][3]) > 0) { goal = Number(dash[g][3]); break; }
  }

  // Order seconds (team): header row 4, weeks in rows 5-16 (column B = week start)
  var os = ss.getSheetByName('IMPORT order seconds').getRange(4, 1, 14, 5).getValues();
  var weekRows = [];
  for (var r = 1; r < os.length; r++) {
    if (os[r][1] instanceof Date) weekRows.push(os[r]);
  }
  weekRows.sort(function (a, b) { return b[1] - a[1]; }); // newest week first (left side)

  var weeks = weekRows.map(function (row) {
    var d = row[1];
    return { key: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), n: String(row[0]).trim(), label: Utilities.formatDate(d, tz, 'M/d') };
  });

  // Order seconds for the whole team (not by person): C = team seconds per unit weighted by units picked (Wed + Fri), D = total units.
  var teamByWeek = {};
  weekRows.forEach(function (row) {
    var sec = row[2];
    if (sec === '' || sec === null || isNaN(Number(sec))) return;
    teamByWeek[Utilities.formatDate(row[1], tz, 'yyyy-MM-dd')] = { sec: Number(sec), units: Number(row[3]) || 0 };
  });
  var order = [{
    name: '',
    vals: weeks.map(function (w) { return teamByWeek[w.key] || null; })
  }];
  if (!order[0].vals.some(function (v) { return v !== null; })) order = [];

  // Finished planting (IMPORT tab, goals in row 4, header row 5, data from row 6): A = week start, B = person,
  // L = hours, P = % of goal speed (hours at goal speed / hours with a goal)
  var pl = ss.getSheetByName('IMPORT finished planting');
  var goalsFinished = [];
  var gnF = pl.getRange(4, 5, 2, 5).getValues(); // row 4 = goals per hour, row 5 = size names (E:I)
  for (var gf = 0; gf < 5; gf++) if (String(gnF[1][gf]).trim() && Number(gnF[0][gf]) > 0) goalsFinished.push({ name: String(gnF[1][gf]).trim(), goal: Number(gnF[0][gf]) });
  var plRows = pl.getLastRow() >= 6 ? pl.getRange(6, 1, pl.getLastRow() - 5, 17).getValues() : [];
  var byPerson = {};
  plRows.forEach(function (row) {
    if (!(row[0] instanceof Date)) return;
    var name = String(row[1]).trim();
    if (!name) return;
    var key = Utilities.formatDate(row[0], tz, 'yyyy-MM-dd');
    var hrs = Number(row[11]) || 0;
    var ratio = row[15];
    var pctF = Math.round(Number(ratio) * 100);
    var noteF = null; // 'few' = fewer hours than the minimum (flagged by the source tab), 'high' = 200% or more
    if (String(row[16]).trim()) noteF = { k: 'few' };
    else if (pctF >= 200) noteF = { k: 'high' };
    if (!byPerson[name]) byPerson[name] = {};
    byPerson[name][key] = (hrs > 0 && ratio !== '' && !isNaN(Number(ratio))) ? { pct: pctF, hrs: Math.round(hrs * 10) / 10, note: noteF } : null;
  });
  var planting = Object.keys(byPerson).map(function (name) {
    var vals = weeks.map(function (w) { return byPerson[name][w.key] || null; });
    return { name: name, vals: vals };
  }).filter(function (p) {
    return p.vals.some(function (v) { return v !== null; });
  });
  planting.sort(function (a, b) { return a.name.localeCompare(b.name); });

  // Plugs (IMPORT tab, goals in row 4, header row 5, data from row 6): A = week start, B = person,
  // J = hours, N = % of goal speed
  var pg = ss.getSheetByName('IMPORT plugs planting');
  var goalsPlugs = [];
  var gnP = pg.getRange(4, 5, 2, 3).getValues(); // row 4 = goals per hour, row 5 = source names (E:G)
  for (var gp = 0; gp < 3; gp++) if (String(gnP[1][gp]).trim() && Number(gnP[0][gp]) > 0) goalsPlugs.push({ name: String(gnP[1][gp]).trim(), goal: Number(gnP[0][gp]) });
  var pgRows = pg.getLastRow() >= 6 ? pg.getRange(6, 1, pg.getLastRow() - 5, 15).getValues() : [];
  var plugByPerson = {};
  pgRows.forEach(function (row) {
    if (!(row[0] instanceof Date)) return;
    var name = String(row[1]).trim();
    if (!name) return;
    var key = Utilities.formatDate(row[0], tz, 'yyyy-MM-dd');
    var hrs = Number(row[9]) || 0;
    var ratio = row[13];
    var pctP = Math.round(Number(ratio) * 100);
    var noteP = null;
    if (String(row[14]).trim()) noteP = { k: 'few' };
    else if (pctP >= 200) noteP = { k: 'high' };
    if (!plugByPerson[name]) plugByPerson[name] = {};
    plugByPerson[name][key] = (hrs > 0 && ratio !== '' && !isNaN(Number(ratio))) ? { pct: pctP, hrs: Math.round(hrs * 10) / 10, note: noteP } : null;
  });
  var plugs = Object.keys(plugByPerson).map(function (name) {
    var vals = weeks.map(function (w) { return plugByPerson[name][w.key] || null; });
    return { name: name, vals: vals };
  }).filter(function (p) {
    return p.vals.some(function (v) { return v !== null; });
  });
  plugs.sort(function (a, b) { return a.name.localeCompare(b.name); });

  // Labels (IMPORT tab, goals in row 4, header row 5, data from row 6): A = week start, B = person,
  // L = hours, P = % of goal speed
  var lb = ss.getSheetByName('IMPORT labels');
  var goalsLabels = [];
  var gnL = lb.getRange(4, 5, 2, 5).getValues(); // row 4 = goals per hour, row 5 = size names (E:I)
  for (var gl = 0; gl < 5; gl++) if (String(gnL[1][gl]).trim() && Number(gnL[0][gl]) > 0) goalsLabels.push({ name: String(gnL[1][gl]).trim(), goal: Number(gnL[0][gl]) });
  var labelRows = lb.getLastRow() >= 6 ? lb.getRange(6, 1, lb.getLastRow() - 5, 17).getValues() : [];
  var lblByPerson = {};
  labelRows.forEach(function (row) {
    if (!(row[0] instanceof Date)) return;
    var name = String(row[1]).trim();
    if (!name) return;
    var key = Utilities.formatDate(row[0], tz, 'yyyy-MM-dd');
    var hrs = Number(row[11]) || 0;
    var ratio = row[15];
    var pctL = Math.round(Number(ratio) * 100);
    var noteL = null;
    if (String(row[16]).trim()) noteL = { k: 'few' };
    else if (pctL >= 200) noteL = { k: 'high' };
    if (!lblByPerson[name]) lblByPerson[name] = {};
    lblByPerson[name][key] = (hrs > 0 && ratio !== '' && !isNaN(Number(ratio))) ? { pct: pctL, hrs: Math.round(hrs * 10) / 10, note: noteL } : null;
  });
  var labels = Object.keys(lblByPerson).map(function (name) {
    var vals = weeks.map(function (w) { return lblByPerson[name][w.key] || null; });
    return { name: name, vals: vals };
  }).filter(function (p) {
    return p.vals.some(function (v) { return v !== null; });
  });
  labels.sort(function (a, b) { return a.name.localeCompare(b.name); });

  // Only people in the race (a block on the DASH tab: header row with "Millas" in column B)
  var raceNames = {};
  dash.forEach(function (dr) {
    var dn = String(dr[0]).trim();
    if (dn && /^(millas|miles)$/i.test(String(dr[1]).trim())) raceNames[dn.replace(/^Tareas\s+/i, '').replace(/'s Tasks$/i, '').replace(/ Tasks$/i, '').trim().toLowerCase()] = true;
  });
  function inRace(p) { return raceNames[String(p.name).trim().toLowerCase()] === true; }
  var extraShown = { blanca: true, jose: true }; // shown in Finished and Plugs even though not in the race (no scores)
  function inRaceOrExtra(p) { return inRace(p) || extraShown[String(p.name).trim().toLowerCase()] === true; }
  planting = planting.filter(inRaceOrExtra);
  plugs = plugs.filter(inRaceOrExtra);
  labels = labels.filter(inRace);

  // Planting entries (IMPORT tab, header row 4, data from row 5): A source, B date, C time, D person, E qty, F plant, G size,
  // I entry type, M goal per hour, N status, O plants counted, P hours credited, Q hours at goal speed
  var entries = [];
  var en = ss.getSheetByName('IMPORT planting entries');
  var enRows = (en && en.getLastRow() >= 5) ? en.getRange(5, 1, en.getLastRow() - 4, 17).getDisplayValues() : [];
  var enLb = ss.getSheetByName('IMPORT label entries'); // label records have their own tab, same columns
  if (enLb && enLb.getLastRow() >= 5) enRows = enRows.concat(enLb.getRange(5, 1, enLb.getLastRow() - 4, 17).getDisplayValues());
  function num(x) { var n = Number(String(x).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
  var weekStarts = weeks.map(function (w) { return w.key; });
  enRows.forEach(function (r) {
    var m = String(r[1]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m || !String(r[3]).trim()) return;
    var src = /^fin/i.test(r[0]) ? 'planting' : (/^plug/i.test(r[0]) ? 'plugs' : (/^lab/i.test(r[0]) ? 'labels' : ''));
    if (!src) return;
    var dk = m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    var wk = '';
    for (var wi = 0; wi < weekStarts.length; wi++) {
      var ws = weekStarts[wi].split('-');
      var t0 = Date.UTC(Number(ws[0]), Number(ws[1]) - 1, Number(ws[2]));
      var t1 = Date.UTC(Number(m[3]), Number(m[1]) - 1, Number(m[2]));
      if (t1 >= t0 && t1 < t0 + 7 * 86400000) { wk = weekStarts[wi]; break; }
    }
    entries.push([src, dk, String(r[2]).trim(), String(r[3]).trim(), String(r[4]).trim() === '' ? '' : num(r[4]), String(r[5]).trim(), String(r[6]).trim(),
      String(r[8]).trim(), String(r[13]).trim(), num(r[14]), Math.round(num(r[15]) * 100) / 100, Math.round(num(r[16]) * 100) / 100, num(r[12]), wk, String(r[9]).trim(), String(r[10]).trim() === '' ? null : num(r[10])]);
  });
  entries.sort(function (a, b) { return a[1] < b[1] ? 1 : (a[1] > b[1] ? -1 : 0); }); // newest day first

  // Weighted average of the last 28 days per person: hours at goal speed / hours with a goal
  var tp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-');
  var cut = new Date(Date.UTC(Number(tp[0]), Number(tp[1]) - 1, Number(tp[2]) - 27));
  var cutKey = Utilities.formatDate(cut, 'UTC', 'yyyy-MM-dd');
  var avgs = { planting: {}, plugs: {}, labels: {} };
  entries.forEach(function (e) {
    if (e[1] < cutKey || e[7] === 'Marker' || !(e[12] > 0)) return;
    var a = avgs[e[0]][e[3]] || (avgs[e[0]][e[3]] = { atGoal: 0, hrs: 0 });
    a.atGoal += e[11]; a.hrs += e[10];
  });
  function addAvg(list, src) {
    list.forEach(function (p) {
      var a = avgs[src][p.name];
      p.avg = (a && a.hrs > 0) ? { pct: Math.round(a.atGoal / a.hrs * 100), hrs: Math.round(a.hrs * 10) / 10 } : null;
    });
  }
  addAvg(planting, 'planting'); addAvg(plugs, 'plugs'); addAvg(labels, 'labels');

  // Newest record in the data (date + time), so people can see how fresh the numbers are
  var latest = null, latestKey = '';
  entries.forEach(function (e) {
    var tm = String(e[2]).match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i), mins = 0;
    if (tm) { var hh = Number(tm[1]) % 12; if (/pm/i.test(tm[3] || '')) hh += 12; else if (!tm[3]) hh = Number(tm[1]); mins = hh * 60 + Number(tm[2]); }
    var lk = e[1] + ('000' + mins).slice(-4);
    if (lk > latestKey) { latestKey = lk; latest = { date: e[1], time: e[2] }; }
  });

  return {
    goal: goal,
    readAt: { ms: Date.now(), txt: Utilities.formatDate(new Date(), tz, 'M/d h:mm a z') },
    latest: latest,
    weeks: weeks.map(function (w) { return { n: w.n, label: w.label, key: w.key }; }),
    order: order,
    planting: planting,
    plugs: plugs,
    labels: labels,
    entries: entries,
    goals: { planting: goalsFinished, plugs: goalsPlugs, labels: goalsLabels }
  };
}


// "Today" view: the planting records of the last 7 days from the tab IMPORT planting entries, one array per record in the same layout as getPerformance().entries
// (week key left empty). Cached for 45 seconds so many viewers do not each read the sheet.
function getToday() {
  var cache = CacheService.getScriptCache();
  try { var hit = cache.get('today1'); if (hit) return JSON.parse(hit); } catch (e1) {}
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), tz = ss.getSpreadsheetTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), q = todayKey.split('-');
  var cutKey = Utilities.formatDate(new Date(Date.UTC(Number(q[0]), Number(q[1]) - 1, Number(q[2]) - 7)), 'UTC', 'yyyy-MM-dd');
  var en = ss.getSheetByName('IMPORT planting entries');
  var rows = (en && en.getLastRow() >= 5) ? en.getRange(5, 1, en.getLastRow() - 4, 17).getDisplayValues() : [];
  var enLb = ss.getSheetByName('IMPORT label entries'); // label records have their own tab, same columns
  if (enLb && enLb.getLastRow() >= 5) rows = rows.concat(enLb.getRange(5, 1, enLb.getLastRow() - 4, 17).getDisplayValues());
  function num(x) { var n = Number(String(x).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
  var entries = [], latest = null, latestKey = '';
  rows.forEach(function (r) {
    var m = String(r[1]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m || !String(r[3]).trim()) return;
    var src = /^fin/i.test(r[0]) ? 'planting' : (/^plug/i.test(r[0]) ? 'plugs' : (/^lab/i.test(r[0]) ? 'labels' : ''));
    if (!src) return;
    var dk = m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    if (dk < cutKey || dk > todayKey) return;
    entries.push([src, dk, String(r[2]).trim(), String(r[3]).trim(), String(r[4]).trim() === '' ? '' : num(r[4]), String(r[5]).trim(), String(r[6]).trim(),
      String(r[8]).trim(), String(r[13]).trim(), num(r[14]), Math.round(num(r[15]) * 100) / 100, Math.round(num(r[16]) * 100) / 100, num(r[12]), '', String(r[9]).trim(), String(r[10]).trim() === '' ? null : num(r[10])]);
    var tm = String(r[2]).match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i), mins = 0;
    if (tm) { var hh = Number(tm[1]) % 12; if (/pm/i.test(tm[3] || '')) hh += 12; else if (!tm[3]) hh = Number(tm[1]); mins = hh * 60 + Number(tm[2]); }
    var lk = dk + ('000' + mins).slice(-4);
    if (lk > latestKey) { latestKey = lk; latest = { date: dk, time: String(r[2]).trim() }; }
  });
  var out = { today: todayKey, entries: entries, latest: latest, readAt: { ms: Date.now(), txt: Utilities.formatDate(new Date(), tz, 'M/d h:mm a z') } };
  try { cache.put('today1', JSON.stringify(out), 45); } catch (e2) {}
  return out;
}


// "Data accuracy" view: per week (Monday start) counts of record issues from IMPORT planting entries (Finished + Plugs), and % good data.
// Real records = entry type not Marker. Good = status OK. Start records ('Empieza', no qty) are normal and not an issue; qty 1 on a real plant instead of Empieza is counted as marker. Cached 5 minutes.
function getAccuracy() {
  var cache = CacheService.getScriptCache();
  try { var hit = cache.get('acc2'); if (hit) return JSON.parse(hit); } catch (e1) {}
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), tz = ss.getSpreadsheetTimeZone();
  var en = ss.getSheetByName('IMPORT planting entries');
  var rows = (en && en.getLastRow() >= 5) ? en.getRange(5, 1, en.getLastRow() - 4, 17).getDisplayValues() : [];
  var wk = {};
  rows.forEach(function (r) {
    var m = String(r[1]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m || !String(r[3]).trim()) return;
    var d = new Date(Date.UTC(Number(m[3]), Number(m[1]) - 1, Number(m[2])));
    var dow = (d.getUTCDay() + 6) % 7; // Monday = 0
    var key = Utilities.formatDate(new Date(d.getTime() - dow * 86400000), 'UTC', 'yyyy-MM-dd');
    var w = wk[key] = wk[key] || { week: key, all: 0, real: 0, ok: 0, start: 0, marker: 0, fast: 0, slow: 0, first: 0, repeat: 0 };
    w.all++;
    var st = String(r[13]).trim();
    if (String(r[8]).trim() === 'Marker' || st === 'Marker') {
      // a proper start record is the plant 'Empieza' with no quantity: not an issue. A start made with qty 1 on a real plant is the issue.
      if (/^empieza$/i.test(String(r[5]).trim()) || String(r[4]).trim() === '') w.start++; else w.marker++;
      return;
    }
    w.real++;
    if (st === 'OK') w.ok++;
    else if (/too fast/i.test(st)) w.fast++;
    else if (/too slow/i.test(st)) w.slow++;
    else if (/first entry/i.test(st)) w.first++;
    if (String(r[11]).trim() === '1') w.repeat++;
  });
  var weeks = Object.keys(wk).sort().slice(-12).map(function (k) { return wk[k]; });
  var out = { weeks: weeks, readAt: { ms: Date.now(), txt: Utilities.formatDate(new Date(), tz, 'M/d h:mm a z') } };
  try { cache.put('acc2', JSON.stringify(out), 300); } catch (e2) {}
  return out;
}

// "Timeline" view: per person and day, when each kind of work happened. Planting, plugs and labels come from the IMPORT tabs of this spreadsheet (one record per point);
// order pulling comes from CG BLANCA (tab 'performance': per day and picker, first and last pick), so pulling is one block per person and day. Last 14 days. Cached 60 seconds.
var PULL_SS_ID = '1PxfE7QZfGWrJXTupcwfLZJ3bQ06VKzMPw6SsPQnkJ0o';
var AVAIL_SS_ID = '1_q9SAZVD_GMBw1TVWCcXXcec5iIPKnV2GIclzUwunmI'; // CG availability sheet, tab AVAIL CG: column BA = last edited in app (time), BB = last edited by
// Cache for big results (a cache entry holds at most 100 KB): the JSON is gzipped, turned into text and split in 90 KB pieces.
function tlCachePut(key, obj, sec) {
  try {
    var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(obj), 'application/octet-stream', 'x')).getBytes());
    var n = Math.ceil(b64.length / 90000), all = {};
    for (var i = 0; i < n; i++) all[key + '_' + i] = b64.substr(i * 90000, 90000);
    all[key] = String(n);
    CacheService.getScriptCache().putAll(all, sec);
  } catch (e) {}
}
function tlCacheGet(key) {
  try {
    var c = CacheService.getScriptCache(), n = Number(c.get(key));
    if (!n) return null;
    var keys = [], i;
    for (i = 0; i < n; i++) keys.push(key + '_' + i);
    var got = c.getAll(keys), s = '';
    for (i = 0; i < n; i++) { if (!got[key + '_' + i]) return null; s += got[key + '_' + i]; }
    return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(s), 'application/x-gzip', 'x')).getDataAsString());
  } catch (e) { return null; }
}
function getTimeline() {
  var cache = CacheService.getScriptCache();
  var hitTl = tlCacheGet('tl5'); if (hitTl) return hitTl;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), tz = ss.getSpreadsheetTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), q = todayKey.split('-');
  var cutKey = Utilities.formatDate(new Date(Date.UTC(Number(q[0]), Number(q[1]) - 1, Number(q[2]) - 14)), 'UTC', 'yyyy-MM-dd');
  function num(x) { var n = Number(String(x).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
  function mins(s) {
    var m = String(s).match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
    if (!m) return -1;
    var h = Number(m[1]) % 12; if (/pm/i.test(m[3] || '')) h += 12; else if (!m[3]) h = Number(m[1]);
    return h * 60 + Number(m[2]);
  }
  // read only from the first row of the last 14 days down (the tabs hold 6 weeks and more), not the whole tab
  function recentRows(sh) {
    var last = sh ? sh.getLastRow() : 0;
    if (last < 5) return [];
    var col = sh.getRange(5, 2, last - 4, 1).getDisplayValues(), first = -1;
    for (var i = 0; i < col.length; i++) {
      var dm = String(col[i][0]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (dm && (dm[3] + '-' + ('0' + dm[1]).slice(-2) + '-' + ('0' + dm[2]).slice(-2)) >= cutKey) { first = i; break; }
    }
    if (first < 0) return [];
    return sh.getRange(5 + first, 1, last - 4 - first, 17).getDisplayValues();
  }
  var rows = recentRows(ss.getSheetByName('IMPORT planting entries')).concat(recentRows(ss.getSheetByName('IMPORT label entries')));
  var recs = [];
  rows.forEach(function (r) {
    var m = String(r[1]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m || !String(r[3]).trim()) return;
    var src = /^fin/i.test(r[0]) ? 'planting' : (/^plug/i.test(r[0]) ? 'plugs' : (/^lab/i.test(r[0]) ? 'labels' : ''));
    if (!src) return;
    var dk = m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    if (dk < cutKey || dk > todayKey) return;
    var tm = mins(r[2]); if (tm < 0) return;
    var gap = String(r[10]).trim() === '' ? 0 : Math.round(num(r[10]) * 60);
    var mk = String(r[8]).trim() === 'Marker' || String(r[13]).trim() === 'Marker';
    // [source, day, person, minute of day, qty planted, status, minutes of credited work before this record, 1 = start marker, plant, qty entered, size, hours credited, hours at goal speed, minute of the previous record (-1 if none that day), goal per hour]
    recs.push([src, dk, String(r[3]).trim(), tm, num(r[14]), String(r[13]).trim(), (!mk && String(r[13]).trim() === 'OK') ? gap : 0, mk ? 1 : 0, String(r[5]).trim(), String(r[4]).trim() === '' ? '' : num(r[4]), String(r[6]).trim(), Math.round(num(r[15]) * 100) / 100, Math.round(num(r[16]) * 100) / 100, (String(r[9]).trim().split(' ')[0] === String(r[1]).trim()) ? mins(String(r[9]).trim().split(' ').slice(1).join(' ')) : -1, num(r[12])]);
  });
  var pa = tlCacheGet('tl5pa'), paHit = !!pa; // order pulling and CG Avail change slowly: cached 10 minutes
  var pulls = paHit ? pa.pulls : [], pullErr = paHit ? pa.pullErr : '';
  if (!paHit) try {
    var ps = SpreadsheetApp.openById(PULL_SS_ID).getSheetByName('performance');
    var lc = ps.getLastColumn(), v = ps.getRange(11, 1, 40, lc).getDisplayValues();
    v.forEach(function (r) {
      var m = String(r[1]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
      if (!m) return;
      var yy = Number(m[3]); if (yy < 100) yy += 2000;
      var dk = yy + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
      if (dk < cutKey || dk > todayKey) return;
      for (var c = 14; c + 9 < lc; c += 10) { // Picker blocks of 10 columns from column O: email, Name, Ayudante, Picks, First, Last, Hours, Breaks, Qty, Seconds
        var nm = String(r[c + 1]).trim(), a = mins(r[c + 4]), b = mins(r[c + 5]);
        if (!nm || a < 0 || b < 0) continue;
        pulls.push([dk, nm, a, b, num(r[c + 3]), num(r[c + 8]), Math.round(num(r[c + 6]) * 100) / 100, num(r[c + 9])]); // day, person, first pick, last pick, picks sent, qty picked, hours picked, seconds per unit
      }
    });
  } catch (e3) { pullErr = String(e3.message || e3); }
  // CG Avail: Jessica's edits, one block per day from her first to her last edit (the sheet keeps only the latest edit time per row)
  var avail = paHit ? pa.avail : [], availErr = paHit ? pa.availErr : '';
  if (!paHit) try {
    var av = SpreadsheetApp.openById(AVAIL_SS_ID).getSheetByName('AVAIL CG');
    var vv = av.getRange(4, 53, Math.max(1, av.getLastRow() - 3), 2).getDisplayValues();
    var byDay = {};
    vv.forEach(function (r) {
      var m = String(r[0]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
      if (!m || !/^jessica/i.test(String(r[1]).trim())) return;
      var dk = m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
      if (dk < cutKey || dk > todayKey) return;
      var mm = Number(m[4]) * 60 + Number(m[5]), o = byDay[dk] = byDay[dk] || { a: mm, b: mm, n: 0 };
      o.a = Math.min(o.a, mm); o.b = Math.max(o.b, mm); o.n++;
    });
    Object.keys(byDay).forEach(function (k) { avail.push([k, 'Jessica', byDay[k].a, byDay[k].b, byDay[k].n]); }); // day, person, first edit, last edit, rows edited
  } catch (e4) { availErr = String(e4.message || e4); }
  var out = { today: todayKey, recs: recs, pulls: pulls, pullErr: pullErr, avail: avail, availErr: availErr, readAt: { ms: Date.now(), txt: Utilities.formatDate(new Date(), tz, 'M/d h:mm a z') } };
  if (!paHit && !pullErr && !availErr) tlCachePut('tl5pa', { pulls: pulls, pullErr: pullErr, avail: avail, availErr: availErr }, 600);
  tlCachePut('tl5', out, 120);
  return out;
}

// Raw planting records for the "See raw data" view: every column of the IMPORT planting entries tab, newest first
function getRawEntries(src) {
  var en = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(src === 'labels' ? 'IMPORT label entries' : 'IMPORT planting entries'); // labels have their own records
  if (!en || en.getLastRow() < 4) return { head: [], rows: [] };
  var vals = en.getRange(4, 1, en.getLastRow() - 3, 17).getDisplayValues();
  var head = vals[0];
  function sortKey(r) {
    var m = String(r[1]).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    var t = String(r[2]).match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
    var mins = 0;
    if (t) { var h = Number(t[1]) % 12; if (/pm/i.test(t[3] || '')) h += 12; else if (!t[3]) h = Number(t[1]); mins = h * 60 + Number(t[2]); }
    return m ? (m[3] + ('0' + m[1]).slice(-2) + ('0' + m[2]).slice(-2) + ('000' + mins).slice(-4)) : '';
  }
  var rows = vals.slice(1).filter(function (r) { return String(r[3]).trim() !== ''; });
  rows.sort(function (a, b) { var ka = sortKey(a), kb = sortKey(b); return ka < kb ? 1 : (ka > kb ? -1 : 0); });
  return { head: head, rows: rows };
}


// Scores screen: each person's DASH table (points per task and week) plus the data behind the system-scored cells.
// Order seconds rows and planting goal rows are recognised by the IMPORT tab their formula reads; every other row is typed in by a reviewer.
// lang 'en' takes the task texts from the DASH EN tab (same layout as DASH).
function getScoreTable(lang) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tz = ss.getSpreadsheetTimeZone();
  var sh = ss.getSheetByName(SHEET_NAME);
  var rng = sh.getDataRange();
  var raw = rng.getValues();
  var shown = rng.getDisplayValues();
  var forms = rng.getFormulas();
  var enSh = lang === 'en' ? ss.getSheetByName('DASH EN') : null;
  var txt = enSh ? enSh.getDataRange().getDisplayValues() : shown;

  function key(d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  function num(x) { return x === '' || x === null || isNaN(Number(x)) ? null : Number(x); }

  // Start date by label in column A, then the 10 weeks (Monday keys) and whether each week is over (system scores appear after it)
  var start = null;
  for (var s = 0; s < Math.min(raw.length, 30); s++) {
    var lb = String(raw[s][0]).trim().toLowerCase();
    if ((lb === 'fecha de inicio' || lb === 'start date') && raw[s][1] instanceof Date) start = raw[s][1];
  }
  if (!start) throw new Error('Start date not found on the DASH tab');
  var sp = key(start).split('-');
  function addDays(n) { return Utilities.formatDate(new Date(Date.UTC(Number(sp[0]), Number(sp[1]) - 1, Number(sp[2]) + n)), 'UTC', 'yyyy-MM-dd'); }
  function md(k) { var q = k.split('-'); return Number(q[1]) + '/' + Number(q[2]); }
  var todayKey = key(new Date());
  var weeks = [];
  for (var w = 1; w <= 10; w++) {
    var k0 = addDays(7 * (w - 1));
    weeks.push({ n: w, key: k0, label: md(k0) + ' \u2013 ' + md(addDays(7 * (w - 1) + 6)), over: todayKey >= addDays(7 * w), endLabel: md(addDays(7 * w)) });
  }

  // Team order seconds per week (IMPORT order seconds: A week #, B week start, C team sec/unit, D units)
  var orders = {};
  var osh = ss.getSheetByName('IMPORT order seconds');
  if (osh) osh.getRange(5, 1, 12, 5).getValues().forEach(function (r) {
    if (r[1] instanceof Date && num(r[2]) !== null) orders[key(r[1])] = { sec: Number(r[2]), units: Number(r[3]) || 0 };
  });

  // Planting per person and week: Finished (L hours, O hours at goal speed) and Plugs (J hours, M hours at goal speed)
  var plantAll = {}, pnames = {};
  function addPlant(tab, hCol, gCol, which) {
    var t = ss.getSheetByName(tab);
    if (!t || t.getLastRow() < 6) return;
    t.getRange(6, 1, t.getLastRow() - 5, 15).getValues().forEach(function (r) {
      if (!(r[0] instanceof Date)) return;
      var nm = String(r[1]).trim();
      if (!nm) return;
      var lk = nm.toLowerCase();
      if (!pnames[lk]) pnames[lk] = nm;
      var byWeek = plantAll[lk] || (plantAll[lk] = {});
      var wk = byWeek[key(r[0])] || (byWeek[key(r[0])] = {});
      var x = wk[which] || (wk[which] = { h: 0, g: 0 });
      x.h += Number(r[hCol]) || 0;
      x.g += Number(r[gCol]) || 0;
    });
  }
  addPlant('IMPORT finished planting', 11, 14, 'f');
  addPlant('IMPORT plugs planting', 9, 12, 'p');
  addPlant('IMPORT labels', 11, 14, 'l'); // Labels: L hours, O hours at goal speed

  // One block per person on DASH: header row (name in A, "Millas" in B) down to the "Total:" row (label in column F)
  var people = [], plant = {};
  for (var i = 0; i < raw.length; i++) {
    var hn = String(shown[i][0]).trim();
    if (!hn || !/^(millas|miles)$/i.test(String(shown[i][1]).trim())) continue;
    var name = hn.replace(/^Tareas\s+/i, '').replace(/'s Tasks$/i, '').replace(/ Tasks$/i, '').trim();
    var rows = [], total = null;
    for (var j = i + 1; j < raw.length; j++) {
      if (/^total:?$/i.test(String(shown[j][5]).trim())) { total = raw[j].slice(6, 16).map(num); break; }
      var ta = String(shown[j][0]).trim();
      if (!ta && !String(shown[j][1]).trim() && !String(shown[j][4]).trim()) break;
      var f = '';
      for (var c = 6; c < 16; c++) if (forms[j][c]) { f = forms[j][c]; break; }
      var kind = /IMPORT order seconds/i.test(f) ? 'order' : (/IMPORT (finished|plugs) planting|IMPORT labels/i.test(f) ? 'planting' : 'typed');
      var tx = (txt[j] && String(txt[j][1]).trim() === String(shown[j][1]).trim()) ? txt[j] : shown[j];
      rows.push({
        task: String(tx[0]).trim(), pts: Number(raw[j][1]) || 0, how: String(tx[2]).trim(), goal: Number(raw[j][3]) || 0,
        who: String(tx[4]).trim(), when: String(tx[5]).trim(), bonus: /^(bono|bonus)/i.test(ta), kind: kind,
        v: raw[j].slice(6, 16).map(num),
        fm: forms[j].slice(6, 16).map(function (x) { return !!x; })
      });
    }
    var lk2 = name.toLowerCase();
    people.push({ name: name, key: lk2, pn: pnames[lk2] || name, rows: rows, total: total });
    if (plantAll[lk2]) plant[lk2] = plantAll[lk2];
  }

  return { weeks: weeks, orders: orders, plant: plant, people: people };
}


// ---------- Start tab: links by theme, featured links, favorites per person ----------
// Sheet tab 'START links': rows 1-3 notes, header row 4, data from row 5.
// Columns: A Theme, B Title, C Link, D Description, E Order, F Added by, G Date, H Featured (x = shown in "Lo mas importante"), I Icon (one emoji), J Description EN (English description; D is the Spanish one), K Featured order (position inside Most important). Tab 'START themes' (created by the app): A Theme, B Theme (English), C Icon, D Order (card position)
// Sheet tab 'START favorites': rows 1-3 notes, header row 4, data from row 5. Columns: A Person, B Title, C Link, D Date

// Names that can save favorites: the people on the DASH tab (same rule as getOverview), Blanca, Jose and the managers
function startPeople() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  var out = [], seen = {};
  function add(n) { n = String(n).trim(); if (n && !seen[n.toLowerCase()]) { seen[n.toLowerCase()] = true; out.push(n); } }
  vals.forEach(function (r) {
    var a = String(r[0]).trim(), b = String(r[1]).trim();
    if (a && /^(millas|miles)$/i.test(b)) add(a.replace(/^Tareas\s+/i, '').replace(/'s Tasks$/i, '').replace(/ Tasks$/i, ''));
  });
  add('Blanca'); add('Jose');
  MANAGERS.forEach(add);
  out.sort(function (a, b) { return a.localeCompare(b); });
  return out;
}

// Links saved as favorites by one person
function startFavs(person) {
  person = String(person || '').trim().toLowerCase();
  if (!person) return [];
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('START favorites');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 5) return [];
  var v = sh.getRange(5, 1, last - 4, 3).getDisplayValues(), out = [];
  v.forEach(function (r) { if (String(r[0]).trim().toLowerCase() === person && String(r[2]).trim()) out.push(String(r[2]).trim()); });
  return out;
}

// All links + the list of people + the favorites of "person" (optional) + the themes in display order
function getLinks(person) {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('START links');
  if (!sh) return { links: [], missing: true, people: [], favs: [], themes: [] };
  var rows = startRows_(sh);
  var links = rows.map(function (r) {
    return { row: r.row, theme: r.theme, title: r.title, link: r.link, desc: r.desc, descEn: r.descEn, order: r.order, feat: r.feat, featOrd: r.featOrd, icon: r.icon };
  });
  var themes = startThemeState_(rows).map(function (e) { return { name: e.name, nameEn: e.nameEn, icon: e.icon }; });
  return { links: links, missing: false, people: startPeople(), favs: startFavs(person), themes: themes };
}

function startSafe_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; } // keeps text that starts with = + - @ from becoming a formula

// Data rows of 'START links' (row 5 on) as objects; column K (Featured order) is read only if the sheet has it
function startRows_(sh) {
  var last = sh.getLastRow();
  if (last < 5) return [];
  var cols = Math.min(11, sh.getMaxColumns());
  var vals = sh.getRange(5, 1, last - 4, cols).getDisplayValues(), out = [];
  function num(v) { v = String(v == null ? '' : v).trim(); if (v === '') return ''; var n = Number(v); return isNaN(n) ? '' : n; }
  vals.forEach(function (r, i) {
    var title = String(r[1]).trim(), link = String(r[2]).trim();
    if (!title && !link) return;
    out.push({
      row: i + 5, theme: String(r[0]).trim() || 'General', title: title, link: link, desc: String(r[3]).trim(), descEn: String(r[9] || '').trim(),
      order: num(r[4]), feat: /^(x|y|yes|si|sí|true|1|★|✓)$/i.test(String(r[7]).trim()),
      featOrd: num(r[10]), icon: String(r[8] || '').trim()
    });
  });
  return out;
}

function startSortKey_(v) { return v === '' ? 1e9 : v; }
function startSortLinks_(a, b) { return (startSortKey_(a.order) - startSortKey_(b.order)) || (a.row - b.row); }
function startSortFeat_(a, b) { return (startSortKey_(a.featOrd) - startSortKey_(b.featOrd)) || startSortLinks_(a, b); }

// Tab 'START themes' (rows 1-3 notes, header row 4): A Theme, B Theme (English), C Icon, D Order. Created on the first theme edit.
function startThemesSheet_(create) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName('START themes');
  if (!sh && create) {
    sh = ss.insertSheet('START themes');
    sh.getRange(1, 1, 3, 1).setValues([
      ['Sections (themes) of the Start page: English name, icon and the position of each card. Changed from the app (Start tab > Edit).'],
      ['Theme must match column A of the tab START links exactly. A theme that has no row here is shown after the others, in the order of its first link.'],
      ['You can change it here too. Order is a number (1 = first card).']
    ]);
    sh.getRange(4, 1, 1, 4).setValues([['Theme', 'Theme (English)', 'Icon', 'Order']]).setFontWeight('bold');
    sh.setFrozenRows(4);
  }
  return sh;
}

// Themes in display order: [{name, nameEn, icon, order}]. Themes with an Order number first, then the others in the order of their first link.
function startThemeState_(rows) {
  var meta = {}, tsh = startThemesSheet_(false);
  if (tsh && tsh.getLastRow() >= 5) {
    tsh.getRange(5, 1, tsh.getLastRow() - 4, 4).getDisplayValues().forEach(function (r) {
      var k = String(r[0]).trim().toLowerCase();
      if (k && !meta[k]) { var o = String(r[3]).trim(); meta[k] = { nameEn: String(r[1]).trim(), icon: String(r[2]).trim(), order: o === '' || isNaN(Number(o)) ? '' : Number(o) }; }
    });
  }
  var list = [], seen = {};
  rows.forEach(function (r) {
    var k = r.theme.toLowerCase();
    if (seen[k]) return;
    var m = meta[k] || { nameEn: '', icon: '', order: '' };
    seen[k] = { name: r.theme, nameEn: m.nameEn, icon: m.icon, order: m.order, first: list.length };
    list.push(seen[k]);
  });
  list.sort(function (a, b) { return (startSortKey_(a.order) - startSortKey_(b.order)) || (a.first - b.first); });
  return list;
}

// Rewrites the whole 'START themes' table from the list (position = Order 1..n); themes without links are dropped
function startThemeWrite_(entries) {
  var sh = startThemesSheet_(true);
  var last = sh.getLastRow();
  if (last >= 5) sh.getRange(5, 1, last - 4, Math.max(4, sh.getMaxColumns())).clearContent();
  if (!entries.length) return;
  var data = entries.map(function (e, i) { return [startSafe_(e.name), startSafe_(e.nameEn || ''), startSafe_(e.icon || ''), i + 1]; });
  if (5 + data.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 5 + data.length - 1 - sh.getMaxRows());
  sh.getRange(5, 1, data.length, 4).setValues(data);
}

function startLinksSheet_() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('START links');
  if (!sh) throw new Error("The tab 'START links' is missing in the sheet.");
  if (sh.getMaxColumns() < 11) sh.insertColumnsAfter(sh.getMaxColumns(), 11 - sh.getMaxColumns());
  if (!String(sh.getRange(4, 11).getDisplayValue()).trim()) sh.getRange(4, 11).setValue('Featured order').setFontWeight('bold');
  return sh;
}

// Runs fn(sh, rows) under the script lock and returns the fresh links
function startWrite_(person, fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = startLinksSheet_();
    fn(sh, startRows_(sh));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getLinks(person);
}

// Adds a link (item.row = 0) or updates one (item.row = sheet row; origTitle/origLink must still match that row)
function saveLink(item) {
  var theme = String(item.theme || '').trim() || 'General';
  var title = String(item.title || '').trim();
  var link = String(item.link || '').trim();
  var desc = String(item.desc || '').trim();
  var descEn = String(item.descEn || '').trim();
  var icon = String(item.icon || '').trim();
  var feat = item.feat === true ? 'x' : '';
  var ordRaw = String(item.order == null ? '' : item.order).trim();
  if (!title || !/^https?:\/\/[^\s]+$/i.test(link)) throw new Error('Title and a link starting with http:// or https:// are required.');
  if (theme.length > 40 || title.length > 120 || link.length > 500 || desc.length > 400 || descEn.length > 400 || icon.length > 8) throw new Error('Text is too long.');
  var ord = '';
  if (ordRaw !== '') { ord = Number(ordRaw); if (isNaN(ord)) throw new Error('Order must be a number.'); }
  var safe = startSafe_;
  return startWrite_(item.person, function (sh, rows) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var row = Number(item.row) || 0;
    // a theme typed in another spelling (upper/lower case) joins the existing theme
    rows.forEach(function (r) { if (r.theme.toLowerCase() === theme.toLowerCase()) theme = r.theme; });
    if (row) {
      if (row < 5) throw new Error('Invalid row.');
      var cur = sh.getRange(row, 1, 1, 7).getDisplayValues()[0];
      if (String(cur[1]).trim() !== String(item.origTitle || '').trim() || String(cur[2]).trim() !== String(item.origLink || '').trim()) throw new Error('This link was changed in the meantime. Reload and try again.');
      sh.getRange(row, 1, 1, 5).setValues([[safe(theme), safe(title), link, safe(desc), ord]]);
      sh.getRange(row, 8, 1, 3).setValues([[feat, safe(icon), safe(descEn)]]);
    } else {
      var last = sh.getLastRow(), target = 5;
      if (last >= 5) {
        var bc = sh.getRange(5, 2, last - 4, 2).getDisplayValues();
        for (var i = bc.length - 1; i >= 0; i--) { if (String(bc[i][0]).trim() || String(bc[i][1]).trim()) { target = i + 6; break; } }
      }
      var who = '';
      try { who = Session.getActiveUser().getEmail() || ''; } catch (e) {}
      var day = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), target - sh.getMaxRows());
      sh.getRange(target, 1, 1, 7).setValues([[safe(theme), safe(title), link, safe(desc), ord, who, day]]);
      sh.getRange(target, 8, 1, 3).setValues([[feat, safe(icon), safe(descEn)]]);
    }
  });
}

// Deletes a link row (origTitle/origLink must still match that row)
function deleteLink(row, origTitle, origLink, person) {
  row = Number(row) || 0;
  if (row < 5) throw new Error('Invalid row.');
  return startWrite_(person, function (sh) {
    var cur = sh.getRange(row, 1, 1, 7).getDisplayValues()[0];
    if (String(cur[1]).trim() !== String(origTitle || '').trim() || String(cur[2]).trim() !== String(origLink || '').trim()) throw new Error('This link was changed in the meantime. Reload and try again.');
    sh.deleteRow(row);
  });
}

// Saves a new order of one list of links. items = [{row, title, link}] in the new order.
// featured = true: the list "Most important" (column K), else all links of one theme (column E). Positions are renumbered 1..n.
function reorderLinks(items, featured, person) {
  if (!items || !items.length || items.length > 500) throw new Error('Invalid list.');
  var changed = 'This link was changed in the meantime. Reload and try again.';
  return startWrite_(person, function (sh, rows) {
    var by = {}, seen = {}, picked = [], isF = featured === true || featured === 'true';
    rows.forEach(function (r) { by[r.row] = r; });
    items.forEach(function (it) {
      var r = by[Number(it.row)];
      if (!r || seen[r.row] || r.title !== String(it.title || '').trim() || r.link !== String(it.link || '').trim()) throw new Error(changed);
      seen[r.row] = true; picked.push(r);
    });
    var theme = picked[0].theme.toLowerCase();
    var cur = rows.filter(function (r) { return isF ? r.feat : r.theme.toLowerCase() === theme; });
    if (cur.length !== picked.length) throw new Error(changed);
    picked.forEach(function (r) { if (isF ? !r.feat : r.theme.toLowerCase() !== theme) throw new Error(changed); });
    picked.forEach(function (r, k) { sh.getRange(r.row, isF ? 11 : 5).setValue(k + 1); });
  });
}

// Renames a theme (all its links), sets its English name and icon. Renaming to the name of another theme merges them.
function renameTheme(oldName, name, nameEn, icon, person) {
  oldName = String(oldName || '').trim(); name = String(name || '').trim(); nameEn = String(nameEn || '').trim(); icon = String(icon || '').trim();
  if (!oldName || !name) throw new Error('The theme needs a name.');
  if (name.length > 40 || nameEn.length > 40 || icon.length > 8) throw new Error('Text is too long.');
  return startWrite_(person, function (sh, rows) {
    var entries = startThemeState_(rows), mine = null, other = null;
    entries.forEach(function (e) {
      if (e.name.toLowerCase() === oldName.toLowerCase()) mine = e;
      else if (e.name.toLowerCase() === name.toLowerCase()) other = e;
    });
    if (!mine) throw new Error('This theme was changed in the meantime. Reload and try again.');
    rows.forEach(function (r) { if (r.theme.toLowerCase() === oldName.toLowerCase() && r.theme !== (other ? other.name : name)) sh.getRange(r.row, 1).setValue(startSafe_(other ? other.name : name)); });
    if (other) { other.nameEn = nameEn; other.icon = icon; entries.splice(entries.indexOf(mine), 1); }
    else { mine.name = name; mine.nameEn = nameEn; mine.icon = icon; }
    startThemeWrite_(entries);
  });
}

// Deletes a theme with all its links
function deleteTheme(name, person) {
  name = String(name || '').trim();
  if (!name) throw new Error('Invalid theme.');
  return startWrite_(person, function (sh, rows) {
    var entries = startThemeState_(rows), keep = [], found = false;
    entries.forEach(function (e) { if (e.name.toLowerCase() === name.toLowerCase()) found = true; else keep.push(e); });
    if (!found) throw new Error('This theme was changed in the meantime. Reload and try again.');
    var del = rows.filter(function (r) { return r.theme.toLowerCase() === name.toLowerCase(); }).map(function (r) { return r.row; }).sort(function (a, b) { return b - a; });
    del.forEach(function (r) { sh.deleteRow(r); });
    startThemeWrite_(keep);
  });
}

// Saves a new order of the theme cards. names = all theme names in the new order.
function reorderThemes(names, person) {
  var changed = 'This theme was changed in the meantime. Reload and try again.';
  return startWrite_(person, function (sh, rows) {
    var entries = startThemeState_(rows), map = {}, seen = {}, out = [];
    names = names || [];
    if (names.length !== entries.length) throw new Error(changed);
    entries.forEach(function (e) { map[e.name.toLowerCase()] = e; });
    names.forEach(function (n) {
      var k = String(n || '').trim().toLowerCase(), e = map[k];
      if (!e || seen[k]) throw new Error(changed);
      seen[k] = true; out.push(e);
    });
    startThemeWrite_(out);
  });
}

// Saves (on = true) or removes (on = false) one favorite of one person; returns that person's favorites
function toggleFavorite(person, title, link, on) {
  person = String(person || '').trim();
  link = String(link || '').trim();
  title = String(title || '').trim();
  person = person.replace(/\s+/g, ' ');
  if (person.length < 2 || person.length > 40) throw new Error('Invalid name.');
  var canon = '';
  startPeople().forEach(function (n) { if (n.toLowerCase() === person.toLowerCase()) canon = n; }); // listed people keep their spelling
  if (!/^https?:\/\/[^\s]+$/i.test(link) || link.length > 500) throw new Error('Invalid link.');
  function safe(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName('START favorites');
    if (!sh) throw new Error("The tab 'START favorites' is missing in the sheet.");
    var last = sh.getLastRow();
    var rows = last >= 5 ? sh.getRange(5, 1, last - 4, 3).getDisplayValues() : [];
    if (!canon) rows.forEach(function (r) { if (!canon && String(r[0]).trim().toLowerCase() === person.toLowerCase()) canon = String(r[0]).trim(); }); // same typed name, same spelling as before
    if (!canon) canon = person;
    var hit = [];
    rows.forEach(function (r, i) { if (String(r[0]).trim().toLowerCase() === canon.toLowerCase() && String(r[2]).trim() === link) hit.push(i + 5); });
    if (on === true || on === 'true') {
      if (!hit.length) {
        var target = Math.max(5, last + 1);
        var day = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), target - sh.getMaxRows());
        sh.getRange(target, 1, 1, 4).setValues([[canon, safe(title.slice(0, 120)), link, day]]);
      }
    } else {
      for (var i = hit.length - 1; i >= 0; i--) sh.deleteRow(hit[i]);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return startFavs(canon);
}

// Priority 1-4 per plant #, read-only from the CG DATABASE sheet (tab 'PLANT LIST connected': column A = plant #, column P = URC prioridades 1-4). Plant # per item comes from the tab 'Incoming' (Plant # col Q, ID col V) of the Plug / Finished Planting sheet.
var PRIO_DB_ID = '1jO4v8F2rxtn0_PxFjbFCESLyYAUSxi6CjtxAr5zbNU8', PLUG_SHEET_ID = '1OK8tZaGvYjAKAPHNaqFfOqk0U9WfCxOTJscKKHY9LO4';
function getPrioMap_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('prioMap1');
  if (hit) return JSON.parse(hit);
  var sh = SpreadsheetApp.openById(PRIO_DB_ID).getSheetByName('PLANT LIST connected');
  var map = {};
  if (sh && sh.getLastRow() >= 6) {
    var n = sh.getLastRow() - 5, a = sh.getRange(6, 1, n, 1).getDisplayValues(), p = sh.getRange(6, 16, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      var k = String(a[i][0]).trim(), pr = Number(p[i][0]);
      if (k && pr >= 1 && pr <= 4 && pr === Math.floor(pr)) map[k] = pr;
    }
  }
  try { cache.put('prioMap1', JSON.stringify(map), 600); } catch (e) {}
  return map;
}
function getPrioByItemId_(ids) { // { itemId: {plant, prio} } for the ids asked
  var out = {}, map = getPrioMap_();
  var sh = SpreadsheetApp.openById(PLUG_SHEET_ID).getSheetByName('Incoming');
  if (!sh || sh.getLastRow() < 4) return out;
  var want = {}; ids.forEach(function (x) { if (x) want[x] = 1; });
  var v = sh.getRange(4, 17, sh.getLastRow() - 3, 6).getDisplayValues(); // Q..V
  for (var r = 0; r < v.length; r++) {
    var id = String(v[r][5]).trim();
    if (!want[id]) continue;
    var pl = String(v[r][0]).trim();
    out[id] = { plant: pl, prio: map[pl] || 0 };
  }
  return out;
}

// ---------- Work week tab: incoming only (one week at a time), data from the tab 'IMPORT work week' ----------
// That tab imports from the sheet 'Plug / Finished Planting': overview B11:I13 of 'Labor Summary (goal rates)', and the incoming items of the week from 'Incoming_arrived'.
function getWorkWeek() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('IMPORT work week');
  if (!sh) return { missing: true };
  var head = sh.getRange(4, 2, 2, 1).getValues();
  var week = head[0][0], target = head[1][0];
  if (week === '' || String(week).charAt(0) === '#') return { error: 'import' };
  var ov = sh.getRange(9, 1, 3, 8), ovV = ov.getValues(), ovD = ov.getDisplayValues();
  var overview = [];
  for (var i = 0; i < 3; i++) {
    var area = String(ovD[i][0]).trim();
    if (!area) continue;
    if (area.charAt(0) === '#') return { error: 'import' };
    var planned = parseFloat(String(ovD[i][4]).replace(/,/g, '')); // what the sheet shows: a dash or empty means nothing planned, so no bar
    var planted = parseFloat(String(ovD[i][3]).replace(/,/g, '')) || 0;
    var rem = ovV[i][2] === '' ? NaN : Number(ovV[i][2]);
    overview.push({
      key: /plug/i.test(area) ? 'plug' : (/finished/i.test(area) ? 'finished' : 'hb'),
      area: area, period: String(ovD[i][1]), remaining: String(ovD[i][2]), planted: String(ovD[i][3]), planned: String(ovD[i][4]),
      pct: String(ovD[i][5]), hours: String(ovD[i][6]), workers: String(ovD[i][7]),
      remainingNum: isNaN(rem) ? null : rem,
      left: (isNaN(planned) || planned <= 0) ? null : Math.max(0, Math.min(1, 1 - planted / planned)) // same rule as the status bar in the source sheet
    });
  }
  var items = [];
  var last = sh.getLastRow();
  if (last >= 14) {
    var rg = sh.getRange(14, 1, last - 13, 17), v = rg.getValues(), d = rg.getDisplayValues();
    for (var r = 0; r < v.length; r++) {
      var size = String(d[r][0]).trim(), genus = String(d[r][1]).trim(), variety = String(d[r][2]).trim();
      if (!genus && !size) continue;
      if (size.charAt(0) === '#') continue;
      var hasIn = /in/i.test(size), isHB = hasIn && /HB/i.test(variety); // same rules as the source sheet: size contains "in" = finished pots, variety contains HB = HBs
      items.push({
        size: size, genus: genus, variety: variety, goal: Number(v[r][3]) || 0, flats: String(d[r][4]), plug: String(d[r][5]).trim(),
        loc: String(d[r][6]).trim(), plantDate: String(d[r][7]), to: String(d[r][9]).trim(), vendor: String(d[r][10]).trim(),
        planted: v[r][13] === true || String(d[r][13]).toUpperCase() === 'TRUE', qtyPlanted: String(d[r][14]),
        area: isHB ? 'hb' : (hasIn ? 'finished' : 'plug'), id: String(d[r][12]).trim(), from: String(d[r][8]).trim(), plant: '', prio: 0
      });
    }
  }
  var prioOk = true;
  try {
    var pm = getPrioByItemId_(items.map(function (x) { return x.id; }));
    items.forEach(function (x) { var q = pm[x.id]; if (q) { x.plant = q.plant; x.prio = q.prio; } });
  } catch (e) { prioOk = false; } // the Work week still works without priorities
  return {
    prioOk: prioOk, missing: false, week: String(week), target: Number(target) || 0, overview: overview, items: items,
    readAt: { ms: Date.now(), txt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d h:mm a z') },
    sheetUrl: 'https://docs.google.com/spreadsheets/d/1CD6mirTQD70kw1Y1oR37-gbTuUoezTGtK5dwTH3XqVA/edit?gid=1205559600#gid=1205559600&range=A1'
  };
}


// ---------- Ask tab (chat-style data lookup, no AI yet) ----------
// Reads the CG Production Management sheet: 'Production Plan' (header row 4, data from row 5), 'URC Orders' (header row 2, data from row 3)
// and 'Incoming Plan 2' (header row 3, data from row 4). The question is split into month / week / size / source words; every other word must appear in the plant name.
// To switch AI on later, replace the body of askQuestion with a call that sends the question plus these sections to the model.
var ASK_INC_ID = '1OK8tZaGvYjAKAPHNaqFfOqk0U9WfCxOTJscKKHY9LO4'; // Plug / Finished Planting: tab Incoming (header row 3, current week only)
var ASK_URC_ID = '1uYWyLoK2ujBrVxoFKKTE7zGQbXDNY-FtxcIKNgj82zE'; // CG URCs: tab URC ORDERS (header row 3)
var ASK_PROD_ID = '1eF53WTexVdCWSgOZtCrVjser62Dj2I3HXd6LWP9_PQE'; // CG Production Plan: tab Production Plan (header row 3)
var ASK_FIN_ID = '1uMv0lhxSOXaXw9XhmPjnyl2s6W10SPe09yZ7ac0RTlg'; // CG Inventory App: tab GF_Planting (finished planting records)
var ASK_PLUG_ID = '1rIkxAMl_mrXiMA-xyUWBS4GJas_Sb3I1FWQcPzMVuS8'; // CG INVENTORY: tab GF_PlugPlanting (plug planting records)
function askQuestion(question) {
  var q = String(question || '').trim().slice(0, 300);
  var t0 = Date.now();
  var tz = 'America/Los_Angeles'; // the business runs on Pacific time, whatever the spreadsheet's own time zone is
  function norm(s) { return String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }
  function dkey(v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    var s = String(v).trim(); if (!s) return '';
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    var d = new Date(s); return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  function num(x) { var n = Number(String(x).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
  function isoWeek(k) { var p = k.split('-'), d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]))); var day = d.getUTCDay() || 7; d.setUTCDate(d.getUTCDate() + 4 - day); var y0 = new Date(Date.UTC(d.getUTCFullYear(), 0, 1)); return Math.ceil(((d - y0) / 86400000 + 1) / 7); }
  function addDays(k, n) { var p = k.split('-'); return Utilities.formatDate(new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + n)), 'UTC', 'yyyy-MM-dd'); }
  var MN = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  var MS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var n = norm(q), rest = ' ' + n.replace(/[^a-z0-9 ]/g, ' ') + ' ';
  var notes = [], filterTxt = [];

  // when: month (optionally with a year), this week or next week
  var from = '', to = '', yearTxt = (rest.match(/ (20\d\d) /) || [])[1] || '';
  var mo = -1;
  for (var i = 0; i < 12; i++) if (rest.indexOf(' ' + MN[i] + ' ') >= 0 || rest.indexOf(' ' + MS[i] + ' ') >= 0) { mo = i; rest = rest.replace(' ' + MN[i] + ' ', ' ').replace(' ' + MS[i] + ' ', ' '); break; }
  if (mo >= 0) {
    var cy = Number(todayKey.slice(0, 4)), cm = Number(todayKey.slice(5, 7)) - 1;
    var yy = yearTxt ? Number(yearTxt) : (mo >= cm ? cy : cy + 1);
    from = yy + '-' + ('0' + (mo + 1)).slice(-2) + '-01';
    to = Utilities.formatDate(new Date(Date.UTC(yy, mo + 1, 0)), 'UTC', 'yyyy-MM-dd');
    filterTxt.push(MN[mo].charAt(0).toUpperCase() + MN[mo].slice(1) + ' ' + yy);
    if (!yearTxt) notes.push('No year given: using ' + yy + '.');
  } else {
    var wk = /( this week | esta semana )/.test(rest) ? 0 : (/( next week | proxima semana | semana que viene )/.test(rest) ? 1 : -1);
    if (wk >= 0) {
      var dow = new Date(todayKey + 'T12:00:00Z').getUTCDay(), mon = addDays(todayKey, -((dow + 6) % 7) + 7 * wk);
      from = mon; to = addDays(mon, 6); filterTxt.push((wk ? 'Next week' : 'This week') + ' (' + from + ' to ' + to + ')');
    }
  }
  rest = rest.replace(/ (this week|next week|esta semana|proxima semana|semana que viene|20\d\d) /g, ' ');

  // size, e.g. 4in, 4 in, 4inHB
  var size = '';
  var sm = rest.match(/ (\d{1,2}) ?(in|inch|inches|pulgada|pulgadas)( hb)? /);
  if (sm) { size = sm[1] + 'in' + (sm[3] ? 'hb' : ''); rest = rest.replace(sm[0], ' '); filterTxt.push(size); }
  else { var sm2 = rest.match(/ (\d{1,2}in)(hb)? /); if (sm2) { size = sm2[1] + (sm2[2] || ''); rest = rest.replace(sm2[0], ' '); filterTxt.push(size); } }

  // which data: incoming (URC orders) and/or the production plan
  var wantIn = /( incoming | arriv\w* | urc | urcs | llega\w* | pedid\w* | order | orders | vendor | received | recib\w* )/.test(rest);
  var wantPl = /( plan | planned | plant | planted | planting | production | produccion | ready | listo | goal | meta | siembra | sembrar )/.test(rest);
  if (!wantIn && !wantPl) { wantIn = true; wantPl = true; }
  var stop = ' how many much is are of this that for in on at the a an what which check show me tell give list all any to from by with and or do we our there be will when was were have has it its i can you please total qty quantity quantities amount need needed expected planned plan plant planted planting production ready goal incoming arriving arrives arrive urc urcs order orders vendor como cuantos cuantas cuanto cuanta hay de la el los las para en que se esta estan por con y o un una del al mes plantas planta pedidos pedido llega llegan meta ';
  var words = rest.split(' ').filter(function (w) { return w && stop.indexOf(' ' + w + ' ') < 0 && !/^\d+$/.test(w); });
  if (words.length) filterTxt.unshift('"' + words.join(' ') + '"');
  if (!words.length && !from && !size) {
    return { question: q, ok: false, sections: [], notes: [], filter: '', ms: Date.now() - t0 };
  }
  function nameOk(h) { var hh = norm(h); for (var w = 0; w < words.length; w++) if (hh.indexOf(words[w]) < 0) return false; return true; }
  function inRange(k) { return !!k && (!from || (k >= from && k <= to)); }
  var CAP = 500, sections = [];
  // The big tabs are read once and kept for 15 minutes in the script cache (compact rows, in slices under 100 KB per key)
  var cache = CacheService.getScriptCache();
  function load(key, build) {
    var cnt = Number(cache.get(key + '_n'));
    if (cnt) {
      var keys = []; for (var c = 0; c < cnt; c++) keys.push(key + '_' + c);
      var got = cache.getAll(keys), txt = '', whole = true;
      keys.forEach(function (kk) { if (got[kk] == null) whole = false; else txt += got[kk]; });
      if (whole) { try { return JSON.parse(txt); } catch (e1) {} }
    }
    var data = build(), js = JSON.stringify(data), parts = {}, np = 0;
    for (var p = 0; p < js.length; p += 30000) { parts[key + '_' + np] = js.slice(p, p + 30000); np++; }
    parts[key + '_n'] = String(np);
    try { cache.putAll(parts, 900); } catch (e2) {}
    return data;
  }

  if (wantPl) {
    var plan = load('askPlan2', function () {
      var pp = SpreadsheetApp.openById(ASK_PROD_ID).getSheetByName('Production Plan'), out = [];
      pp.getRange(4, 1, Math.max(1, pp.getLastRow() - 3), 18).getValues().forEach(function (r) {
        var goal = num(r[5]); if (!(goal > 0)) return;
        out.push([String(r[0]), (String(r[1]) + ' ' + String(r[2])).trim(), String(r[9]).trim(), dkey(r[3]), dkey(r[4]), goal, String(r[10]).trim(), String(r[11]).trim(), String(r[16]).trim()]);
      });
      return out;
    });
    var hit = [];
    plan.forEach(function (r) {
      if (size && norm(r[0]).replace(/\s/g, '') !== size) return;
      if (words.length && !nameOk(r[1] + ' ' + r[2])) return;
      hit.push({ size: r[0], plant: r[1], asst: r[2], pk: r[3], rk: r[4], goal: r[5], fc: r[6], tc: r[7], vendor: r[8] });
    });
    notes.push('Production Plan lines with Goal 0 are not shown. Data is refreshed every 15 minutes.');
    var head = ['Plant date', 'Size', 'Plant', 'Goal', 'From', 'To', 'Vendor'];
    function prow(h) { return [h.pk, h.size, h.plant + (h.asst ? ' (' + h.asst + ')' : ''), h.goal, h.fc, h.tc, h.vendor]; }
    if (from) {
      var pl = hit.filter(function (h) { return inRange(h.pk); }).sort(function (a, b) { return a.pk < b.pk ? -1 : 1; });
      sections.push({ title: 'Production Plan: to plant in this period (Plant date)', head: head, rows: pl.slice(0, CAP).map(prow), more: Math.max(0, pl.length - CAP), total: pl.reduce(function (s, h) { return s + h.goal; }, 0), totalLabel: 'Goal', count: pl.length });
    } else {
      var byM = {};
      hit.forEach(function (h) { if (!h.pk) return; var m = h.pk.slice(0, 7); byM[m] = byM[m] || { g: 0, c: 0 }; byM[m].g += h.goal; byM[m].c++; });
      var ms = Object.keys(byM).filter(function (m) { return m >= todayKey.slice(0, 7); }).sort();
      sections.push({ title: 'Production Plan: Goal by plant month (from this month on)', head: ['Plant month', 'Goal', 'Lines'], rows: ms.slice(0, 24).map(function (m) { return [m, byM[m].g, byM[m].c]; }), more: Math.max(0, ms.length - 24), total: ms.reduce(function (s, m) { return s + byM[m].g; }, 0), totalLabel: 'Goal', count: ms.length });
    }
  }

  if (wantPl) { // records that were actually planted, with the ready date the sheet already calculated
    var rcut = addDays(todayKey, -14); // keep records whose ready date is not older than 2 weeks
    function gfSection(key, srcId, tab, rows0, cols, title, typ) {
      var data = load(key, function () {
        var sh = SpreadsheetApp.openById(srcId).getSheetByName(tab), out = [];
        if (!sh || sh.getLastRow() < 3) return out;
        sh.getRange(3, 1, sh.getLastRow() - 2, 26).getValues().forEach(function (r) {
          var nm = String(r[cols.name]).trim(); if (!nm) return;
          var rk = dkey(r[cols.ready]); if (!rk || rk < rcut) return;
          out.push([rk, dkey(r[cols.ts]), nm, num(r[cols.qty]), String(r[cols.size]).trim()]);
        });
        return out;
      });
      var gh = [];
      data.forEach(function (r) {
        if (from ? !inRange(r[0]) : r[0] < todayKey) return;
        if (size && typ === 'Finished' && norm(r[4]).replace(/\s/g, '') !== size) return;
        if (words.length && !nameOk(r[2])) return;
        gh.push(r);
      });
      gh.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
      sections.push({ title: title, head: ['Ready date', 'Planted', 'Plant', 'Qty', 'Size'], rows: gh.slice(0, CAP), more: Math.max(0, gh.length - CAP), total: gh.reduce(function (s, h) { return s + h[3]; }, 0), totalLabel: 'Qty', count: gh.length });
    }
    gfSection('askGfFin', ASK_FIN_ID, 'GF_Planting', 3, { name: 2, qty: 3, ts: 7, ready: 20, size: 22 }, 'Planted (Finished): getting ready ' + (from ? 'in this period' : 'from today on') + ' (ready date from GF_Planting)', 'Finished');
    gfSection('askGfPlug', ASK_PLUG_ID, 'GF_PlugPlanting', 3, { name: 1, qty: 2, ts: 3, ready: 17, size: 7 }, 'Planted (Plugs): getting ready ' + (from ? 'in this period' : 'from today on') + ' (ready date from GF_PlugPlanting)', 'Plugs');
    if (size) notes.push('The size filter applies to Finished planting records only (plug records show the tray size).');
  }

  if (wantIn) {
    var cutoff = addDays(todayKey, -365);
    var urc = load('askUrc2', function () {
      var uo = SpreadsheetApp.openById(ASK_URC_ID).getSheetByName('URC ORDERS'), out = [];
      uo.getRange(4, 1, Math.max(1, uo.getLastRow() - 3), 12).getValues().forEach(function (r) {
        if (!String(r[2]).trim()) return;
        var k = dkey(r[5]); if (!k || k < cutoff) return;
        out.push([k, String(r[0]), String(r[2]).trim(), num(r[4]), String(r[8]).trim() === '' ? '' : num(r[8]), String(r[9]).trim(), String(r[10]).trim(), String(r[11]).trim()]);
      });
      return out;
    });
    if (from && from < cutoff) notes.push('URC Orders older than ' + cutoff + ' are not included.');
    var uh = [];
    urc.forEach(function (r) {
      var k = r[0];
      if (from ? !inRange(k) : k < todayKey) return;
      if (words.length && !nameOk(r[2])) return;
      uh.push({ k: k, urc: r[1], name: r[2], qty: r[3], arr: r[4], from: r[5], to: r[6], vendor: r[7] });
    });
    uh.sort(function (a, b) { return a.k < b.k ? -1 : 1; });
    if (size) notes.push('URC Orders has no size column: the size filter is not applied there.');
    sections.push({ title: 'URC Orders: ' + (from ? 'ordered for this period' : 'ordered from today on'), head: ['Date', 'URC #', 'Name', 'Qty projected', 'Qty arrived', 'From', 'Vendor'], rows: uh.slice(0, CAP).map(function (h) { return [h.k, h.urc, h.name, h.qty, h.arr, h.from, h.vendor]; }), more: Math.max(0, uh.length - CAP), total: uh.reduce(function (s, h) { return s + h.qty; }, 0), totalLabel: 'Qty projected', count: uh.length });

    var ip = SpreadsheetApp.openById(ASK_INC_ID).getSheetByName('Incoming'); // only holds the current incoming week
    var iv = ip.getRange(4, 1, Math.max(1, ip.getLastRow() - 3), 21).getValues();
    var ih = [], isWeekQ = !!from && to === addDays(from, 6), fromWk = from ? isoWeek(from) : 0, listWk = {};
    iv.forEach(function (r) {
      var nm = (String(r[1]) + ' ' + String(r[2])).trim(); if (!nm) return;
      var wk = String(r[19]).trim(); listWk[wk] = 1;
      var pk = dkey(r[9]);
      if (from) { // a one-week question matches the list's Incoming Week (col T); a longer period matches the plant date
        if (isWeekQ) { if (wk !== String(fromWk)) return; } else if (!inRange(pk)) return;
      }
      if (size && norm(r[11]).replace(/\s/g, '') !== size) return;
      if (words.length && !nameOk(nm)) return;
      ih.push({ pk: pk, rk: dkey(r[20]), name: nm, qty: num(r[3]), goal: num(r[4]), to: String(r[11]).trim(), urc: String(r[18]), vendor: String(r[14]).trim() });
    });
    ih.sort(function (a, b) { return (a.pk || '9') < (b.pk || '9') ? -1 : 1; });
    notes.push('Incoming only holds the current incoming week (week ' + Object.keys(listWk).filter(Boolean).join(', ') + ')' + (isWeekQ ? '; this question is week ' + fromWk + '.' : '.'));
    sections.push({ title: 'Incoming: this week\'s incoming list (CTG Planting app)', head: ['Plant date', 'Ready date', 'Name', 'Qty projected', 'Goal', 'To', 'URC #', 'Vendor'], rows: ih.slice(0, CAP).map(function (h) { return [h.pk, h.rk, h.name, h.qty, h.goal, h.to, h.urc, h.vendor]; }), more: Math.max(0, ih.length - CAP), total: ih.reduce(function (s, h) { return s + h.goal; }, 0), totalLabel: 'Goal', count: ih.length });
  }
  return { question: q, ok: true, sections: sections, notes: notes, filter: filterTxt.join(' · '), ms: Date.now() - t0 };
}

// ---------- API for the GitHub Pages version of the app (Google sign-in) ----------
// The page at https://cascadetropicals.github.io/ctg-dash-app/ cannot use google.script.run, so it sends POST requests here (text/plain JSON).
// 1) action 'login': the page sends the Google ID token; it is checked with Google and answered with our own signed session (valid API_SESSION_DAYS days).
// 2) action 'rpc': the page sends the session plus the name of one of the functions below and its arguments; the function runs and its result is returned.
// Any Google account can sign in; each person then picks their own name in the app, as before. The old script.google.com address keeps working through doGet.
var GOOGLE_CLIENT_ID = '796335216653-im30hq9jmqe028l7fsjapeqnu5ptrqt6.apps.googleusercontent.com';
var API_SESSION_DAYS = 7;

function apiFunctions_() {
  return {
    getInit: getInit, getTasks: getTasks, saveScores: saveScores, getOverview: getOverview, getPerformance: getPerformance, getToday: getToday,
    getAccuracy: getAccuracy, getTimeline: getTimeline, getRawEntries: getRawEntries, getScoreTable: getScoreTable, startPeople: startPeople,
    startFavs: startFavs, getLinks: getLinks, saveLink: saveLink, deleteLink: deleteLink, reorderLinks: reorderLinks, renameTheme: renameTheme,
    deleteTheme: deleteTheme, reorderThemes: reorderThemes, toggleFavorite: toggleFavorite, getWorkWeek: getWorkWeek, askQuestion: askQuestion
  };
}

function apiSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('API_SESSION_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('API_SESSION_SECRET', s); }
  return s;
}

function apiSign_(text) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, apiSecret_())); }

// Checks the Google ID token with Google; returns { email, name } or throws.
function apiVerifyGoogle_(idToken) {
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(String(idToken || '')), { muteHttpExceptions: true });
  var j = {};
  try { j = JSON.parse(r.getContentText()); } catch (e) {}
  if (r.getResponseCode() !== 200 || j.aud !== GOOGLE_CLIENT_ID || String(j.email_verified) !== 'true' || !j.email || Number(j.exp) * 1000 < Date.now())
    throw new Error('GOOGLE_SIGNIN_FAILED');
  return { email: String(j.email).toLowerCase(), name: String(j.name || '') };
}

function apiMakeSession_(email) {
  var body = Utilities.base64EncodeWebSafe(JSON.stringify({ e: email, x: Date.now() + API_SESSION_DAYS * 86400000 }));
  return body + '.' + apiSign_(body);
}

// Returns the signed-in email or throws SESSION_EXPIRED.
function apiCheckSession_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 2 || apiSign_(parts[0]) !== parts[1]) throw new Error('SESSION_EXPIRED');
  var s = {};
  try { s = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); } catch (e) {}
  if (!s.e || Number(s.x) < Date.now()) throw new Error('SESSION_EXPIRED');
  return s.e;
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'login') {
      var u = apiVerifyGoogle_(req.idToken);
      out = { ok: true, session: apiMakeSession_(u.email), email: u.email, name: u.name };
    } else if (req.action === 'rpc') {
      apiCheckSession_(req.session);
      var fn = apiFunctions_()[String(req.fn)];
      if (typeof fn !== 'function') throw new Error('UNKNOWN_FUNCTION');
      var res = fn.apply(null, req.args || []);
      out = { ok: true, result: res === undefined ? null : res };
    } else {
      throw new Error('UNKNOWN_ACTION');
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
