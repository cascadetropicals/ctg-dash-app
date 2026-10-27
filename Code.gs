// CTG Dash App – manager scoring app
// Reads tasks from the DASH tab (Spanish layout) and writes scores straight into DASH (week columns G:P).

var SPREADSHEET_ID = '1ZCSHtPryT8Y13g7PL3pa-k62X2OC5j13YkYMh0R5maA';
var SHEET_NAME = 'DASH';
var MANAGERS = ['Dago', 'Jessica', 'Andrej'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CTG Dash App')
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
      var d = new Date(start.getFullYear(), start.getMonth(), start.getDate() + (w - 1) * 7);
      weekDate = Utilities.formatDate(d, tz, 'M/d');
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
    people: people
  };

  if (start instanceof Date) {
    var ws = new Date(start.getFullYear(), start.getMonth(), start.getDate() + (weekForData - 1) * 7);
    var we = new Date(ws.getFullYear(), ws.getMonth(), ws.getDate() + 6);
    var pe = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 69);
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

  // Order seconds: header row 4 (people from column C), weeks in rows 5-16 (column B = week start)
  var os = ss.getSheetByName('IMPORT order seconds').getRange(4, 1, 14, 30).getValues();
  var hdr = os[0];
  var weekRows = [];
  for (var r = 1; r < os.length; r++) {
    if (os[r][1] instanceof Date) weekRows.push(os[r]);
  }
  weekRows.sort(function (a, b) { return b[1] - a[1]; }); // newest week first (left side)

  var weeks = weekRows.map(function (row) {
    var d = row[1];
    return { key: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), n: String(row[0]).trim(), label: (d.getMonth() + 1) + '/' + d.getDate() };
  });

  var order = [];
  for (var c = 2; c < hdr.length; c++) {
    var nm = String(hdr[c]).trim();
    if (!nm) continue;
    var vals = weekRows.map(function (row) {
      var v = row[c];
      return v === '' || v === null || isNaN(Number(v)) ? null : Number(v);
    });
    if (vals.some(function (v) { return v !== null; })) order.push({ name: nm, vals: vals });
  }
  order.sort(function (a, b) { return a.name.localeCompare(b.name); });

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
  var lb = ss.getSheetByName('IMPORT labels planting');
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

  return {
    goal: goal,
    weeks: weeks.map(function (w) { return { n: w.n, label: w.label }; }),
    order: order,
    planting: planting,
    plugs: plugs,
    labels: labels,
    goals: { planting: goalsFinished, plugs: goalsPlugs, labels: goalsLabels }
  };
}
