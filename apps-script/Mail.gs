// ---------- Friday reminder mail ----------
// Run this from a weekly time-driven trigger set for Friday noon Pacific (Triggers > Add trigger > sendManagerReminders > Time-driven > Week timer).
// The trigger decides the day and time; the function sends whenever it runs. Run by hand from the editor it sends too.
// Settings and manager emails are read from the MAIL tab.
// A manager gets a mail only if they still have blank scores for the current week (a 0 counts as scored; bonus rows are ignored).
function sendManagerReminders(e) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mail = ss.getSheetByName('MAIL');
  if (!mail) throw new Error('MAIL tab not found');
  var mv = mail.getDataRange().getValues();

  // Settings (label in column A, value in column B) and the manager table (header row starts with "Manager")
  var settings = {};
  var managers = [];
  var inTable = false;
  var lastRunRow = 0;
  for (var i = 0; i < mv.length; i++) {
    var a = String(mv[i][0]).trim();
    if (!a) continue;
    if (a.toLowerCase() === 'manager') { inTable = true; continue; }
    if (inTable) {
      managers.push({ name: a, email: String(mv[i][1]).trim(), lang: String(mv[i][2]).trim().toLowerCase() === 'en' ? 'en' : 'es' });
    } else {
      settings[a.toLowerCase()] = mv[i][1];
      if (a.toLowerCase() === 'last run') lastRunRow = i + 1;
    }
  }
  var sendOn = String(settings['send mails (yes/no)'] || '').trim().toLowerCase();
  if (sendOn !== 'yes' && sendOn !== 'si' && sendOn !== 'sí') { Logger.log('Mails are switched off on the MAIL tab.'); return; }
  var appLink = String(settings['app link'] || '').trim();
  if (!appLink) throw new Error('App link is empty on the MAIL tab');
  var testTo = String(settings['test email (leave empty to send to the managers)'] || '').trim();

  // Current week by label on DASH (outside weeks 1-10 the program is not running: no mails)
  var dash = ss.getSheetByName(SHEET_NAME);
  var top = dash.getRange(1, 1, 30, 2).getValues();
  var week = 0;
  for (var t = 0; t < top.length; t++) {
    var label = String(top[t][0]).trim().toLowerCase();
    if (label === 'semana actual' || label === 'current week') week = Number(top[t][1]) || 0;
  }
  if (week < 1 || week > 10) { Logger.log('Program is not running this week (week ' + week + ').'); return; }

  var sent = [];
  managers.forEach(function (m) {
    if (!m.email && !testTo) { Logger.log('No email for ' + m.name); return; }
    var groups = getTasks(m.name, week, m.lang);
    var lines = [];
    var count = 0;
    groups.forEach(function (g) {
      var blanks = g.tasks.filter(function (tk) { return !tk.isBonus && tk.value === ''; }).map(function (tk) { return tk.task; });
      if (blanks.length) { count += blanks.length; lines.push({ person: g.name, tasks: blanks }); }
    });
    if (!count) return;

    var url = appLink + (appLink.indexOf('?') < 0 ? '?' : '&') + 'manager=' + encodeURIComponent(m.name) + '&lang=' + m.lang;
    var es = m.lang === 'es';
    var subject = (testTo ? '[TEST ' + m.name + '] ' : '') + (es ? 'Millas: ' + (count === 1 ? 'falta 1 calificación' : 'faltan ' + count + ' calificaciones') + ' de la semana ' + week : 'Millas: ' + count + (count === 1 ? ' score is' : ' scores are') + ' still missing for week ' + week);
    var intro = es
      ? 'Hola ' + m.name + ', ya pasó el viernes al mediodía (hora del Pacífico) y todavía faltan calificaciones de la semana ' + week + ':'
      : 'Hi ' + m.name + ', it is past Friday noon (Pacific time) and some scores for week ' + week + ' are still missing:';
    var button = es ? 'Calificar ahora' : 'Score now';
    var html = '<div style="font-family:Arial,sans-serif;font-size:15px;color:#17241b">' +
      '<p>' + intro + '</p><ul>' +
      lines.map(function (l) { return '<li><b>' + l.person + '</b>: ' + l.tasks.join(', ') + '</li>'; }).join('') +
      '</ul><p><a href="' + url + '" style="display:inline-block;background:#2f6b3c;color:#fff;text-decoration:none;padding:12px 20px;border-radius:8px;font-weight:bold">' + button + '</a></p>' +
      '<p style="color:#6b776e;font-size:13px">' + url + '</p></div>';
    var text = intro + '\n' + lines.map(function (l) { return '- ' + l.person + ': ' + l.tasks.join(', '); }).join('\n') + '\n\n' + button + ': ' + url;
    MailApp.sendEmail({ to: testTo || m.email, subject: subject, body: text, htmlBody: html, name: 'CTG App' });
    sent.push(m.name + ' (' + count + ')');
  });

  Logger.log('Week ' + week + ': ' + (sent.length ? 'mails sent to ' + sent.join(', ') : 'nobody has missing scores'));
  if (lastRunRow) mail.getRange(lastRunRow, 2).setValue(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd HH:mm') + ' Pacific, week ' + week + ': ' + (sent.length ? sent.join(', ') : 'nobody missing'));
}
