// ---------- Paste each employee's DASH table into their own sheet
// Reads the SHEETS tab (Employee | Sheet link | Tab name | Status | Last pasted), finds that person's block on the DASH tab,
// and overwrites the tab named in column C (default "Millas") in their sheet. Creates the tab if it is missing.
// DASH columns E and F ("Quien revisa", "Cuando") are left out of the paste.
// Run it from the editor (function dropdown > pasteDashToEmployeeSheets > Run). A weekly time-driven trigger (Sunday night Pacific) runs it automatically.
function pasteDashToEmployeeSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dash = ss.getSheetByName('DASH');
  var list = ss.getSheetByName('SHEETS');
  if (!dash || !list) throw new Error('Tabs DASH and SHEETS are required');

  // 1) Find every employee block on DASH: header row has the name in A and "Millas" in B, block ends at the "Total:" row
  var all = dash.getDataRange().getDisplayValues();
  var blocks = {};
  for (var i = 0; i < all.length; i++) {
    if (String(all[i][0]).trim() === '' || !/^millas$/i.test(String(all[i][1]).trim())) continue;
    var width = 0;
    for (var c = 0; c < all[i].length; c++) if (String(all[i][c]).trim() !== '') width = c + 1;
    var end = i;
    for (var j = i + 1; j < all.length; j++) {
      end = j;
      if (/^total:?$/i.test(String(all[j][5]).trim())) break;
      if (String(all[j][0]).trim() === '' && String(all[j][1]).trim() === '' && String(all[j][4]).trim() === '') { end = j - 1; break; }
    }
    blocks[String(all[i][0]).trim().toLowerCase()] = { row: i + 1, rows: all.slice(i, end + 1).map(function (r) { return r.slice(0, width); }) };
  }

  // 2) Find the header row on the SHEETS tab
  var rows = list.getDataRange().getDisplayValues();
  var head = -1;
  for (var h = 0; h < rows.length; h++) if (String(rows[h][0]).trim().toLowerCase() === 'employee') { head = h; break; }
  if (head < 0) throw new Error('Header "Employee" not found on the SHEETS tab');

  var stamp = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'MM/dd/yyyy h:mm a') + ' Pacific';

  // 3) Paste for each employee
  for (var r = head + 1; r < rows.length; r++) {
    var name = String(rows[r][0]).trim();
    if (name === '') continue;
    var url = String(rows[r][1]).trim();
    var tabName = String(rows[r][2]).trim() || 'Millas';
    var status = '', when = '';
    try {
      var block = blocks[name.toLowerCase()];
      var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (url === '') { status = 'No sheet link yet'; }
      else if (!block) { status = 'Name not found on DASH'; }
      else if (!m) { status = 'Link not valid'; }
      else {
        var target = SpreadsheetApp.openById(m[1]);
        var sh = target.getSheetByName(tabName) || target.insertSheet(tabName);
        var wOrig = block.rows[0].length;
        // Columns E and F on DASH ("Quien revisa", "Cuando") are not pasted. The "Total:" label sits in column F, so it moves to column A.
        var rowsOut = block.rows.map(function (r) {
          var keep = r.filter(function (v, i) { return i !== 4 && i !== 5; });
          if (/^total:?$/i.test(String(r[5]).trim()) && String(keep[0]).trim() === '') keep[0] = 'Total:';
          return keep;
        });
        var w = rowsOut[0].length;
        var notes = [
          ['Automatically updated every Sunday night'],
          ['']
        ].map(function (n) { var x = [n[0]]; while (x.length < w) x.push(''); return x; });
        var out = notes.concat(rowsOut);
        sh.clear();
        sh.getRange(1, 1, out.length, w).setValues(out);
        sh.getRange(1, 1).setFontWeight('bold');
        // Copy the look of the table from DASH: fonts, colors, borders, number formats, wrapping, column widths
        // (copyTo does not work between two spreadsheets, so copy the DASH sheet in as a temporary tab, take the format from it, then delete it)
        var tmp = dash.copyTo(target);
        try {
          tmp.getRange(block.row, 1, block.rows.length, Math.min(4, wOrig)).copyTo(sh.getRange(notes.length + 1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          if (wOrig > 6) tmp.getRange(block.row, 7, block.rows.length, wOrig - 6).copyTo(sh.getRange(notes.length + 1, 5), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        } finally {
          target.deleteSheet(tmp);
        }
        for (var col = 1; col <= w; col++) sh.setColumnWidth(col, dash.getColumnWidth(col <= 4 ? col : col + 2));
        status = 'OK - ' + block.rows.length + ' rows pasted';
        when = stamp;
      }
    } catch (err) {
      status = 'Error: ' + (err && err.message ? err.message : err);
    }
    list.getRange(r + 1, 4, 1, 2).setValues([[status, when]]);
  }
}
