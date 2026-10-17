/**
 * Free Time Ideas – save taps from the phone page into the "Ideas" sheet.
 * Deploy: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */
var SHEET_ID = '11UVa9eHWh-wRbGfe7uc-xjv5cZ0Hp3h51t48NFjikec';
var TAB = 'Ideas';
var TOKEN = 'fti-DbziM3rXemYo8RaX1cxfh5cC'; // must match TOKEN in index.html
var ALLOWED = {
  'Status':   ['Want to try', 'Booked', 'Done'],
  'Reaction': ['', 'Like', 'Not for us'],
  'Hidden':   ['', 'TRUE']
};

function doGet() {
  return json_({ ok: true, service: 'free-time-ideas', time: new Date().toISOString() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (p.token !== TOKEN) return json_({ ok: false, error: 'not allowed' });

    var field = String(p.field || '');
    if (!ALLOWED.hasOwnProperty(field)) return json_({ ok: false, error: 'unknown field' });
    var value = p.value == null ? '' : String(p.value);
    if (ALLOWED[field].indexOf(value) < 0) return json_({ ok: false, error: 'bad value' });

    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
    var lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    var ideaCol = header.indexOf('Idea') + 1, fieldCol = header.indexOf(field) + 1;
    if (!ideaCol || !fieldCol) return json_({ ok: false, error: 'column "' + field + '" not found' });

    var name = String(p.idea || '').trim();
    if (!name) return json_({ ok: false, error: 'missing idea' });
    var row = Number(p.row) || 0;
    if (!(row >= 2 && row <= lastRow && String(sh.getRange(row, ideaCol).getValue()).trim() === name)) {
      var names = sh.getRange(2, ideaCol, Math.max(lastRow - 1, 1), 1).getValues()
        .map(function (r) { return String(r[0]).trim(); });
      var i = names.indexOf(name);
      if (i < 0) return json_({ ok: false, error: 'idea not found' });
      row = i + 2;
    }

    var cell = sh.getRange(row, fieldCol);
    if (field === 'Hidden') cell.setValue(value === 'TRUE' ? true : '');
    else cell.setValue(value);
    return json_({ ok: true, row: row, field: field, value: value });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
