/**
 * This Time Nepal — shared checklist backend.
 *
 * Deploy this bound to a Google Sheet with a tab named "Checklist"
 * whose header row is: item_id | text | checked | checked_by | checked_at
 *
 * Deploy > New deployment > type "Web app" > Execute as: Me,
 * Who has access: Anyone. Paste the resulting URL into SHEET_SYNC_URL
 * in index.html.
 */

var SHEET_NAME = 'Checklist';

function doGet(e) {
  return respond_(readAll_());
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var id = body.item_id;
  var checked = !!body.checked;
  var by = body.checked_by || '';

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 3).setValue(checked);
      sheet.getRange(i + 1, 4).setValue(checked ? by : '');
      sheet.getRange(i + 1, 5).setValue(checked ? new Date() : '');
      break;
    }
  }
  return respond_(readAll_());
}

function readAll_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    items.push({
      item_id: data[i][0],
      text: data[i][1],
      checked: data[i][2] === true,
      checked_by: data[i][3],
      checked_at: data[i][4] ? String(data[i][4]) : ''
    });
  }
  return items;
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
