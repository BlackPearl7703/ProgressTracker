const SHEET_NAME = "Daily log";

function doGet() {
  const sheet = getSheet();
  removeConnectionTestRows(sheet);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return json({ entries: [] });

  const headers = values.shift();
  const entries = values.map((row) =>
    Object.fromEntries(headers.map((header, index) => [header, row[index]])),
  );
  return json({ entries });
}

function doPost(event) {
  const payload = JSON.parse(event.postData.contents);
  const entry = payload.entry;
  if (entry.id === "connection-test-20260822") return json({ ok: true });
  const sheet = getSheet();
  removeConnectionTestRows(sheet);
  ensureHeaders(sheet);
  sheet.appendRow([
    entry.id,
    entry.student,
    entry.date,
    entry.learned,
    entry.practiced,
    entry.questions,
    entry.remarks,
    entry.cuetPaperYear,
    entry.cuetPaperScore,
    entry.clatPaperYear,
    entry.clatPaperScore,
    entry.duration,
    entry.confidence,
    entry.nextFocus,
  ]);
  return json({ ok: true });
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  ensureHeaders(sheet);
  return sheet;
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow([
    "id",
    "student",
    "date",
    "learned",
    "practiced",
    "questions",
    "remarks",
    "cuetPaperYear",
    "cuetPaperScore",
    "clatPaperYear",
    "clatPaperScore",
    "duration",
    "confidence",
    "nextFocus",
  ]);
}

function removeConnectionTestRows(sheet) {
  for (let row = sheet.getLastRow(); row >= 2; row--) {
    if (sheet.getRange(row, 1).getValue() === "connection-test-20260822") {
      sheet.deleteRow(row);
    }
  }
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
