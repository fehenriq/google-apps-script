function freezeValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ORCAMENTO");
  const range = sheet.getRange("B27:D1000");

  range.copyTo(range, { contentsOnly: true });
}

function standardKit(interval, query) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ORCAMENTO");

  let row = 15;
  while (sheet.getRange(row, 2).isBlank() == false) {
    row += 1;
  }

  const rangeHeader = sheet.getRange(interval);
  const rangeQuery = sheet.getRange(query);

  const pasteHeader = sheet.getRange(`B${row}`);
  const pasteQuery = sheet.getRange(`B${row + 1}`);

  rangeHeader.copyTo(pasteHeader, { contentsOnly: false });
  rangeQuery.copyTo(pasteQuery, { contentsOnly: false });
}

function insertKit() {
  standardKit("B13:L20", "B14");
}

function insertKitBT() {
  standardKit("B21:L21", "B22");
}

function insertKitMT() {
  standardKit("B23:L23", "B24");
}

function insertKitOT() {
  standardKit("B25:L25", "B26");
}
