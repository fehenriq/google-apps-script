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

  if (query == "B14") {
    let newRow = 3;
    while (sheet.getRange(newRow, 3).isBlank() == false) {
      newRow += 1;
    }

    const itemOrigin = [[`=C${row + 2}`]];
    const produtoOrigin = [[`=C${row + 3}`]];
    const qtdOrigin = [[`=C${row + 4}`]];
    const tagOrigin = [[`=C${row + 5}`]];
    const docOrigin = [[`=C${row + 6}`]];
    const valorOrigin = [[`=J${row + 6}`]];
    const noICMSOrigin = [[`=K${row + 6}`]];

    const itemDestiny = sheet.getRange(`B${newRow}`);
    const produtoDestiny = sheet.getRange(`C${newRow}`);
    const qtdDestiny = sheet.getRange(`D${newRow}`);
    const tagDestiny = sheet.getRange(`E${newRow}`);
    const docDestiny = sheet.getRange(`F${newRow}`);
    const valorDestiny = sheet.getRange(`I${newRow}`);
    const noICMSDestiny = sheet.getRange(`J${newRow}`);

    itemDestiny.setValues(itemOrigin);
    produtoDestiny.setValues(produtoOrigin);
    qtdDestiny.setValues(qtdOrigin);
    tagDestiny.setValues(tagOrigin);
    docDestiny.setValues(docOrigin);
    valorDestiny.setValues(valorOrigin);
    noICMSDestiny.setValues(noICMSOrigin);
  }
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
