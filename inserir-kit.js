const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("ORCAMENTO");

// Retirar fórmulas
function freezeValues() {
  const range = sheet.getRange("B16:D1000");

  range.copyTo(range, { contentsOnly: true });
}

// Inserir item e kits
function standardKit(interval, query) {
  let row = 2;
  while (sheet.getRange(row, 2).isBlank() == false) {
    row += 1;
  }

  const rangeHeader = sheet.getRange(interval);
  const rangeQuery = sheet.getRange(query);

  const pasteHeader = sheet.getRange(`B${row}`);
  const pasteQuery = sheet.getRange(`B${row + 1}`);

  rangeHeader.copyTo(pasteHeader, { contentsOnly: false });
  rangeQuery.copyTo(pasteQuery, { contentsOnly: false });

  // Inserir item
  if (query == "B3") {
    const sheetTB = ss.getSheetByName("TB_ITENS");

    let newRow = 2;
    while (sheetTB.getRange(newRow, 2).isBlank() == false) {
      newRow += 1;
    }

    const itemOrigin = [[`=IFERROR(TB_ITENS!$B$${newRow - 1} + 1; 1)`]];
    const produtoOrigin = [[`=ORCAMENTO!C${row + 3}`]];
    const linhaOrigin = [[`=ORCAMENTO!F${row + 3}`]];
    const qtdOrigin = [[`=ORCAMENTO!C${row + 4}`]];
    const tagOrigin = [[`=ORCAMENTO!C${row + 5}`]];
    const docOrigin = [[`=ORCAMENTO!C${row + 6}`]];
    const valorOrigin = [[`=ORCAMENTO!J${row + 6}*D${newRow}`]];
    const noICMSOrigin = [[`=ORCAMENTO!K${row + 6}*D${newRow}`]];
    const formulaIPIOrigin = [[`=(I${newRow}*H${newRow}%)+I${newRow}`]];

    const itemDestiny = sheetTB.getRange(`B${newRow}`);
    const produtoDestiny = sheetTB.getRange(`C${newRow}`);
    const qtdDestiny = sheetTB.getRange(`D${newRow}`);
    const tagDestiny = sheetTB.getRange(`E${newRow}`);
    const docDestiny = sheetTB.getRange(`F${newRow}`);
    const valorDestiny = sheetTB.getRange(`I${newRow}`);
    const noICMSDestiny = sheetTB.getRange(`J${newRow}`);
    const formulaIPIDestiny = sheetTB.getRange(`K${newRow}`);
    const linhaDestiny = sheetTB.getRange(`L${newRow}`);

    itemDestiny.setValues(itemOrigin);
    produtoDestiny.setValues(produtoOrigin);
    qtdDestiny.setValues(qtdOrigin);
    tagDestiny.setValues(tagOrigin);
    docDestiny.setValues(docOrigin);
    valorDestiny.setValues(valorOrigin);
    noICMSDestiny.setValues(noICMSOrigin);
    formulaIPIDestiny.setValues(formulaIPIOrigin);
    linhaDestiny.setValues(linhaOrigin);

    const totalSemIPIOrigin = [[`=SUM($I$3:$I$${newRow})`]];
    const totalSemICMSOrigin = [[`=SUM($J$3:$J$${newRow})`]];
    const totalComIPIOrigin = [[`=SUM($K$3:$K$${newRow})`]];
    const totalLinhaOrigin = [[`=UNIQUE($L$3:$L$${newRow})`]];

    const totalSemIPIDestiny = sheetTB.getRange("N4");
    const totalSemICMSDestiny = sheetTB.getRange("O4");
    const totalComIPIDestiny = sheetTB.getRange("Q4");
    const totalLinhaDestiny = sheetTB.getRange("M3");

    totalSemIPIDestiny.setValues(totalSemIPIOrigin);
    totalSemICMSDestiny.setValues(totalSemICMSOrigin);
    totalComIPIDestiny.setValues(totalComIPIOrigin);
    totalLinhaDestiny.setValues(totalLinhaOrigin);

    const oldItemOrigin = sheet.getRange(`C${row + 2}`);

    itemDestiny.copyTo(oldItemOrigin);

    let cont = row + 9;
    let sumQTD = [[`=IF(D${cont} <> ""; D${cont}*$C$${row + 4}; "")`]];
    let qtdTotal = sheet.getRange(`L${cont}`);

    let itemNum = [[`=C${row + 2}`]];
    let itemCell = sheet.getRange(`M${cont}`);

    qtdTotal.setValues(sumQTD);
    itemCell.setValues(itemNum);

    while (cont < 1000) {
      cont += 1;
      sumQTD = [[`=IF(D${cont} <> ""; D${cont}*$C$${row + 4}; "")`]];
      qtdTotal = sheet.getRange(`L${cont}`);

      itemNum = [[`=$C$${row + 2}`]];
      itemCell = sheet.getRange(`M${cont}`);

      qtdTotal.setValues(sumQTD);
      itemCell.setValues(itemNum);
    }
  }
}

function insertItem() {
  standardKit("B2:M9", "B3");
}

function insertKitBT() {
  standardKit("B10:M10", "B11");
}

function insertKitMT() {
  standardKit("B12:M12", "B13");
}

function insertKitOT() {
  standardKit("B14:M14", "B15");
}

// Somar itens da tabela
function sumTB_ITENS() {
  const sumSheet = ss.getSheetByName("TB_ITENS");

  let row = 2;
  while (sumSheet.getRange(row, 2).isBlank() == false) {
    row += 1;
  }

  const originO2 = sumSheet.getRange("O2");
  const originO3 = sumSheet.getRange("O3");
  const originP2 = sumSheet.getRange("P2");
  const originP3 = sumSheet.getRange("P3");
  const originQ2 = sumSheet.getRange("Q2");
  const originQ3 = sumSheet.getRange("Q3");

  const destinyO2 = sumSheet.getRange(`K${row + 1}`);
  const destinyO3 = sumSheet.getRange(`L${row + 1}`);
  const destinyP2 = sumSheet.getRange(`K${row + 2}`);
  const destinyP3 = sumSheet.getRange(`L${row + 2}`);
  const destinyQ2 = sumSheet.getRange(`K${row + 3}`);
  const destinyQ3 = sumSheet.getRange(`L${row + 3}`);

  originO2.copyTo(destinyO2, { contentsOnly: false });
  originO3.copyTo(destinyO3, { contentsOnly: false });
  originP2.copyTo(destinyP2, { contentsOnly: false });
  originP3.copyTo(destinyP3, { contentsOnly: false });
  originQ2.copyTo(destinyQ2, { contentsOnly: false });
  originQ3.copyTo(destinyQ3, { contentsOnly: false });
}

// Totalizar itens do orçamento
function totalization() {
  const sheetTot = ss.getSheetByName("LM-TOTALIZACAO");

  const colC = sheet.getRange(`C18:C1000`);
  const colL = sheet.getRange(`L18:L1000`);

  const colA = sheetTot.getRange("A2:A1000");
  const colB = sheetTot.getRange("B2:B1000");

  colA.clear();
  colB.clear();

  colC.copyTo(colA, { contentsOnly: true });
  colL.copyTo(colB, { contentsOnly: true });
}

// Gravar modificações
function modifications() {
  const sheetMod = ss.getSheetByName("MODIFICACOES");
  const allValues = sheet.getRange("B18:M1000");
  const mod = sheetMod.getRange("B18");

  const initalValues = sheetMod.getRange("A:M");
  initalValues.clear();

  allValues.copyTo(mod, { contentsOnly: true });
}
