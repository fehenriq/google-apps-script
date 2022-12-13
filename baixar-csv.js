function convertToCsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DESCRIÇÃO");
  const name = sheet.getRange("B2").getValues();
  const lastRowPrata = sheet.getRange("A4").getValues();
  const lastRowAmarela = sheet.getRange("A2").getValues();
  const valuesPrata = sheet.getRange(`Z3:AF${lastRowPrata}`).getValues();
  const valuesAmarela = sheet.getRange(`D3:D${lastRowAmarela}`).getValues();
  let contentPrata = "";
  let contentAmarela = "";

  for (let i = 0; i < valuesPrata.length; i++) {
    let row = "";
    for (let j = 0; j < valuesPrata[i].length; j++) {
      if (valuesPrata[i][j]) {
        row += valuesPrata[i][j];
      }
      row += ";";
    }
    contentPrata += row + "\n";
  }

  for (let i = 0; i < valuesAmarela.length; i++) {
    let row = "";
    for (let j = 0; j < valuesAmarela[i].length; j++) {
      if (valuesAmarela[i][j]) {
        row += valuesAmarela[i][j];
      }
    }
    contentAmarela += row + "\n";
  }

  Logger.log("[info] Downloading the file...");
  downloadFile(
    `ETIQUETA PRATA(${name}).csv`,
    contentPrata,
    ss,
    `ETIQUETA AMARELA(${name}).csv`,
    contentAmarela
  );
}

function downloadFile(
  titlePrata = "",
  contentPrata = "",
  ss = "",
  titleAmarela = "",
  contentAmarela = ""
) {
  const f = DriveApp.getFileById(ss.getId());
  const fldrs = f.getParents();
  const fldrID = fldrs.next().getId();

  const filePrata = DriveApp.createFile(titlePrata, contentPrata);
  const fileAmarela = DriveApp.createFile(titleAmarela, contentAmarela);

  const urlPrata = filePrata.getDownloadUrl();
  const urlAmarela = fileAmarela.getDownloadUrl();

  filePrata.moveTo(DriveApp.getFolderById(fldrID));
  fileAmarela.moveTo(DriveApp.getFolderById(fldrID));

  showDownloadModal(urlPrata, titlePrata, urlAmarela, titleAmarela);
}

function showDownloadModal(
  urlPrata,
  etiquetaPrata,
  urlAmarela,
  etiquetaAmarela
) {
  const htmlOutput = HtmlService.createHtmlOutput(
    `<a href="${urlPrata}" style="font-size: 1.2rem; font-family: sans-serif; padding:0">${etiquetaPrata}</a>
    </br></br><a href="${urlAmarela}" style="font-size: 1.2rem; font-family: sans-serif; padding:0">${etiquetaAmarela}</a>`
  );
  htmlOutput.setHeight(100).setWidth(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Tem certeza?");
}
