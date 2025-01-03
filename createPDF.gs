let spreadsheet = SpreadsheetApp.getActive();
const OUTPUT_FOLDER_NAME = 'Rotulos';
const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
const sheet = SpreadsheetApp.getActive().getSheetByName('Rotulos').getSheetId();
let fecha = spreadsheet.getRange('Productos!A2').getValue().toLocaleString().split(', '); 
const pdfName = 'Rotulo ' + fecha[0];

function getFolderByName_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    // La carpeta no existe, puedes crearla si lo deseas
    return DriveApp.createFolder(folderName);
  }
}

function createPDF(ssId, sheet, pdfName) {
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0&" +
    "bottom_margin=0.25&" +
    "left_margin=0&" +
    "right_margin=0&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Obtiene la carpeta en Drive donde se almacenan los PDF.
  const folder = getFolderByName_(OUTPUT_FOLDER_NAME);

  const pdfFile = folder.createFile(blob);

  // Crea una URL para descargar y abrir el PDF.
  const downloadUrl = pdfFile.getDownloadUrl();

  // Crea un enlace con estilo de botón en una sola línea de HTML
  let htmlOutput = HtmlService.createHtmlOutput('<div style="display:grid; justify-content:center;"><a style="display:inline-block;padding:10px 20px;background-color:#096176;color:#fff;text-decoration:none;border-radius:5px;" href="' + downloadUrl + '" target="_blank">Descargar Rotulos</a></div>').setHeight(100).setWidth(300);


  // Abre la interfaz web en una ventana modal.
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Rotulos listos');

  return pdfFile;
}
