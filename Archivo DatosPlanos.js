/*
 *  Archivo .gs que se aloja en el Sheets donde se va a ejecutar
 *  Biblioteca "DatosPlanosBiblioteca" es necesaria 
 *  row - Fila donde empiezan los datos que se quieren aplanar
 *  col - Columna donde se encuentra la fecha a evaluar
 */

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🗒️ Borrar Formato")
  .addItem('Datos Planos','DatosPlanos')
  .addToUi();
}

//////////////////////////////

function DatosPlanos(row, col) {
  var row = 6;
  var col = 1; // Columna donde se encuentra la fecha (Arreglo B = 1)
  var SSID = SpreadsheetApp.getActiveSpreadsheet().getId();
  var SHID = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getName();
  if (!SHID || !SHID.startsWith("S.Gastos ")){ mostrarAlertaConFormato(SHID); return;}
  DatosPlanosBiblioteca.datos(SSID, SHID, row, col);
}

//////////////////////////////

function mostrarAlertaConFormato(nombreHoja) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif; text-align:center; padding:5px; font-size:24px;">
      La hoja:<br>
      ⚠️ <b>"${nombreHoja}"</b> ⚠️<br>
      no es válida.<br><br>
      (Formato válido: "S.Gastos AREA")
     </div>`
  ).setWidth(600).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, `❗ HOJA "${nombreHoja}" INVÁLIDA ❗`);
}
