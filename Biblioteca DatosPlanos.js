/**
 * Biblioteca de utilidades para manipular datos en Google Sheets
 * @author desarrolladorjr@kabzo.org
 * @ID de Implementación: 1VOwFFtMS4Rb1tb7nuxwhgbzz8E4WldwqOt_X-l03qBhTql3n4cRmEE2a
 * Convierte en datos planos las filas cuya fecha sea más antigua a 1 semana.
 * 
 * @param {string} SHIT - Parametro vacio
 * @param {string} SHID - Nombre de la hoja
 * @param {string} SSID - ID del Spreadsheet
 * @param {number} row - Fila inicial desde donde tomar los datos
 * @param {number} col - Índice de la columna donde está la fecha (comenzando en 0)
 **/

function datos(SSID, SHID, row, col) {
  var hoja = SpreadsheetApp.openById(SSID).getSheetByName(SHID);
  var datos = hoja.getRange(row, 1, hoja.getLastRow(), hoja.getLastColumn()).getValues(); // Incluye encabezado
  var hoy = new Date();
  var unaSemana = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - 8);
  // Filtramos los datos que cumplen con la condición
  var datosFiltrados = datos.filter((fila, index) => {
    if (index === 0) return true; // mantener encabezado
    var fecha = new Date(fila[col]); 
    return fecha <= unaSemana;
  });
  
  // Limpiar formato y validaciones de los datos encontrados
  hoja.getRange(row, 1, datosFiltrados.length-1, datosFiltrados[0].length)
    .clearFormat()
    .clearDataValidations();
  
  // Mensaje de alerta
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;">' +
    '<center><h2 style="color:#CC0000;">¡Atención!</h2></center>' +
    `<p>Los datos más antiguos de 1 semana de la hoja: <b>${SHID}</b> se han convertido a planos.</p>` +
    '</div>'
  ).setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, "Datos Planos");
}
