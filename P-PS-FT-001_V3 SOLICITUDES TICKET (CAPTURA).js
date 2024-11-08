function onEdit(e) {
  // Obtener la hoja activa y su nombre
  var hojaActiva = e.source.getActiveSheet();
  var nombreHoja = hojaActiva.getName();
  
  // Verificar si la hoja activa es "FORMATO DE SOLICITUDES"
  if (nombreHoja === "FORMATO DE SOLICITUDES") {
    // Verificar si la celda editada es B5
    var rangoEditado = e.range;
    if (rangoEditado.getA1Notation() === 'B5') {
      // Obtener el valor de la celda A1
      var valorCeldaA1 = hojaActiva.getRange("A1").getValue();
      
      // Verificar si la celda A1 no está vacía
      if (valorCeldaA1 !== "") {
        // Obtener el correo electrónico del usuario activo
        var correo = Session.getActiveUser().getEmail();
        
        // Establecer el correo electrónico en la celda D3
        hojaActiva.getRange("D3").setValue(correo);
        Logger.log(correo);
      }
    }
  } else {
    Logger.log("La función solo debe ejecutarse en la hoja 'FORMATO DE SOLICITUDES'.");
  }
}

function chequeoCeldasllenas() { /*13/08/2024 */ //me quede aqui!!!!!!!!!!!1:00pm
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("FORMATO DE SOLICITUDES");

  //var range = sheet.getRange("B7");
  var mapeoCeldas = [
    {origen: sheet.getRange('A5')},
    {origen: sheet.getRange('B5')},
    {origen: sheet.getRange('C5')},
    {origen: sheet.getRange('A7')},
    {origen: sheet.getRange('B7:C7')},
    {origen: sheet.getRange('A9')},
    {origen: sheet.getRange('B9')},
    {origen: sheet.getRange('C9')},
    {origen: sheet.getRange('A11')},
    {origen: sheet.getRange('B11')},
    {origen: sheet.getRange('C11')},
    {origen: sheet.getRange('A13')},
    {origen: sheet.getRange('B13')}
  ];
    
    var camposLlenos = true;

  mapeoCeldas.forEach(function (mapeo) {
    var rangos = mapeo.origen.getValues();
    var emptyCells = rangos.some(row => row[0] === "");

    if (emptyCells) {
      SpreadsheetApp.getUi().alert("Debe llenar todos los campos.");
      camposLlenos = false;
      return;
    }
  });
  if(camposLlenos){
       macroinsertar();
    }
}

function macroinsertar() {
  CorreoStatic();

  // Abre el libro de cálculo específico y obtiene la hoja "LISTA DE SOLICITUDES"
  const libroDestino = SpreadsheetApp.openById("1FSIzhw9fcJlVrfFhO6pM3xEFHMEo7W7gyBWQeXEy_3s");
  const listaSheet = libroDestino.getSheetByName("LISTA DE SOLICITUDES");
  const formatoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FORMATO DE SOLICITUDES');
  
  // Inserta una nueva fila después de la última fila en "LISTA DE SOLICITUDES"
  let lastRow = listaSheet.getLastRow() + 1;
  listaSheet.insertRowAfter(lastRow - 1);

  // Copiar y pegar datos de FORMATO DE SOLICITUDES a LISTA DE SOLICITUDES
  listaSheet.getRange('B' + lastRow).setValue(formatoSheet.getRange('A3').getValue());
  listaSheet.getRange('A' + lastRow).setValue(formatoSheet.getRange('C3').getValue());
  listaSheet.getRange('C' + lastRow).setValue(formatoSheet.getRange('B3').getDisplayValue());
  listaSheet.getRange('E' + lastRow).setValue(formatoSheet.getRange('D3').getValue());
  
  listaSheet.getRange('F' + lastRow).setValue(formatoSheet.getRange('A5').getValue());
  listaSheet.getRange('G' + lastRow).setValue(formatoSheet.getRange('B5').getValue());
  listaSheet.getRange('H' + lastRow).setValue(formatoSheet.getRange('C5').getValue());

  listaSheet.getRange('I' + lastRow + ':J' + lastRow).setValues(formatoSheet.getRange('A7:B7').getValues());

  listaSheet.getRange('K' + lastRow).setValue(formatoSheet.getRange('A9').getValue());
  listaSheet.getRange('L' + lastRow).setValue(formatoSheet.getRange('B9').getValue());
  listaSheet.getRange('O' + lastRow).setValue(formatoSheet.getRange('C9').getDisplayValue());

  listaSheet.getRange('M' + lastRow).setValue(formatoSheet.getRange('A11').getValue());
  listaSheet.getRange('R' + lastRow).setValue(formatoSheet.getRange('C11').getValue());
  listaSheet.getRange('P' + lastRow).setValue(formatoSheet.getRange('A13').getValue());
  listaSheet.getRange('Q' + lastRow).setValue(formatoSheet.getRange('B13').getValue());

  var valorB11 = formatoSheet.getRange("B11").getValue();
  listaSheet.getRange('N' + lastRow).setValue(valorB11);


  // Copiar fórmula de la fila anterior
  listaSheet.getRange('D' + lastRow).setFormula(listaSheet.getRange('D' + (lastRow - 1)).getFormula());

  // Agregar fórmulas específicas
  listaSheet.getRange('V' + lastRow).setFormula('=IF(U' + lastRow + '="TERMINADO", IF(V' + lastRow + '="", NOW(), V' + lastRow + '),"")');

  listaSheet.getRange('Y' + lastRow).setFormula('=ROUNDDOWN((W' + lastRow + '-C' + lastRow + '))');

  // Quitar formato y restaurar valores
  ['C', 'F', 'H', 'M'].forEach(col => {
    let cell = listaSheet.getRange(col + lastRow);
    let value = cell.getValue();
    cell.clearDataValidations().clearContent().setValue(value);
  });

  // Enviar correos y agregar borde
  correos();
  listaSheet.getRange(lastRow, 1, 1, listaSheet.getLastColumn())
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  mostrarVentanaEmergenteConImagen();
  macroquitar();
}

function CorreoStatic() {/*modificado 13/08/2024 */ //funciona, si llegan los tickets //para los jefes.
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet1 = ss.getSheetByName('FORMATO DE SOLICITUDES');
  var emailAddress = "sistemas3@kabzo.org, optimizacion@kabzo.org"; /*define las direcciones de correo electrónico a las que se enviará el correo. En este caso, se enviará a dos destinatarios: "arquitectura@kabzo.org" y "ernesto@moreandkitchen.com".*/
  var subject = "Solicitud De Gastos: " + sheet1.getRange(3, 1).getValue();
  var message = "Descripcion De La Solicitud: " + sheet1.getRange(7, 2).getValue() + "\n" +
    "-----------------------------------------\n" +
    "ID: " + sheet1.getRange(3, 3).getValue() + "\n" +
    "-----------------------------------------\n" +
    ss.getUrl();
  MailApp.sendEmail(emailAddress, subject, message);
}

function correos() {/*modificado 13/08/2024 */ //el que esta haciendo el ticket
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet1 = ss.getSheetByName('FORMATO DE SOLICITUDES');
  var emailAddress = Session.getActiveUser().getEmail();
  var subject = "Solicitud De Gastos: " + sheet1.getRange(3, 1).getValue();
  var message = "Descripcion De La Solicitud: " + sheet1.getRange(7, 2).getValue() + "\n" +
    "-----------------------------------------\n" +
    "ID: " + sheet1.getRange(3, 3).getValue() + "\n" +
    "-----------------------------------------\n" +
    ss.getUrl();
  MailApp.sendEmail(emailAddress, subject, message);
}

function mostrarVentanaEmergenteConImagen() {
  var html = HtmlService.createHtmlOutput('<img src="https://www.freeiconspng.com/thumbs/success-icon/success-icon-10.png">')
    .setWidth(412)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '¡Se envió con éxito!!');
}

function macroquitar() {//ELIMINA LA TABLA DE TICKET
  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange('A5').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('B5').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('C5').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('D3').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('A7').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('B7').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('B7').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('A9').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('B9').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('C9').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('A11').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('B11').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('C11').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('A13').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
  spreadsheet.getRange('B13').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
};
