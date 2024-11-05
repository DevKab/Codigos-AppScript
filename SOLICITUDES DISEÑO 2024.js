
function macroquitar() {
  var spreadsheet = SpreadsheetApp.getActive();

  // spreadsheet.getRange('B3').activate();
  // spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
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
};

function mostrarVentanaEmergenteConImagen() {
  var html = HtmlService.createHtmlOutput('<img src="https://www.freeiconspng.com/thumbs/success-icon/success-icon-10.png">')
    .setWidth(412)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '¡Se envió con éxito!!');
}

// function obtenerCorreoUsuarioActivo() {
//   var usuarioActivo = Session.getActiveUser();
//   var correo = usuarioActivo.getEmail();

//   var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   var valorCeldaA7 = hoja.getRange("A7").getValue();

//   if (valorCeldaA7 !== "") {
//     hoja.getRange("d3").setValue(correo);
//     Logger.log(correo)
//   }
// }

function macroinsertar() {
  CorreoStatic()
  var spreadsheet = SpreadsheetApp.getActive();
  var listaSheet = spreadsheet.getSheetByName('LISTA DE SOLICITUDES');
  var formatoSheet = spreadsheet.getSheetByName('FORMATO DE SOLICITUDES');

  var lastRow = listaSheet.getLastRow();

  listaSheet.insertRowAfter(lastRow);
  lastRow++;

  var formulasRange1 = listaSheet.getRange('H' + (lastRow - 1) + ':J' + (lastRow - 1));
  var formulasRange2 = listaSheet.getRange('P' + (lastRow - 1));

  var targetFormulasRange1 = listaSheet.getRange('H' + lastRow + ':J' + lastRow);
  var targetFormulasRange2 = listaSheet.getRange('P' + lastRow);

  formulasRange1.copyTo(targetFormulasRange1, SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  formulasRange2.copyTo(targetFormulasRange2, SpreadsheetApp.CopyPasteType.PASTE_FORMULA);

  // Copy data from FORMATO DE SOLICITUDES to the new row in LISTA DE SOLICITUDES
  var sourceRange1 = formatoSheet.getRange('A3'); // Adjust the range as needed //LISTO
  var sourceRange2 = formatoSheet.getRange('A5'); // Adjust the range as needed  //LISTO
  var sourceRange25 = formatoSheet.getRange('B5'); // Adjust the range as needed  //LISTO

  var sourceRange3 = formatoSheet.getRange('C5'); // Adjust the range as needed //LISTO
  var sourceRange4 = formatoSheet.getRange('D3'); // Adjust the range as needed
  var sourceRange5 = formatoSheet.getRange('A7:B7'); // tarea a realizar
  var id = formatoSheet.getRange("c3");
  var fecha = formatoSheet.getRange("b3")
  var fechaentrega = formatoSheet.getRange("b8")

  var targetRange1 = listaSheet.getRange('A' + lastRow);
  var pegaid = listaSheet.getRange("b" + lastRow);
  var targetRange2 = listaSheet.getRange('c' + lastRow);
  var targetRange25 = listaSheet.getRange('d' + lastRow);
  var correo = listaSheet.getRange('e' + lastRow);
  var targetRange3 = listaSheet.getRange('f' + lastRow);
  var targetRange4 = listaSheet.getRange('g' + lastRow);
  var targetRange5 = listaSheet.getRange('g' + lastRow + ':h' + lastRow);
  var pegaFecha = listaSheet.getRange("I" + lastRow)
  var pegaFechaentrega = listaSheet.getRange("L" + lastRow)

  sourceRange1.copyTo(targetRange1, SpreadsheetApp.CopyPasteType.PASTE_NORMAL);
  sourceRange2.copyTo(targetRange2, SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  sourceRange4.copyTo(correo, SpreadsheetApp.CopyPasteType.PASTE_NORMAL)
  sourceRange3.copyTo(targetRange3, SpreadsheetApp.CopyPasteType.PASTE_NORMAL);
  sourceRange4.copyTo(targetRange4, SpreadsheetApp.CopyPasteType.PASTE_NORMAL);
  fechaentrega.copyTo(pegaFechaentrega, SpreadsheetApp.CopyPasteType.PASTE_VALUES)

  listaSheet.getRange('N' + lastRow).setFormula('=IFS(M' + lastRow + '<>' + '"TERMINADO"' + ',"",' + 'N' + lastRow + '<>' + '"",' + 'N' + lastRow + ',TRUE,NOW())')
  listaSheet.getRange('O' + lastRow).setFormula('=ROUNDDOWN((N' + lastRow + '-I' + lastRow + '))');

  var valTarea = sourceRange5.getValues()
  targetRange5.setValues(valTarea)

  var valuesToCopy = id.getValues();
  pegaid.setValues(valuesToCopy);

  var valArea = sourceRange25.getValues()
  targetRange25.setValues(valArea)

  var fechape = fecha.getValues();
  pegaFecha.setValues(fechape);

  var cLastRowValue = listaSheet.getRange('C' + lastRow).getValue();

  var cCell = listaSheet.getRange('C' + lastRow);
  cCell.clearDataValidations();
  cCell.clearContent();

  cCell.setValue(cLastRowValue);

  var fLastRowValue = listaSheet.getRange('F' + lastRow).getValue();

  var fCell = listaSheet.getRange('F' + lastRow);
  fCell.clearDataValidations();
  fCell.clearContent();

  fCell.setValue(fLastRowValue);

  correos();
  var borderRange = listaSheet.getRange(lastRow, 1, 1, listaSheet.getLastColumn());
  borderRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  var fecent = formatoSheet.getRange('b8');
  fecent.clear({ contentsOnly: true, skipFilteredRows: true });

  mostrarVentanaEmergenteConImagen()
  macroquitar()
}

function macroinsertar2() {
  correosreapertura();
  CorreoStatic()
  var spreadsheet = SpreadsheetApp.getActive();
  var listaSheet = spreadsheet.getSheetByName('LISTA DE SOLICITUDES');
  var formatoSheet = spreadsheet.getSheetByName('FORMATO DE SOLICITUDES');

  var lastRow = listaSheet.getLastRow() + 1;

  // Array de mapeo de celdas a copiar y pegar
  var mapeoCeldas = [
    { origen: formatoSheet.getRange('A14'), destino: listaSheet.getRange('A' + lastRow) }, // Tipo de ticket
    { origen: formatoSheet.getRange('b18'), destino: listaSheet.getRange('D' + lastRow) }, // Area
    { origen: formatoSheet.getRange('B16'), destino: listaSheet.getRange('I' + lastRow) }, // Fecha de solicitud
    { origen: formatoSheet.getRange('C18'), destino: listaSheet.getRange('F' + lastRow) }, // Categoria
    { origen: formatoSheet.getRange('c16'), destino: listaSheet.getRange('E' + lastRow) }, // Solicito
    { origen: formatoSheet.getRange('A18'), destino: listaSheet.getRange('G' + lastRow) }, // Tarea a realizar
    { origen: formatoSheet.getRange('a16'), destino: listaSheet.getRange('c' + lastRow) }, // Oficina
    { origen: formatoSheet.getRange('C14'), destino: listaSheet.getRange('h' + lastRow) }, // Detalles adicionales
    { origen: formatoSheet.getRange('b14'), destino: listaSheet.getRange('b' + lastRow) }, // Id
  ];

  mapeoCeldas.forEach(function (mapeo) {
    var datos = mapeo.origen.getValues();
    mapeo.destino.setValues(datos);
  });

  // Copiar, borrar y pegar el valor de la celda de Fecha en columna J
  var fechaCell = formatoSheet.getRange('C14:D14');
  fechaCell.clear({ contentsOnly: true, skipFilteredRows: true });

  var texto = formatoSheet.getRange('C14:D14');
  texto.clear({ contentsOnly: true, skipFilteredRows: true });

  // Agregar las fórmulas a las columnas
  listaSheet.getRange('N' + lastRow).setFormula('=IFS(M' + lastRow + '<>' + '"TERMINADO"' + ',"",' + 'N' + lastRow + '<>' + '"",' + 'N' + lastRow + ',TRUE,NOW())')
  listaSheet.getRange('J' + lastRow).setFormula('=CONCATENATE(TEXT(I' + lastRow + ',"MM"),TEXT(I' + lastRow + ',"_mmmm"))');
  listaSheet.getRange('O' + lastRow).setFormula('=ROUNDDOWN((N' + lastRow + '-I' + lastRow + '))');
  listaSheet.getRange('P' + lastRow).setFormula('=ROUNDDOWN(L' + lastRow + '-I' + lastRow + ')-O' + lastRow);

  var borderRange = listaSheet.getRange(lastRow, 1, 1, listaSheet.getLastColumn());
  borderRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  mostrarVentanaEmergenteConImagen()

  var iddel = formatoSheet.getRange('b14');
  iddel.clear({ contentsOnly: true, skipFilteredRows: true });


}

function macrotest() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B7').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE SOLICITUDES'), true);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('H156:H157'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('H156:H157').activate();
};

function macrodelete() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('K138').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
};

function correosreapertura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet1 = ss.getSheetByName('FORMATO DE SOLICITUDES');
  var emailAddress = Session.getActiveUser().getEmail();
  var subject = "Ticket Diseño: " + sheet1.getRange(14, 1).getValue();
  var message = "Actividad: " + sheet1.getRange(14, 3).getValue() + "\n" +
    "-----------------------------------------\n" +
    "ID: " + sheet1.getRange(14, 2).getValue() + "\n" +
    "-----------------------------------------\n" +
    ss.getUrl();
  MailApp.sendEmail(emailAddress, subject, message);
}

function correos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet1 = ss.getSheetByName('FORMATO DE SOLICITUDES');
  var emailAddress = Session.getActiveUser().getEmail();
  var subject = "Ticket Diseño: " + sheet1.getRange(3, 1).getValue();
  var message = "Actividad: " + sheet1.getRange(7, 2).getValue() + "\n" +
    "-----------------------------------------\n" +
    "ID: " + sheet1.getRange(3, 3).getValue() + "\n" +
    "-----------------------------------------\n" +
    ss.getUrl();
  MailApp.sendEmail(emailAddress, subject, message);
}

function CorreoStatic() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet1 = ss.getSheetByName('FORMATO DE SOLICITUDES');
  var emailAddress = "edicionvideo@produccionesdobbleb.com";
  var subject = "Ticket Optimzacion: " + sheet1.getRange(3, 1).getValue();
  var message = "Actividad: " + sheet1.getRange(7, 2).getValue() + "\n" +
    "-----------------------------------------\n" +
    "ID: " + sheet1.getRange(3, 3).getValue() + "\n" +
    "-----------------------------------------\n" +
    ss.getUrl();
  MailApp.sendEmail(emailAddress, subject, message);
}

function macropruebaformato() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('167:167').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
};

function actualizarCeldaBe8() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("FORMATO DE SOLICITUDES");

  var range = sheet.getRange("B8");
  var values = range.getValues();
  var emptyCells = values.some(row => row[0] === "");

  if (emptyCells) {
    SpreadsheetApp.getUi().alert("Debe llenar todos los campos.");
  } else {
    macroinsertar();
  }
}
