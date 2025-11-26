  //  Constantes globales
const SSID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SH_NAME = `PRUEBA`;

  //  Funcion para realizar el corte de dia
function corteDia() {
  var sheet = SpreadsheetApp.openById(SSID).getSheetByName(SH_NAME);

    //  Se obtiene la region de datos de la tabla
  var dataRegion = sheet.getRange("D6").getDataRegion();
  var valores = (dataRegion.getValues().filter(fila => fila[0] !== "" && fila[0] !== null)).map(fila => fila[0]).flat();
  var lastRow = valores.length+3;

    //  Se obtiene la columna que contiene las fechas para corroborar que sea el dia correcto
  var colFechas = sheet.getRange(6, 8, valores.length-3,1).getValues();
  const fecha = Utilities.formatDate(
      new Date(sheet.getRange(lastRow,8).getValue()),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );
  var fechasFiltradas = colFechas.filter(fila => {
    var fechaCelda = Utilities.formatDate(
      new Date(fila[0]),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );
    return fechaCelda === fecha;
  });

    //  Se filtran los datos para sumar los montos que corresponden al dia correcto
  var valoresFiltrados = valores.slice(-(fechasFiltradas.length+1));
  var montoTotal = valoresFiltrados.reduce((acc, num) => {
    return acc + (typeof num === `number` ? num : 0);
  }, 0);

    //  Se construye la estructura dedatos que se insertara al lado derecho
  var totalConf = [
    [``,``,``,`CONFIRMACION`,``],
    [``,``,``,`PROMOTOR`,`INTERNO`],
    // [`TOTAL AL DIA`, fechaAyer, montoTotal, false, false],
    [`TOTAL AL DIA`, fecha, montoTotal, false, false],
  ];
  var finalRange = sheet.getRange(lastRow-2, 11, 3, 5);
  finalRange.setValues(totalConf);

    //  Creacion de regla de validacion (Casilla check)
  var rule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();

    //  Aplica formato a los datos insertados
  sheet.getRange(lastRow-2, 14, 1, 2).merge();
  sheet.getRange(lastRow-2, 11, 3, 5).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(lastRow, 13).setNumberFormat("$#,##0.00");
  sheet.getRange(lastRow, 12, 1, 2).setHorizontalAlignment("right");
  sheet.getRange(lastRow, 14, 1, 2).setDataValidation(rule);
}
