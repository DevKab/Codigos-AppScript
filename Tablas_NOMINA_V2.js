function tablaNomina() {
  const regla = SpreadsheetApp.newDataValidation()
    .requireCheckbox("TRUE", "FALSE")
    .build();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PLANTILLA = ss.getSheetByName("Plantilla_Tablas");
  const SHEET = ss.getSheetByName("Formato Nomina Ejemplo");

  var totalRange = (SHEET.getRange(1001,11,800,6).getValues().slice(1)).filter(fila =>
    fila[0] != "" && 
    fila[0] != null && 
    fila[0] != `USUARIO FINAL` &&
    fila[2] != 0);

  var nomInfo = totalRange.filter(fila => 
    fila[5] == `NOMINA SEMANAL`).map(fila => [fila[0], fila[2], fila[1]]);

  totalRange = totalRange.map(fila => [fila[0], fila[5], fila[2]]);

  var afiInfo = totalRange.filter(fila => 
    fila[1] == `IMPUESTO SOBRE LA NOMINA` || 
    fila[1] == `AGUINALDOS`);

  var bonoInfo = totalRange.filter(fila => 
    fila[1] == `BONO LEALTAD` || 
    fila[1] == `BONO DESPENSA` || 
    fila[1] == `BONO TRANSPORTE`);
     
  var empInfo = totalRange.filter(fila => 
    fila[1] == `EMPLEADO DEL MES`);

  var kpiInfo = totalRange.filter(fila => 
    fila[1] == `BONO MENSUAL`);

  // Validacion para Bonos segun semana
  var diasBono = 22; // 22 Cuarto Viernes para Bonos Bienestar 🟢
  // var diasBono = 1; // 1 Primer Viernes (PRUEBA)
  (getFridayOfMonth(diasBono))?afiInfo = [...afiInfo, ...bonoInfo]:0;

  var diasBono = 8; // 8 Segundo Viernes para Bono KPIs 🔴
  // var diasBono = 1; // 1 Primer Viernes (PRUEBA)
  (getFridayOfMonth(diasBono))?afiInfo = [...afiInfo, ...kpiInfo]:0;

  var diasBono = 15; // 15 Tercer Viernes para Bono Empleado del Mes 🟡
  // var diasBono = 1; // 1 Primer Viernes (PRUEBA)
  (getFridayOfMonth(diasBono))?afiInfo = [...afiInfo, ...empInfo]:0;

  //  afiInfo contiene todos los datos que se muestran en la segunda tabla.
  //  V2 no se muestran datos de Aguinaldos ni Impuesto sobre la Nomina
  //  Incluir tabla aparte de Horas Extras

  const plantEncabezados2 = PLANTILLA.getRange(1,1,7,16);
  const plantNomFormatRange = PLANTILLA.getRange(8,1,nomInfo.length,5);
  plantEncabezados2.copyFormatToRange(SHEET, 1, 11, 6, 12);
  SHEET.getRange(`A6:P12`).setValues(plantEncabezados2.getValues());
  plantNomFormatRange.copyFormatToRange(SHEET, 1, 5, 13, nomInfo.length+7);

  SHEET.getRange(`A7`).setFormula(PLANTILLA.getRange("A2").getFormula());
  SHEET.getRange(`G7`).setFormula(PLANTILLA.getRange("G2").getFormula());  
  SHEET.getRange(13,5,nomInfo.length,1).setFormulas(PLANTILLA.getRange(8,5,nomInfo.length,1).getFormulas());
  SHEET.getRange(13,4,nomInfo.length,1).setValues(PLANTILLA.getRange(8,4,nomInfo.length,1).getValues());
  SHEET.getRange(13,1,nomInfo.length,3).setValues(nomInfo);

  if(afiInfo.length>0){
    const plantAllFormatRange = PLANTILLA.getRange(8,7,afiInfo.length,5);
    plantAllFormatRange.copyFormatToRange(SHEET, 7, 11, 13, afiInfo.length+12);
    SHEET.getRange(13,11,afiInfo.length,1).setFormulas(PLANTILLA.getRange(8,11,afiInfo.length,1).getFormulas());
    // SHEET.getRange(13+afiInfo.filter(fila => fila[1] == `AGUINALDOS` || fila[1] == `IMPUESTO SOBRE LA NOMINA`).length,10,afiInfo.length,1).setDataValidation(regla);
    try {
      const offset = afiInfo.filter(
        fila => fila[1] === 'AGUINALDOS' || fila[1] === 'IMPUESTO SOBRE LA NOMINA'
      ).length;
      const startRow = 13 + offset;
      const numRows = afiInfo.filter(
        fila => fila[1] != 'AGUINALDOS' && fila[1] != 'IMPUESTO SOBRE LA NOMINA'
      ).length;
      if (numRows > 0) {
        SHEET
          .getRange(startRow, 10, numRows, 1)
          .setDataValidation(regla);
      }
    } catch (err) {
      Logger.log('Error al aplicar validación: ' + err.message);
      SpreadsheetApp.getActiveSpreadsheet().toast('Error al aplicar validación: ' + err.message)
    }
    SHEET.getRange(13,7,afiInfo.length,3).setValues(afiInfo);
    (SHEET.getRange(afiInfo.length+12,8,1,1).getValue()==`EMPLEADO DEL MES`)?SHEET.getRange(afiInfo.length+12,7,1,1)
      .setDataValidation(SHEET.getRange(`K1402`).getDataValidation()):0;
    return
  }
  SHEET.getRange(`G10:K12`).clear();
}


// Función actualizada para cuarto viernes del mes (day = 5)
function getFridayOfMonth(diasBono) {
// function getFridayOfMonth() {
  // var diasBono = 15;
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`Formato Nomina Ejemplo`);
  if(hoja.getRange(`R5`).getValue()){
    var pruebaFecha = hoja.getRange(`R2`).getValue();
    var year = (new Date(pruebaFecha)).getFullYear();
    var month = (new Date(pruebaFecha)).getMonth();
    var date = (new Date(pruebaFecha)).getDate();
    const fF = new Date(year, month, diasBono); // 0-Primera, 7-Segunda, 14-Tercera, 21-Cuarta
    const thisFriday = new Date(year, month, date+2);
    while (fF.getDay() !== 5) fF.setDate(fF.getDate() + 1); // 5 = viernes
    // console.log(`${getWeekNumber(fF)} - ${getWeekNumber(thisFriday)}`);
    return getWeekNumber(fF)==getWeekNumber(thisFriday);
  }

  var year = (new Date()).getFullYear();
  var month = (new Date()).getMonth();
  var date = (new Date()).getDate();
  const fF = new Date(year, month, diasBono); // 0-Primera, 7-Segunda, 14-Tercera, 21-Cuarta
  const thisFriday = new Date(year, month, date+2);
  while (fF.getDay() !== 5) fF.setDate(fF.getDate() + 1); // 5 = viernes
  // console.log(getWeekNumber(fF)==getWeekNumber(thisFriday));
  return getWeekNumber(fF)==getWeekNumber(thisFriday);
}

//////////////////////////////

function resetAfi(){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`Formato Nomina Ejemplo`);
  const nombres = hoja.getRange("K1602:K1700").getValues().flat().filter(String);
  var values = Array(nombres.length).fill([0]);
  hoja.getRange(1602, 13, nombres.length, 1).setValues(values);
  hoja.getRange(1702, 13, nombres.length, 1).setValues(values);
}

//////////////////////////////

function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 5; // Viernes = 5
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNum = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return weekNum;
}

//////////////////////////////

function probarSemana() {
  var fecha = new Date("2025-09-19"); // ejemplo
  var numSemana = getWeekNumber(fecha);
  Logger.log("Semana ISO: " + numSemana);
}

function delTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET = ss.getSheetByName("Formato Nomina Ejemplo");

  const nomRange = SHEET.getRange("A4:P1000");
  nomRange.clearContent().clearFormat().clearDataValidations();
}

//////////////////////////////

function dataValidation(){
  const regla = SpreadsheetApp.newDataValidation()
    .requireCheckbox("TRUE", "FALSE")
    .build();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plantilla_Tablas");
    var range = sheet.getRange("J8");
    range.setDataValidation(regla);
}
