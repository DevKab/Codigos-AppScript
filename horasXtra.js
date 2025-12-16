const SHEET_NAME = `Formato Nomina Ejemplo`;
const PLANTILLA_NAME = `Plantilla_Tablas`;

function horasXtra(data) {
  const ss = SpreadsheetApp.openById(SSID);
  const SHEET = ss.getSheetByName(SHEET_NAME);
  const PLANTILLA = ss.getSheetByName(PLANTILLA_NAME);
  const lastRow = SHEET.getRange(`M12`).getDataRegion().getLastRow()+1;
  var arrNom = SHEET.getRange(`K1002:M1100`).getValues().filter(fila => fila[0] != "" && fila[0] != null)
    .map(fila => [fila[0], fila[2]]);
  var objNom = arrToObject(arrNom);
  var puNom = objNom[data.name];    //  Precio Unitario Nominas (Obtener dependiendo del data.name)
    // data.name
    // data.horas
  var formula = 
  `=IF(O${lastRow}*1>8,
  IF(O${lastRow}*1>16,(N${lastRow}*8)+(N${lastRow}*16)+(N${lastRow}*(O${lastRow}-16)*3),(N${lastRow}*8)+(N${lastRow}*(O${lastRow}-8)*2)),
N${lastRow}*O${lastRow})`;

  var arrXtra = [[
    data.name,
    (puNom/8),   //  (Precio Unitario Nomina Semanal)/8 (?)
    data.horas,
    formula
  ]];

  PLANTILLA.getRange(8,13,1,4).copyFormatToRange(SHEET,13,16,lastRow,lastRow);
  // SHEET.getRange(lastRow,13,1,1).setValues(PLANTILLA.getRange(8,13,1,1).getValues());
  SHEET.getRange(lastRow,13,1,4).setValues(arrXtra);
}

//////////////////////////////

function arrToObject(data) {
  let obj = {};
  data.forEach(fila => {
    obj[fila[0]] = fila[1];
  });
  return obj;
}
