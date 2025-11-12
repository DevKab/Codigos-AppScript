const SSID = SpreadsheetApp.getActiveSpreadsheet().getId();
const TABLAS_SHEET = "Formato Nomina Ejemplo";

function addColab(data) {
  if(
    data.quienSol == "" || data.quienSol == null ||
    data.dptoSol == "" || data.dptoSol == null ||
    data.dondeAplica == "" || data.dondeAplica == null ||
    data.colab == "" || data.colab == null ||
    data.sueldoBase == "" || data.sueldoBase == null ||
    data.fechaIngreso == "" || data.fechaIngreso == null
  ){
      var htmlIncompleto = HtmlService.createHtmlOutput(`
    <html>
      <body style="
        text-align:center;
        margin:0;
        padding:20px;
        background:white;
        font-family:Arial, sans-serif;
      ">
        <h3>Debe llenar todos los campos!</h3>
        <p>Favor de intentarlo de nuevo.</p>
        <button 
          style="
            margin-top:15px;
            background-color:#1a73e8;
            color:white;
            border:none;
            padding:8px 16px;
            border-radius:6px;
            cursor:pointer;
            font-size:14px;
          "
          onclick="google.script.host.close()"
        >
          Aceptar
        </button>
      </body>
    </html>
  `).setWidth(300)
  .setHeight(210);
  SpreadsheetApp.getUi().showModalDialog(htmlIncompleto, 'Datos Incompletos!');
    return
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(`Agregando...`);
  var tablasSheet = SpreadsheetApp.openById(SSID).getSheetByName(TABLAS_SHEET);
  var listaSheet = SpreadsheetApp.openById(SSID).getSheetByName("Colaboradores");
  var semanaLastRow = (tablasSheet.getRange(1001, 1).getDataRegion().getLastRow() + 1);
  var lealtadLastRow = (tablasSheet.getRange(1101, 1).getDataRegion().getLastRow() + 1);
  var despensaLastRow = (tablasSheet.getRange(1201, 1).getDataRegion().getLastRow() + 1);
  var transporteLastRow = (tablasSheet.getRange(1301, 1).getDataRegion().getLastRow() + 1);
  var kpiLastRow = (tablasSheet.getRange(1501, 1).getDataRegion().getLastRow() + 1);

  var listaLastRow = (listaSheet.getRange(`G12`).getDataRegion().getLastRow()+1)
  listaSheet.getRange(`G${listaLastRow}`).setValue((data.colab).toUpperCase());


  var colabSemArray = [[
    "",
    new Date().toISOString().substring(0,10),
    data.quienSol,
    data.dptoSol,
    "DESPACHO",
    "SEMANAL", // Periodicidad - 5
    data.dondeAplica,
    "NOMINAS",
    "NOMINA", // Subcategoria - 8
    "",
    (data.colab).toUpperCase(),
    5, // Cantidad - 11
    (data.sueldoBase/5), // Precio Unitario 12
    "N/A",
    "N/A",
    "NOMINA SEMANAL", // Descripcion - 15
    "SERVICIO",
    "","","","","","",
    `=-(M${semanaLastRow}*L${semanaLastRow})`, // Monto/importe - 23
    "",
    new Date(data.fechaIngreso).toISOString().substring(0,10),
    `=WEEKNUM(TODAY())`,
    `=MONTH(TODAY())`,
    2025,
  ]];

  // BONO LEALTAD
  var colabLealArray = colabSemArray.map(fila => fila.slice());
  colabLealArray[0][5] = "MENSUAL";
  colabLealArray[0][8] = "BONO LEALTAD";
  colabLealArray[0][11] = 1;
  colabLealArray[0][12] = 1000;
  colabLealArray[0][15] = colabLealArray[0][8];
  colabLealArray[0][23] = `=-(M${lealtadLastRow}*L${lealtadLastRow})`;

  // BONO DESPENSA
  var colabDespArray = colabLealArray.map(fila => fila.slice());
  colabDespArray[0][8] = "BONO DESPENSA";
  colabDespArray[0][12] = 400;
  colabDespArray[0][15] = colabDespArray[0][8];
  colabDespArray[0][23] = `=-(M${despensaLastRow}*L${despensaLastRow})`;

  // BONO TRANSPORTE
  var colabTransArray = colabDespArray.map(fila => fila.slice());
  colabTransArray[0][8] = "BONO TRANSPORTE";
  colabTransArray[0][15] = colabTransArray[0][8];
  colabTransArray[0][23] = `=-(M${transporteLastRow}*L${transporteLastRow})`;

  // BONO KPI
  var colabKPIArray = colabDespArray.map(fila => fila.slice());
  colabKPIArray[0][8] = "BONO MENSUAL";
  colabKPIArray[0][12] = 0;
  colabKPIArray[0][15] = colabKPIArray[0][8];
  colabKPIArray[0][23] = `=-(M${kpiLastRow}*L${kpiLastRow})`;

    // NOMINA SEMANAL
  var semanaRange = tablasSheet.getRange(semanaLastRow-1,1,1,29);
  var semanaRange2 = tablasSheet.getRange(semanaLastRow,1,1,29);
  semanaRange.copyFormatToRange(tablasSheet,1,29,semanaLastRow,semanaLastRow);
  semanaRange2.setDataValidations(semanaRange.getDataValidations());
  tablasSheet.getRange(semanaLastRow,1,1,29).setValues(colabSemArray);

    // BONO LEALTAD
  var lealtadRange = tablasSheet.getRange(lealtadLastRow-1,1,1,29);
  var lealtadRange2 = tablasSheet.getRange(lealtadLastRow,1,1,29);
  lealtadRange.copyFormatToRange(tablasSheet,1,29,lealtadLastRow,lealtadLastRow);
  lealtadRange2.setDataValidations(semanaRange.getDataValidations());
  tablasSheet.getRange(lealtadLastRow,1,1,29).setValues(colabLealArray);

    // BONO DESPENSA
  var despRange = tablasSheet.getRange(despensaLastRow-1,1,1,29);
  var despRange2 = tablasSheet.getRange(despensaLastRow,1,1,29);
  despRange.copyFormatToRange(tablasSheet,1,29,despensaLastRow,despensaLastRow);
  despRange2.setDataValidations(semanaRange.getDataValidations());
  tablasSheet.getRange(despensaLastRow,1,1,29).setValues(colabDespArray);

    // BONO TRANSPORTE
  var transRange = tablasSheet.getRange(transporteLastRow-1,1,1,29);
  var transRange2 = tablasSheet.getRange(transporteLastRow,1,1,29);
  transRange.copyFormatToRange(tablasSheet,1,29,transporteLastRow,transporteLastRow);
  transRange2.setDataValidations(semanaRange.getDataValidations());
  tablasSheet.getRange(transporteLastRow,1,1,29).setValues(colabTransArray);

    // BONO PRODUCTIVIDAD
  var kpiRange = tablasSheet.getRange(kpiLastRow-1,1,1,29);
  var kpiRange2 = tablasSheet.getRange(kpiLastRow,1,1,29);
  kpiRange.copyFormatToRange(tablasSheet,1,29,kpiLastRow,kpiLastRow);
  kpiRange2.setDataValidations(semanaRange.getDataValidations());
  tablasSheet.getRange(kpiLastRow,1,1,29).setValues(colabKPIArray);

  var html = HtmlService.createHtmlOutput(`
    <html>
      <body style="
        text-align:center;
        margin:0;
        padding:20px;
        background:white;
        font-family:Arial, sans-serif;
      ">
        <h3>Se ha Agregado Correctamente a:</h3>
        <p>${(data.colab).toUpperCase()}</p>
        <button 
          style="
            margin-top:15px;
            background-color:#1a73e8;
            color:white;
            border:none;
            padding:8px 16px;
            border-radius:6px;
            cursor:pointer;
            font-size:14px;
          "
          onclick="google.script.host.close()"
        >
          Aceptar
        </button>
      </body>
    </html>
  `).setWidth(300)
  .setHeight(210);
  SpreadsheetApp.getUi().showModalDialog(html, 'Agregar Colaborador');
  return colabSemArray;
}

function removeColab(colaborador){
  SpreadsheetApp.getActiveSpreadsheet().toast(`Quitando...`);
  const tablasHoja = SpreadsheetApp.openById(SSID).getSheetByName(TABLAS_SHEET);
  const colaboradoresHoja = SpreadsheetApp.openById(SSID).getSheetByName("Colaboradores");
  const nombreColab = (colaborador.name);

  const colaboradores = tablasHoja.getRange("K:K").getValues().flat();
  for (let i = colaboradores.length - 1; i >= 0; i--) {
    if (colaboradores[i] === nombreColab) {
      tablasHoja.deleteRow(i + 1);
      tablasHoja.insertRowAfter(tablasHoja.getRange(i,11).getDataRegion().getLastRow()+1);
      console.log(`Fila ${i + 1} eliminada (Nombre: ${nombreColab})`);
    }
  }

  const colaboradoresNombres = colaboradoresHoja.getRange("G:G").getValues().flat();
  for (let i = colaboradoresNombres.length - 1; i >= 0; i--) {
    if (colaboradoresNombres[i] === nombreColab) {
      colaboradoresHoja.deleteRow(i + 1);
      colaboradoresHoja.insertRowAfter(colaboradoresHoja.getRange(i,11).getDataRegion().getLastRow());
      console.log(`Fila ${i + 1} eliminada (Nombre: ${nombreColab})`);
    }
  }
  pollo(nombreColab);
}

function pollo(nombre){
// function pollo(){
//   var nombre = "POLLO LOCO";
  var html = HtmlService.createHtmlOutput(`<html>
      <body style="
        text-align:center;
        margin:0;
        padding:20px;
        background:white;
        font-family:Arial, sans-serif;
      ">
        <img 
          src="https://static.vecteezy.com/system/resources/previews/020/952/293/non_2x/chicken-hen-standing-free-png.png"
          style="max-width:200px; height:auto;"
        >
        <h3>🚨 ¡Pollo Alert! 🚨</h3>
        <p>Esta es tu alerta del pollo.</p>
        <p>Se ha dado de baja correctamente a:</p>
        <p>${nombre}</p>
        <button 
          style="
            margin-top:15px;
            background-color:#1a73e8;
            color:white;
            border:none;
            padding:8px 16px;
            border-radius:6px;
            cursor:pointer;
            font-size:14px;
          "
          onclick="google.script.host.close()"
        >
          Aceptar
        </button>
      </body>
    </html>`)
  .setWidth(350)
  .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, '🚨 Pollo Alert 🚨');
}

function polloAlert(){
  var html = HtmlService.createHtmlOutput(`<html>
      <body style="
        text-align:center;
        margin:0;
        padding:20px;
        background:white;
        font-family:Arial, sans-serif;
      ">
        <img 
          src="https://static.vecteezy.com/system/resources/previews/020/952/293/non_2x/chicken-hen-standing-free-png.png"
          style="max-width:200px; height:auto;"
        >
        <h3>🚨 ¡Pollo Alert! 🚨</h3>
        <p>Esta es tu alerta del pollo.</p>
        <button 
          style="
            margin-top:15px;
            background-color:#1a73e8;
            color:white;
            border:none;
            padding:8px 16px;
            border-radius:6px;
            cursor:pointer;
            font-size:14px;
          "
          onclick="google.script.host.close()"
        >
          Aceptar
        </button>
      </body>
    </html>`)
  .setWidth(350)
  .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, '🚨 Pollo Alert 🚨');
}

function doGet(e){
  var template = HtmlService.createTemplateFromFile("modalColab");
  var html = template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return html;
}

function agregarColaborador (){
  openDialog('modalColab', 'Agregar Nuevo Colaborador');
}

function quitarColaborador (){
  openDialog('quitarColab', 'Quitar Colaborador');
}

function obtenerNombres(){
  const hoja = SpreadsheetApp.openById(SSID).getSheetByName("Colaboradores");
  const nombres = hoja.getRange("G13:G").getValues().flat().filter(String);
  return nombres;
}

// Opens the html file corresponding to the selected option
function openDialog(temp,title) {
  // var html = HtmlService.createHtmlOutputFromFile(temp);
  var template = HtmlService.createTemplateFromFile(temp);
  var html = template.evaluate().setTitle(title).setHeight(350);
  SpreadsheetApp.getUi()
  .showModalDialog(html, title);
}

function permisos(){
  SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getRootFolder();
  console.log("Permisos verificados correctamente");
}
