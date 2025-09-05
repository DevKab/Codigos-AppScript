function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📃 Duplicacion ')
    .addItem('📄 Duplicar hoja OPERACION / VERIFICACION B', 'func')
    .addToUi();
}


function func(){
  nombreHoja()
  limpiadoHojasOPERACIONES_VERIFICAION_B()
}

function nombreHoja(){
  duplicarYAplicar("OPERACION");
  duplicarYAplicar("VERIFICACION B");
}

function duplicarYAplicar(nombreHoja) {//funciona con formato y texto
  var hojaNombre = nombreHoja;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(hojaNombre);
  var currentDate = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");

  if (!hoja) {
    Logger.log("La hoja '" + hojaNombre + "' no existe.");
    return;
  }

  // Duplicar la hoja
  var nuevaHoja = hoja.copyTo(ss);
  nuevaHoja.setName(hojaNombre + " " + currentDate);
  
  // Obtener el rango usado en la nueva hoja
  var rango = nuevaHoja.getDataRange();
  var formulas = rango.getFormulas();
  var valores = rango.getValues();

  // Reemplazar donde haya fórmula por el valor correspondiente
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j]) { // Si hay fórmula
        rango.getCell(i + 1, j + 1).setValue(valores[i][j]); // Poner valor
      }
    }
  }

  Logger.log("Hoja duplicada y bloqueos aplicados correctamente.");
}

//CREAR BOTON PARA SACAR COPIA Y LIMPIAR LA BASE
function limpiadoHojasOPERACIONES_VERIFICAION_B() {//FUNCIONA 08/08/2025
  try {
    var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();// Obtén la hoja de cálculo activa

      var hojas = [
        { nombre: "OPERACION", rangos: ["C2", "C16", "C30", "C44", "N2", "N16", "N30", "N44", "Y2", "Y16", "Y30", "Y44", "AJ2", "AJ16", "AJ30", "AJ44", "AU2", "AU16", "AU30", "AU44", "BF2", "BF16", "BF30", "BF44",
        "C4:C14", "C18:C28", "C32:C42", "C46:C56", "G4:G14", "G18:G28", "G32:G42", "G46:G56", "N4:N14", "N18:N28", "N32:N42", "N46:N56", "R4:R14", "R18:R28", "R32:R42", "R46:R56", "Y4:Y14", "Y18:Y28", "Y32:Y42", "Y46:Y56", "AC4:AC14", "AC18:AC28", "AC32:AC42", "AC46:AC56", "AJ4:AJ14", "AJ18:AJ28", "AJ32:AJ42", "AJ46:AJ56", "AN4:AN14", "AN18:AN28", "AN32:AN42", "AN46:AN56", "AU4:AU14", "AU18:AU28", "AU32:AU42", "AU46:AU56", "AY4:AY14", "AY18:AY28", "AY32:AY42", "AY46:AY56", "BF4:BF14", "BF18:BF28", "BF32:BF42", "BF46:BF56", "BJ4:BJ14", "BJ18:BJ28", "BJ32:BJ42", "BJ46:BJ56"] },
        {
          nombre: "VERIFICACION B", rangos: ["C3:C44", "C47:C84", "G3:G44", "G47:G84", "K3:K44", "K47:K84", "O3:O44", "O47:O84", "S3:S44", "S47:S84", "W3:W44", "W47:W84", "AA3:AA44", "AA47:AA84", "AE3:AE44", "AE47:AE84", "AI3:AI44", "AI47:AI84", "AM3:AM44", "AM47:AM84", "AQ3:AQ44", "AQ47:AQ84",
          "C1", "G1", "K1", "O1", "S1", "W1", "AA1", "AE1", "AI1", "AM1", "AQ1"]
        }
      ];

      hojas.forEach(function (hoja) {
        var sheet = hojaDeCalculo.getSheetByName(hoja.nombre);
        hoja.rangos.forEach(function (rango) {
          sheet.getRange(rango).clearContent();
        });
      });

  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}

///codigo limpio//
function botonCopiaLimpioBanck(){
  copiarhojasOPERACIONES_VERIFICACION_B();
  limpiadoHojasOPERACIONES_VERIFICAION_B()
}

//CREAR BOTON PARA SACAR COPIA Y LIMPIAR LA BASE
function limpiadoHojasOPERACIONES_VERIFICAION_B() {//FUNCIONA 08/08/2025
  try {
    var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();// Obtén la hoja de cálculo activa

      var hojas = [
        { nombre: "OPERACION", rangos: ["C2", "C16", "C30", "C44", "N2", "N16", "N30", "N44", "Y2", "Y16", "Y30", "Y44", "AJ2", "AJ16", "AJ30", "AJ44", "AU2", "AU16", "AU30", "AU44", "BF2", "BF16", "BF30", "BF44",
        "C4:C14", "C18:C28", "C32:C42", "C46:C56", "G4:G14", "G18:G28", "G32:G42", "G46:G56", "N4:N14", "N18:N28", "N32:N42", "N46:N56", "R4:R14", "R18:R28", "R32:R42", "R46:R56", "Y4:Y14", "Y18:Y28", "Y32:Y42", "Y46:Y56", "AC4:AC14", "AC18:AC28", "AC32:AC42", "AC46:AC56", "AJ4:AJ14", "AJ18:AJ28", "AJ32:AJ42", "AJ46:AJ56", "AN4:AN14", "AN18:AN28", "AN32:AN42", "AN46:AN56", "AU4:AU14", "AU18:AU28", "AU32:AU42", "AU46:AU56", "AY4:AY14", "AY18:AY28", "AY32:AY42", "AY46:AY56", "BF4:BF14", "BF18:BF28", "BF32:BF42", "BF46:BF56", "BJ4:BJ14", "BJ18:BJ28", "BJ32:BJ42", "BJ46:BJ56"] },
        {
          nombre: "VERIFICACION B", rangos: ["C3:C44", "C47:C84", "G3:G44", "G47:G84", "K3:K44", "K47:K84", "O3:O44", "O47:O84", "S3:S44", "S47:S84", "W3:W44", "W47:W84", "AA3:AA44", "AA47:AA84", "AE3:AE44", "AE47:AE84", "AI3:AI44", "AI47:AI84", "AM3:AM44", "AM47:AM84", "AQ3:AQ44", "AQ47:AQ84",
          "C1", "G1", "K1", "O1", "S1", "W1", "AA1", "AE1", "AI1", "AM1", "AQ1"]
        }
      ];

      hojas.forEach(function (hoja) {
        var sheet = hojaDeCalculo.getSheetByName(hoja.nombre);
        hoja.rangos.forEach(function (rango) {
          sheet.getRange(rango).clearContent();
        });
      });

  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}

function copiarhojasOPERACIONES_VERIFICACION_B() { //saca a una copia de g2 y de ENTRECUENTAS ==funciona == 09/01/2025
  var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();
  var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  var nombreArchivo = hojaDeCalculo.getName();
  var hojasDatos = ["VERIFICACION B", "OPERACION"];

  var carpetaBackup = DriveApp.getFolderById("1DIdhlDHlUvUFlK2yRGL4yu0ARAhXmp9-");//id de la carpeta DE OPERACIONES. 

  //Crear un nuevo archivo donde se copiaran las hojas
  var nombreBackup = 'Backup - ' + nombreArchivo + ' - ' + currentDate;
  nuevaHojaDeCalculo = SpreadsheetApp.create(nombreBackup);

  hojasDatos.forEach(function (hojaNombre) {
    var hojaOrigen = hojaDeCalculo.getSheetByName(hojaNombre);
    if (!hojaOrigen) {
      Logger.log('No se encontró la hoja con el nombre: ' + hojaNombre);
      return;
    }

    // Copiar la hoja al archivo nuevo
    var hojaNueva = hojaOrigen.copyTo(nuevaHojaDeCalculo);
    hojaNueva.setName(hojaNombre);
  });

  // Eliminar la hoja inicial creada al momento de crear el nuevo archivo
  var hojaInicial = nuevaHojaDeCalculo.getSheets()[0];
  nuevaHojaDeCalculo.deleteSheet(hojaInicial);

  // Mover el archivo a la carpeta de respaldo
  var idNuevoArchivo = nuevaHojaDeCalculo.getId();
  DriveApp.getFileById(idNuevoArchivo).moveTo(carpetaBackup);
}
