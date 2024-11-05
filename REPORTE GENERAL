function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu("💳 BACKUP");
  menu.addSeparator();
  menu.addItem("💾 Crear Backup", "copiarArchivos");
  menu.addItem("📡 Enviar info. archivo Yessi", "MandaInfoThisToYessi");
  menu.addSeparator();
  menu.addToUi();
}

function Copiaaaaas() {
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");
  var name = "REPORTE GENERAL pruebas " + formattedDate;
  var destination = DriveApp.getFolderById("1NV8_cqvNfXfkDEzxE7VMaxUa4J2nUhUh"); //test

  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
  file.makeCopy(name, destination);
  var nombreFile = name;
  var carpeta = destination;
  var docs = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (docs.hasNext()) {
    var doc = docs.next();
    if (doc == nombreFile) {
      const libroOrigen = SpreadsheetApp.openById('1U9xTBA9bi3lJRNX-q1QgBmSv_LtT0VDoPNXn_sYuAko'); // hoja donde se copia la informacion

      const hojaOrigen = libroOrigen.getSheetByName("CONCENTRADO") //nombre hoja donde se copia
      const hojaOrigen2 = libroOrigen.getSheetByName("EC SD") //nombre hoja donde se copia
      const hojaOrigen3 = libroOrigen.getSheetByName("MOVS DIARIOS") //nombre hoja donde se copia

      const libroDestino = SpreadsheetApp.openById(doc.getId()); // id archivo nuevo

      const hojaDestino = libroDestino.getSheetByName("CONCENTRADO") // hoja donde se va pegar
      const hojaDestino2 = libroDestino.getSheetByName("EC SD") // hoja donde se va pegar
      const hojaDestino3 = libroDestino.getSheetByName("MOVS DIARIOS") // hoja donde se va pegar

      const rangoOrigen = hojaOrigen.getRange(2, 1, hojaOrigen.getLastRow(), hojaOrigen.getLastColumn()).getValues();
      const rangoDestino = hojaDestino.getRange(2, 1, hojaOrigen.getLastRow(), hojaOrigen.getLastColumn());

      const rangoOrigen2 = hojaOrigen2.getRange(2, 1, hojaOrigen2.getLastRow(), hojaOrigen2.getLastColumn()).getValues();
      const rangoDestino2 = hojaDestino2.getRange(2, 1, hojaOrigen2.getLastRow(), hojaOrigen2.getLastColumn());

      const rangoOrigen3 = hojaOrigen3.getRange(2, 1, hojaOrigen3.getLastRow(), hojaOrigen3.getLastColumn()).getValues();
      const rangoDestino3 = hojaDestino3.getRange(2, 1, hojaOrigen3.getLastRow(), hojaOrigen3.getLastColumn());

      rangoDestino.setValues(rangoOrigen);
      rangoDestino2.setValues(rangoOrigen2);
      rangoDestino3.setValues(rangoOrigen3);
    }
  }
}


function MandaInfoThisToYessi() {
  const libroOrigen = SpreadsheetApp.openById('1U9xTBA9bi3lJRNX-q1QgBmSv_LtT0VDoPNXn_sYuAko'); // hoja donde se copia la informacion
  const hojaOrigen = libroOrigen.getSheetByName("MANDAINFOAYESSI") //nombre hoja donde se copia
  const libroDestino = SpreadsheetApp.openById("1y_TtHo1YaT3pWkdhGKkhoD0f5EQ2akiG__tNpfZbJss"); // id donde se va pegar archivo de cada mes //1zSd9R02ke91YEo8ZEPh7FIieZfaBEhIjgQN-KiOu9ZM
  const hojaDestino = libroDestino.getSheetByName("CONCENTRADO") // hoja donde se va pegar
  const rangoOrigen = hojaOrigen.getRange(2, 1, hojaOrigen.getLastRow(), hojaOrigen.getLastColumn()).getValues();
  const UFila = hojaOrigen.getLastRow();
  hojaDestino.insertRowsBefore(2, UFila)
  const rangoDestino = hojaDestino.getRange(2, 1, hojaOrigen.getLastRow(), hojaOrigen.getLastColumn());
  rangoDestino.setValues(rangoOrigen);
}


function copiarArchivos() {
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");
  var nombreArchivo = "REPORTE GENERAL OP " + formattedDate;
  var carpetaDestino = DriveApp.getFolderById("1NV8_cqvNfXfkDEzxE7VMaxUa4J2nUhUh");
  var archivoActivo = SpreadsheetApp.getActiveSpreadsheet();
  var idArchivoActivo = archivoActivo.getId();
  var copiaArchivo = DriveApp.getFileById(idArchivoActivo).makeCopy(nombreArchivo, carpetaDestino);
  var idCopiaArchivo = copiaArchivo.getId();
  var libroOrigen = SpreadsheetApp.openById('1U9xTBA9bi3lJRNX-q1QgBmSv_LtT0VDoPNXn_sYuAko');
  var hojasOrigen = {
    "CONCENTRADO": libroOrigen.getSheetByName("CONCENTRADO"),
    "EC SD": libroOrigen.getSheetByName("EC SD"),
    "MOVS DIARIOS": libroOrigen.getSheetByName("MOVS DIARIOS")
  };
  var libroDestino = SpreadsheetApp.openById(idCopiaArchivo);
  var hojasDestino = {
    "CONCENTRADO": libroDestino.getSheetByName("CONCENTRADO"),
    "EC SD": libroDestino.getSheetByName("EC SD"),
    "MOVS DIARIOS": libroDestino.getSheetByName("MOVS DIARIOS")
  };
  for (var hoja in hojasOrigen) {
    var rangoOrigen = hojasOrigen[hoja].getDataRange();
    var rangoDestino = hojasDestino[hoja].getRange(1, 1, rangoOrigen.getNumRows(), rangoOrigen.getNumColumns());
    rangoDestino.setValues(rangoOrigen.getValues());
  }
    Browser.msgBox("La copia de archivos se ha completado correctamente.");

}


