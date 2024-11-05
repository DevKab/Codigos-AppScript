
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var mensaje = "Recuerda que esto es una plantilla automatizada:"
    + "\n-  ❌ No agregar o quitar columnas y filas."
    + "\n-  ❌ No alterar fórmulas."
    + "\n-  ❌ No modificar la posición de las tablas o el rango."
    + "\n-  ❌ No copiar y pegar en los rangos de listas anidadas."
    + "\n-  ✔️ Contacta a 'Optimización' para realizar modificaciones.";
  ui.alert('IMPORTANTE', mensaje,
    ui.ButtonSet.OK);
  ui.createMenu('📤 ~ Enviar menu')
    .addItem('🥘 ~ Menu listo', 'actualizarUsuarioActivo')
    .addItem('Backup', 'backupSpreadsheet')
    .addToUi();
}


function actualizarUsuarioActivo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var rangos = [
    { rango: 'A2:J58', escribirCorreo: 'H3', escribirFecha: 'F3' },
    { rango: 'A60:J116', escribirCorreo: 'H61', escribirFecha: 'F61' },
    { rango: 'A118:J174', escribirCorreo: 'H119', escribirFecha: 'F119' },
    { rango: 'A176:J232', escribirCorreo: 'H177', escribirFecha: 'F177' },
    { rango: 'A234:J290', escribirCorreo: 'H235', escribirFecha: 'F235' }
  ];

  var user = Session.getActiveUser().getEmail();
  var fechaActual = new Date();

  rangos.forEach(function (item) {
    var rangoMonitoreo = sheet.getRange(item.rango);
    var escribirCorreo = sheet.getRange(item.escribirCorreo);
    var escribirFecha = sheet.getRange(item.escribirFecha);

    var cambiosRealizados = rangoMonitoreo.getValues().some(row => row.some(cell => cell !== ""));

    if (cambiosRealizados) {
      console.log('Usuario activo:', user);
      escribirCorreo.setValue(user);
      escribirFecha.setValue(fechaActual);
      // console.log(user, fechaActual);
    } else {
      console.log(`No se han realizado cambios dentro del rango ${item.rango}`);
    }
  });
  convertEditorsToViewers()
}

function convertEditorsToViewers() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetId = spreadsheet.getId();
  var file = DriveApp.getFileById(sheetId);
  var editors = file.getEditors();
  for (var i = 0; i < editors.length; i++) {
    var editor = editors[i];
    file.removeEditor(editor);
    file.addViewer(editor);
  }

  Logger.log('Todos los editores han sido cambiados a lectores.');
}

function backupSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); // Obtiene la hoja activa
    var sheet = ss.getSheetByName("Menu vigente"); // Obtiene la hoja "Menu vigente"

    if (!sheet) {
      Logger.log('La hoja "Menu vigente" no existe en el archivo original.');
      return;
    }
    // Obtener el mes actual y sumar un mes
    var date = new Date();
    date.setMonth(date.getMonth() + 1);  // Sumar un mes

    // Opciones para mostrar el mes en formato largo
    var options = { month: 'long' };
    var nextMonthName = date.toLocaleDateString('es-ES', options);

    // Crear el nombre del archivo de backup
    var backupName = `P-PD-FT-003_REV 0.2_${nextMonthName}_${date.getFullYear()}`;

    // Crear una copia de la hoja de cálculo
    var backupSpreadsheet = ss.copy(backupName);

    Logger.log(`Copia de seguridad creada con el nombre: ${backupName}`);

    // Obtener la hoja "Menu vigente" de la copia
    var backupSheet = backupSpreadsheet.getSheetByName("Menu vigente");

    if (!backupSheet) {
      Logger.log('La hoja "Menu vigente" no se encontró en la copia de backup.');
      return;
    }

    // Rango de celdas a limpiar
    var rangesToClear = [
      "A5:C5", "D6:J12", "D15:J18", "D20:J21", "D23:J27", "D30:J33", "D35:J36", "D38:J58",
      "A63:C63", "D64:J70", "D73:J76", "D78:J79", "D82:J85", "D88:J91", "D93:J94", "D96:J116",
      "A121:C121", "D122:J128", "D131:J134", "D136:J137", "D139:J143", "D146:J149", "D151:J152", "D154:J174",
      "A179:C179", "D180:J186", "D189:J192", "D194:J195", "D197:J201", "D204:J207", "D209:J210", "D212:I232",
      "A237:C237", "D238:J244", "D247:J250", "D252:J253", "D255:J259", "D262:J265", "D267:J268", "D270:J290"
    ];




    // Limpiar los rangos especificados en la copia
    rangesToClear.forEach(function (range) {
      backupSheet.getRange(range).clearContent();
    });

    Logger.log('Rangos especificados limpiados en la copia.');

    // Guardar la copia en Google Drive
    var folderId = '1lCHAW32Kdxv39aKRGtAigOw8zqA8tsP0'; // ID de la carpeta en Drive
    var folder = DriveApp.getFolderById(folderId);
    var file = DriveApp.getFileById(backupSpreadsheet.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // Elimina el archivo de la raíz de My Drive

    Logger.log(`Backup guardado en la carpeta con ID: ${folderId}`);
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}
