function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var mensaje = "Recuerda que esto es una plantilla automatizada:"
               + "\n- No agregar o quitar columnas y filas."
               + "\n- No alterar fórmulas."
               + "\n- No modificar la posición de las tablas o el rango."
               + "\n- Contacta a 'Optimización' para realizar modificaciones.";
  ui.alert('Mensaje de alerta', mensaje, 
  ui.ButtonSet.OK);
  ui.createMenu('Díario')
    .addItem('Crear Backup', 'btns')
    .addToUi();
}

// function mostrarVentanaEmergenteConImagen() {
//   var html = HtmlService.createHtmlOutput('<div style="text-align: center;">' +
//     '<img src="https://i.imgur.com/oSIxVZi.jpeg" style="width: 100%; max-width: 700px; height: auto;">' +
//     '<div style="margin-top: 20px;">' +
//     '<button style="width: 100%; padding: 10px; background-color: green; color: white; border: none; font-size: 16px; cursor: pointer;" onclick="google.script.host.close()">Aceptar</button>' +
//     '</div>' +
//     '</div>')
//     .setWidth(700)
//     .setHeight(800);
//   SpreadsheetApp.getUi().showModalDialog(html, '¡ALERTA!');
// }




function btns() {
  pruebascoloresporminuto()
  mostrarFilasColumnasOcultas()
  eliminarReglasFormatoCondicional()
  eliminarFiltroDeTabla()
  copiarArchivoASpecificFolder()
  bloquearHojasEspecificas()
  copiarYpegarDatos()
}

function mostrarFilasColumnasOcultas() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  // Mostrar filas ocultas
  for (var row = 1; row <= lastRow; row++) {
    if (sheet.isRowHiddenByUser(row)) {
      sheet.showRows(row);
    }
  }

  // Mostrar columnas ocultas
  for (var column = 1; column <= lastColumn; column++) {
    if (sheet.isColumnHiddenByUser(column)) {
      sheet.showColumns(column);
    }
  }
}

function eliminarReglasFormatoCondicional() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Obtén la hoja específica por nombre
  var sheet = spreadsheet.getSheetByName("test");

  if (!sheet) {
    Logger.log("La hoja especificada no existe: " + sheet)
    return;
  }

  var reglasFormatoCondicional = sheet.getConditionalFormatRules();

  // Crear una nueva colección de reglas sin las reglas de formato condicional
  var nuevasReglasFormatoCondicional = [];

  // Aplicar las nuevas reglas de formato condicional a la hoja
  sheet.setConditionalFormatRules(nuevasReglasFormatoCondicional);
}


function eliminarFiltroDeTabla() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("test"); // Cambia por el nombre de la hoja que contiene la tabla

  var startRow = 2; // Fila donde comienza la tabla
  var startColumn = 1; // Columna donde comienza la tabla
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  var rangeWithTable = sheet.getRange(startRow, startColumn, lastRow - startRow + 1, lastColumn - startColumn + 1);

  var filter = rangeWithTable.getFilter();
  if (filter !== null) {
    filter.remove();
  }
}

function pruebascoloresporminuto() {
  var sheetName = 'test'; // Reemplaza 'NombreDeTuHoja' con el nombre de la hoja en la que deseas aplicar el formato.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('La hoja especificada no se encontró en la hoja de cálculo.');
    return;
  }

  // Define los rangos específicos que deseas formatear
  var blueRanges = [
    'D6:D69', 'F6:F69', 'H6:H69', 'J6:J69', 'L6:L69', 'N6:N69', 'P6:P69'
  ];
  var whiteRanges = [
    'E6:E69', 'G6:G69', 'I6:I69', 'K6:K69', 'M6:M69', 'O6:O69', 'Q6:Q69'
  ];

  var blueBackgroundColor = '#c9daf8'; // Color de fondo azul
  var whiteBackgroundColor = '#FFFFFF'; // Color de fondo blanco

  formatRangeStyles(sheet, blueRanges, blueBackgroundColor);
  formatRangeStyles(sheet, whiteRanges, whiteBackgroundColor);
}

function formatRangeStyles(sheet, ranges, backgroundColor) {
  for (var i = 0; i < ranges.length; i++) {
    var range = sheet.getRange(ranges[i]);

    // Cambiar el color de fondo
    range.setBackground(backgroundColor);

    // Cambiar el color de las letras
    range.setFontColor('#000000'); // Cambiar a color negro (#000000)

    // Cambiar el tamaño de la fuente
    range.setFontSize(12); // Cambiar a tamaño de fuente 12
  }
}

function copiarArchivoASpecificFolder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetId = sheet.getId();
  var file = DriveApp.getFileById(sheetId);
  var folders = file.getParents(); // Obtiene todas las carpetas en las que se encuentra el archivo

  if (folders.hasNext()) {
    var folder = folders.next(); // Obtiene la primera carpeta en la que se encuentra el archivo
    var folderId = folder.getId();
    var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
    var newName = sheet.getName() + " - " + currentDate;
    var copy = file.makeCopy(newName, DriveApp.getFolderById(folderId));
    var newFileId = copy.getId();
    var newSpreadsheet = SpreadsheetApp.openById(newFileId);
    var newSheet = newSpreadsheet.getSheetByName("test");

    var rangesToClear = [
      "D6:Q69", "S35:T70", "Y5:AR92", "E80:F155", "G80:H154", "I80:J155", "K80:L154", "N79:O103", "P79:Q102", "N105:O129", "P105:Q128", "N131:O155", "P131:Q154", "N157:O181", "P157:Q180",
      "N183:O207", "P183:Q206", "N209:O233", "P209:Q232", "N235:O259", "P235:Q258", "N261:O285", "P261:Q284", "N287:O311", "P287:Q310", "N313:O337", "P313:Q336", "N339:O363", "P339:Q362", "S79:W86", "S89:W96", "S99:W106", "S109:W116", "S119:W126", "S129:W136", "S139:W146", "S149:W156", "S159:W166", "S169:W176", "S179:W186", "S189:W196", "S199:W206", "S209:W216", "S219:W226", "S229:W236", "S239:W246", "S249:W256", "S259:W266", "S269:W276", "S279:W286", "S289:W296", "S299:W306", "S309:W316", "S319:W326", "S329:W336", "V339:W346", "Y100:AO187", "AE195:AI282", "AN195:AO282", "Y290:AI377", "Y194:Z282", "AB195:AC234", "AB239:AC282", "AK195:AL214", "AK219:AL238", "AK243:AL262", "AK267:AL286", "AK291:AL310", "AK315:AL334", "E157:L231", "E234:L308", "E311:L385", "E388:L462", "AK339:AL358", "E465:F540", "G465:L539"
    ];

    rangesToClear.forEach(function (range) {
      newSheet.getRange(range).clearContent();
    });

    var secondSheet = newSpreadsheet.getSheetByName("1"); // Reemplaza "segundaHoja" con el nombre de la segunda hoja

    var rangesToClearInSecondSheet = [
      "D2:D150", "F2:J150", "M2:M150", "Q2:W150", "D153:D193", "F153:J193", "M153:M193", "Q153:W193"];

    rangesToClearInSecondSheet.forEach(function (range) {
      secondSheet.getRange(range).clearContent();
    });

    var originalSheet = sheet.getSheetByName("test");
    var dataToCopy = originalSheet.getRange("D70:Q70").getValues();

    newSheet.getRange("D5").offset(0, 0, dataToCopy.length, dataToCopy[0].length).setValues(dataToCopy);
  } else {
    Logger.log("El archivo no se encuentra en ninguna carpeta.");
  }
}


//bloqueo estapa 2 y copiado de info. concentrado -------------

function bloquearHojasEspecificas() {
  var hojasABloquear = ["1", "test"]; // Nombres de las hojas a bloquear
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  for (var i = 0; i < hojasABloquear.length; i++) {
    var hoja = spreadsheet.getSheetByName(hojasABloquear[i]);

    if (hoja) {
      var protection = hoja.protect().setDescription('Bloqueo automático'); // Protege la hoja

      // Eliminar a todos los editores actuales de la hoja protegida
      var editoresActuales = protection.getEditors();
      for (var j = 0; j < editoresActuales.length; j++) {
        protection.removeEditor(editoresActuales[j]);
      }
    } else {
      Logger.log("No se encontró la hoja " + hojasABloquear[i]);
    }
  }
}

function copiarYpegarDatos() {
  var archivoOrigen = SpreadsheetApp.getActiveSpreadsheet(); // Archivo actual
  var hojaOrigen = archivoOrigen.getSheetByName("test"); // Hoja de origen
  var datos = hojaOrigen.getRange("E701:L709").getValues(); // Obtener datos del rango

  var archivoDestinoId = "1sWV2HXIcrfaJWOZWRqW_xyVwtRamidQHfJBJpzexbdg"; // ID del archivo de destino
  var archivoDestino = SpreadsheetApp.openById(archivoDestinoId); // Abrir el archivo de destino
  var hojaDestino = archivoDestino.getSheetByName("CONCENTRADO2024"); // Hoja de destino
  var ultimaFila = hojaDestino.getLastRow() + 1;
  var fechaActual = new Date();

  var nombreArchivoOrigen = archivoOrigen.getName();
  for (var i = 0; i < datos.length; i++) {
    var filaDestino = ultimaFila + i;
    hojaDestino.getRange(filaDestino, 1, 1, datos[0].length).setValues([datos[i]]);
    hojaDestino.getRange(filaDestino, 9).setValue(fechaActual);
    hojaDestino.getRange(filaDestino, 10).setValue(nombreArchivoOrigen);
  }
}

