function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 | Pruebas menú')
    .addItem('Backup', 'allFunct')
    .addToUi();
}

function allFunct() {
  copiaraFolder();
}

function copiaraFolder() {
  copiarFormatoAGoogleDrive();
  //arrastrarFormulasV1(); //acomulados 29/05/2024
  copiarArchivosG123(); //implementado 03/06/2024
}


function copiaYpegarDatos_SD(hojaOrigen, hojaDestino, rangoOrigen, columnaInicio, rangoLetras) {
  var datos = hojaOrigen.getRange(rangoOrigen).getValues();
  let ultimaFila = encontrarUltimaFilaEnColumna(hojaDestino, rangoLetras);
  hojaDestino.getRange(ultimaFila + 1, columnaInicio, datos.length, datos[0].length).setValues(datos);
  //Logger.log("ultima fila:" + ultimaFila)
}

function encontrarUltimaFilaEnColumna(hojaDestino, rangoLetras) {
  var valores = hojaDestino.getRange(rangoLetras).getValues();
  var ultimaFila = 0;

  for (var i = valores.length - 1; i >= 0; i--) {
    if (valores[i].some(cell => cell !== "")) {
      ultimaFila = i + 1;
      break;
    }
  }

  return ultimaFila;
}



function copiarYpegarDatos_MK(hojaOrigen, hojaDestino, rangoOrigen, columnaInicio) {
  var datos = hojaOrigen.getRange(rangoOrigen).getValues();
  var ultimaFilaDestino = test(1, hojaDestino);
  hojaDestino.getRange(ultimaFilaDestino + 2, columnaInicio, datos.length, datos[0].length).setValues(datos);
}

function test(col, hoja) {
  const ultimaFila = hoja.getMaxRows();
  const rango = hoja.getRange(1, col, ultimaFila).getValues();
  for (let i = ultimaFila - 1; i > 0; i--) {
    if (rango[i][0]) {
      return i;
    }
  }
}

function copiarYpegarDatos_FT12(hojaOrigen, hojaDestino, rangoOrigen, columnaInicio) { //FUNCIONA 2:32
  // Obtener los datos desde la hoja de origen (D5:N407)
  var datos = hojaOrigen.getRange(rangoOrigen).getDisplayValues(); //correccion 29/04/2024

  // Encontrar la última fila con valores en la hoja de destino (columna C)
  var ultimaFilaDestino = hojaDestino.getLastRow();

  // Pegar los datos en la hoja de destino (C:M) después de la última fila con valores
  hojaDestino.getRange(ultimaFilaDestino + 1, columnaInicio, datos.length, datos[0].length).setValues(datos);
}

function copiarFormatoAGoogleDrive() {
  try {
    var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();// Obtén la hoja de cálculo activa
    var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
    var nombreArchivo = hojaDeCalculo.getName();// Obtén el nombre de la hoja de cálculo
    var nuevaHojaDeCalculo = hojaDeCalculo.copy('copia ' + nombreArchivo + currentDate); // Crea una nueva hoja de cálculo
    var idNuevoArchivo = nuevaHojaDeCalculo.getId();// Obtén la ID del archivo de la nueva hoja de cálculo
    var nuevoNombre = 'Copia de ' + nombreArchivo; // Cambia el nombre del archivo copiado
    DriveApp.getFileById(idNuevoArchivo).setName(nuevoNombre); // Puedes ajustar el nuevo nombre según tus necesidades
    var carpetaDestino = DriveApp.getFolderById('18wLy1lxnOMg15i00eYnbTs0gOj8ZxxAA'); // Reemplaza 'ID_DE_LA_CARPETA' con la ID de la carpeta destino
    DriveApp.getFileById(idNuevoArchivo).moveTo(carpetaDestino); // Mueve el nuevo archivo a la carpeta de destino
    Logger.log('Copia de formato creada y guardada en la carpeta destino. Nombre del archivo: ' + nuevoNombre); // Registra el nombre del archivo en el registro

    var hojasDatos = [
      { origen: "G3", destino: "MARCO", rango: "D42:O69", columnaInicio: 1 }, //D5:O32 //probar
      { origen: "G3", destino: "RSM", rango: "D72:O105", columnaInicio: 1 } //inicia en 5 era antes 1 //D35:O68 //funciona
    ];

    hojasDatos.forEach(function (hoja) {
      var hojaOrigen = hojaDeCalculo.getSheetByName(hoja.origen);
      var hojaDestino = nuevaHojaDeCalculo.getSheetByName(hoja.destino);
      copiarYpegarDatos_MK(hojaOrigen, hojaDestino, hoja.rango, hoja.columnaInicio);
    });

    /*g1 Y g1 FONDEO DE TARJETAS */
    var hojasDatosFT = [
      { origen: "G1", destino: "FONDEO DE TARJETAS", rango: "X120:AC150", columnaInicio: 3 } //funciona
    ];

    hojasDatosFT.forEach(function (hoja) {
      var hojaOrigen = hojaDeCalculo.getSheetByName(hoja.origen);
      var hojaDestino = nuevaHojaDeCalculo.getSheetByName(hoja.destino);
      copiarYpegarDatos_FT12(hojaOrigen, hojaDestino, hoja.rango, hoja.columnaInicio);
    });

    var hojaOrigenG1 = hojaDeCalculo.getSheetByName("G1");
    var libroDestino = SpreadsheetApp.openById("1yPuiDY905qyuMwO1G93RSIoISSLS6H6UaKkoJ5Vpnn8"); // id donde se va pegar archivo de cada año
    var hojaDestino = libroDestino.getSheetByName("DESPACHO");
    copiarYpegarDatos_Acumulado_despacho(hojaOrigenG1, hojaDestino, "D5:P286", 1, "A:M");


    var hojasSD = [
      { origen: "G1", destino: "SD", rango: "X5:Z56", columnaInicio: 2, rangoLetras: "B:D" },//funciona
      { origen: "G3", destino: "SD", rango: "Z5:AB55", columnaInicio: 6, rangoLetras: "F:H" },//por probar//no copia bien.

    ];

    hojasSD.forEach(function (hoja) {
      var hojaOrigen = hojaDeCalculo.getSheetByName(hoja.origen);
      var hojaDestino = nuevaHojaDeCalculo.getSheetByName(hoja.destino);
      copiaYpegarDatos_SD(hojaOrigen, hojaDestino, hoja.rango, hoja.columnaInicio, hoja.rangoLetras);
    });

    limpiarCeldasEnHojas(nuevaHojaDeCalculo);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}

function limpiarCeldasEnHojas(nuevaHojaDeCalculo) {
  var hojas = [
    {
      nombre: "G1", rangos: ["D5:P286", "R5:T28", "R31:T49", "R52:T63", "R66:T73", "R76:T106", "X5:Z56", "X63:AC117", "X120:AC150", "X154:Y162", "X164:Y168", "X170:Y177", "X179:Y183",
        "X185:Y189", "X191:Y195", "X197:Y204", "X206:Y210", "X212:Y216", "X218:Y222", "X224:Y231", "X233:Y237", "X239:Y243", "X245:Y249", "X251:Y255", "X257:Y261", "X263:Y267", "X269:Y273",
        "X275:Y279", "X281:Y285", "X287:Y291", "X293:Y297", "X299:Y303", "X305:Y309", "X311:Y315", "AC154:AD158", "AC160:AD164", "AC166:AD170", "AC172:AD176", "AC178:AD185", "AC187:AD191",
        "AC193:AD197", "AC199:AD206", "AC208:AD215", "AC217:AD221", "AC223:AD230", "AC232:AD236", "AC238:AD242", "AC244:AD248", "AC250:AD254", "AC256:AD258", "AC260:AD264", "AC266:AD270",
        "AC272:AD276", "AC278:AD285", "AC287:AD295", "AC297:AD305", "AC307:AD311", "AC313:AD321", "AC323:AD327", "AC329:AD333", "AC335:AD339", "AC341:AD345"]
    },
    {
      nombre: "G3", rangos: ["D72:O105", "D42:O69", "U5:W27", "U30:W48", "U51:W62", "U65:W72", "U75:W101", "Z5:AB55", "Z59:AA67", "Z69:AA73", "Z75:AA82", "Z84:AA88", "Z90:AA92", "Z94:AA99", "Z101:AA108", "Z110:AA114", "Z116:AA120", "Z122:AA126", "Z128:AA135", "Z137:AA141", "Z143:AA147", "Z149:AA153", "Z155:AA159", "Z161:AA165", "Z167:AA171", "Z173:AA177",
        "Z179:AA183", "Z185:AA189", "Z191:AA195", "Z197:AA201", "Z203:AA207", "Z209:AA213", "Z215:AA219", "AE59:AF63", "AE65:AF69", "AE71:AF75", "AE77:AF81", "AE83:AF90", "AE92:AF97", "AE99:AF103", "AE105:AF112", "AE114:AF121", "AE123:AF127", "AE129:AF136", "AE138:AF142", "AE144:AF148", "AE150:AF154", "AE156:AF160", "AE162:AF167", "AE169:AF173", "AE175:AF178", "AE180:AF184", "AE186:AF193", "AE195:AF203", "AE205:AF213", "AE215:AF219", "AE221:AF225", "AE227:AF231", "AE233:AF237", "AE239:AF243", "AE245:AF249"]
    }
  ];

  hojas.forEach(function (hoja) {
    var sheet = nuevaHojaDeCalculo.getSheetByName(hoja.nombre);
    hoja.rangos.forEach(function (rango) {
      sheet.getRange(rango).clearContent();
    });
  });
}



function copiarArchivosG123() { //saca a LOS tres g1, g2, g3.
  var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();
  var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  var nombreArchivo = hojaDeCalculo.getName();
  var hojasDatos = ["G1", "G3"]; //"G1", "G2", "G3"
  var carpetaBackup = DriveApp.getFolderById("18wLy1lxnOMg15i00eYnbTs0gOj8ZxxAA");

  hojasDatos.forEach(function (hojaNombre) {
    var hojaOrigen = hojaDeCalculo.getSheetByName(hojaNombre);
    if (!hojaOrigen) {
      Logger.log('No se encontró la hoja con el nombre: ' + hojaNombre);
      return;
    }

    var nombreBackup = hojaNombre + ' - ' + nombreArchivo + ' - ' + currentDate;
    var nuevaHojaDeCalculo = SpreadsheetApp.create(nombreBackup);
    var hojaNueva = hojaOrigen.copyTo(nuevaHojaDeCalculo);
    hojaNueva.setName(hojaNombre);

    // Eliminar la hoja inicial creada al momento de crear el nuevo archivo
    var hojaInicial = nuevaHojaDeCalculo.getSheetByName('Hoja 1');
    if (hojaInicial) {
      nuevaHojaDeCalculo.deleteSheet(hojaInicial);
    }

    var idNuevoArchivo = nuevaHojaDeCalculo.getId();
    DriveApp.getFileById(idNuevoArchivo).moveTo(carpetaBackup);
  });
}

