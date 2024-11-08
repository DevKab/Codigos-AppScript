function onEdit(e) {
  var celdaActiva = e.range;
  var filaActiva = celdaActiva.getRow();
  var hojaActiva = celdaActiva.getSheet();

  var colDeseadaF = 4; // columna F
  var colDeseadaG = 7; // columna G
  var filaNodeseada = 1;
  var hojaDeseada = "PENDIENTES GRALES"; // cambie el nombre a la hoja que desee

  if (hojaActiva.getName() == hojaDeseada && filaActiva != filaNodeseada) {
    if (celdaActiva.getColumn() == colDeseadaF) {
      var email = Session.getActiveUser().getEmail();
      var targetSheet = hojaActiva.getParent().getSheetByName(hojaDeseada);
      targetSheet.getRange(filaActiva, 2).setValue(email); // insertar el correo en la columna J de la fila activa
    } else if (celdaActiva.getColumn() == colDeseadaG) {
      var valorCeldaG = celdaActiva.getValue();
      if (valorCeldaG.toLowerCase() == "listo") {
        var email = Session.getActiveUser().getEmail();
        var targetSheet = hojaActiva.getParent().getSheetByName(hojaDeseada);
        targetSheet.getRange(filaActiva, 11).setValue(email); // insertar el correo en la columna K de la fila activa
      }
    }
  }
}


function VloqueaUsuarios() {
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nombrehojadeseada = "PENDIENTES GRALES";
  
  var columnaBusqueda = 1;  // Cambia este valor si deseas buscar en otra columna
  var datos = hojaActiva.getRange(1, columnaBusqueda, hojaActiva.getMaxRows(), 1).getValues();
  var ultimaFila = 0;
  
  // Encuentra la última fila con datos en la columna especificada
  for (var i = datos.length - 1; i >= 0; i--) {
    if (datos[i][0] !== "") {
      ultimaFila = i + 1;
      break;
    }
  }

  var rangosABloquear = [
    { startCol: 1, endCol: 12 }  // A-L
    //{ startCol: 23, endCol:  24} 
  ];


  // Verificar si la hoja activa es la deseada
  if (hojaActiva.getName() == nombrehojadeseada) {
    // Iterar sobre cada rango y bloquear las columnas
    rangosABloquear.forEach(function (rango) {
      var startCol = rango.startCol;
      var endCol = rango.endCol;

      // Encontrar la última fila con datos en el rango de columnas
      var startRow = columnaBusqueda;
      var endRow = encontrarUltimaFilaConDatosEnRango(hojaActiva, startRow, ultimaFila,startCol, endCol);

      // Bloquear el rango especificado
      var rangeToProtect = hojaActiva.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
      var protection = rangeToProtect.protect().setDescription('Bloquear filas con datos');

      //Definir quieres pueden definir las filas bloquedas
      var editores = protection.getEditors();
      protection.removeEditors(editores);
    });
  }
}

function encontrarUltimaFilaConDatosEnRango(hoja, startRow, endRow, startCol, endCol) {
  var data = hoja.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] != "") {
        return i + startRow;
      }
    }
  }
  return startRow; // Devuelve startRow si no se encontraron datos
}//funciona v3


function bloquearUltimasFilasV10() { 
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nombreHoja = "PENDIENTES GRALES";

  var columnaBusqueda = 1;  // Cambia este valor si deseas buscar en otra columna
  var datos = hoja.getRange(1, columnaBusqueda, hoja.getMaxRows(), 1).getValues();
  var ultimaFila = 0;

  // Encuentra la última fila con datos en la columna especificada
  for (var i = datos.length - 1; i >= 0; i--) {
    if (datos[i][0] !== "") {
      ultimaFila = i + 1;
      break;
    }
  }

  var ultimaColumna = 12; // Limitar la protección hasta la columna L (columna 12)
  var totalFilas = hoja.getMaxRows();
  var primeraFilaVacia = ultimaFila + 1;

  if (hoja.getName() == nombreHoja) {
    // Proteger filas con datos
    if (ultimaFila > 0) {
      var rangoProtegido = hoja.getRange(1, 1, ultimaFila, ultimaColumna);
      var proteccion = rangoProtegido.protect().setDescription('Bloquear filas con datos');

      // Definir quiénes pueden editar las filas bloqueadas
      var editores = proteccion.getEditors();
      proteccion.removeEditors(editores);

      // Si quieres que solo el propietario del documento pueda editar las filas bloqueadas, descomenta la siguiente línea:
      // proteccion.addEditor(SpreadsheetApp.getActiveSpreadsheet().getOwner());

      Logger.log('Protección aplicada desde fila 1 hasta fila: ' + ultimaFila);
    }

    // Desproteger filas vacías desde la última fila con datos hasta la última fila de la hoja
    if (primeraFilaVacia <= totalFilas) {
      var datosFilasVacias = hoja.getRange(primeraFilaVacia, columnaBusqueda, totalFilas - primeraFilaVacia + 1, 1).getValues();
      var primeraFilaVaciaEncontrada = null;

      // Encontrar la primera fila vacía
      for (var i = 0; i < datosFilasVacias.length; i++) {
        if (datosFilasVacias[i][0] === "") {
          primeraFilaVaciaEncontrada = primeraFilaVacia + i;
          break;
        }
      }

      if (primeraFilaVaciaEncontrada !== null) {
        Logger.log('Protección revisada desde fila: ' + primeraFilaVaciaEncontrada + ' hasta la última fila de la hoja.');
      } else {
        Logger.log('No se encontraron filas vacías después de la última fila con datos.');
      }
    }
  }
  Logger.log('Protección de filas completada.');
}


function bloquearUltimasFilasV10v3() {
  Logger.log('Inicio de la función bloquearUltimasFilasV10');

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nombreHoja = "PENDIENTES GRALES";

  if (hoja.getName() != nombreHoja) {
    Logger.log('La hoja activa no es la hoja deseada.');
    return;
  }

  var columnaBusqueda = 1;  // Cambia este valor si deseas buscar en otra columna
  var datos = hoja.getRange(1, columnaBusqueda, hoja.getMaxRows(), 1).getValues();
  var ultimaFila = 0;

  // Encuentra la última fila con datos en la columna especificada
  for (var i = datos.length - 1; i >= 0; i--) {
    if (datos[i][0] !== "") {
      ultimaFila = i + 1;
      break;
    }
  }

  Logger.log('Última fila con datos: ' + ultimaFila);

  var ultimaColumna = 12; // Limitar la protección hasta la columna L (columna 12)
  var totalFilas = hoja.getMaxRows();
  var primeraFilaVacia = ultimaFila + 1;

  // Proteger filas con datos
  if (ultimaFila > 0) {
    var rangoProtegido = hoja.getRange(1, 1, ultimaFila, ultimaColumna);
    var proteccion = rangoProtegido.protect().setDescription('Bloquear filas con datos');

    // Definir quiénes pueden editar las filas bloqueadas
    var editores = proteccion.getEditors();
    proteccion.removeEditors(editores);

    // Si quieres que solo el propietario del documento pueda editar las filas bloqueadas, descomenta la siguiente línea:
   proteccion.addEditor(SpreadsheetApp.getActiveSpreadsheet().getOwner());

    Logger.log('Protección aplicada desde fila 1 hasta fila: ' + ultimaFila);
  } else {
    Logger.log('No hay filas con datos para proteger.');
  }

  // Desproteger filas vacías desde la última fila con datos hasta la última fila de la hoja
  if (primeraFilaVacia <= totalFilas) {
    var datosFilasVacias = hoja.getRange(primeraFilaVacia, columnaBusqueda, totalFilas - primeraFilaVacia + 1, 1).getValues();
    var primeraFilaVaciaEncontrada = null;

    // Encontrar la primera fila vacía
    for (var i = 0; i < datosFilasVacias.length; i++) {
      if (datosFilasVacias[i][0] === "") {
        primeraFilaVaciaEncontrada = primeraFilaVacia + i;
        break;
      }
    }

    if (primeraFilaVaciaEncontrada !== null) {
      Logger.log('Protección revisada desde fila: ' + primeraFilaVaciaEncontrada + ' hasta la última fila de la hoja.');
    } else {
      Logger.log('No se encontraron filas vacías después de la última fila con datos.');
    }
  }

  Logger.log('Protección de filas completada.');
}

