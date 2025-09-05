//P-PS-FT-007- 5-R. -REPORTE GASTOS_R0.1/150825 V17
function onOpen() { 
    var ui = SpreadsheetApp.getUi();
  var mensaje = "Recuerda que esta plantilla contiene listas anidadas y recibe información de otros archivos:"
    + "\n- 🚫 No agregar o quitar columnas y filas."
    + "\n- 🚫 No alterar fórmulas."
    + "\n- 🚫 No modificar la posición de las tablas o el rango."
    + "\n- ✅ Para un uso adecuado del archivo consulta tu instrucción de trabajo P-PS-IT-002_ SOLICITUD DE GASTOS DESPACHO DIRECCIÓN SOLICITANTE"
    + "\n- ☎︎ Contacta a 'Optimización' para realizar modificaciones. V16";

  ui.alert(mensaje);


    ui.createMenu('📅 | Backup')
    .addItem('1. Informacion del Temporal| 📄', 'ExtraerInfoTemp')
    .addItem('2. Backup del 5-R | 📁', 'allFunct')
    .addToUi();
}

function ExtraerInfoTemp(){
  ejemploFuncion()
}

function allFunct() {
  copiarArchivosG1(); //implementado 03/06/2024 
  copiarFormatoAGoogleDrive();
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
    copiarTemporalAlMaster();//copiado G1 al concentrado.
    var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();// Obtén la hoja de cálculo activa
    var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
    var nombreArchivo = hojaDeCalculo.getName();// Obtén el nombre de la hoja de cálculo
    var nuevaHojaDeCalculo = hojaDeCalculo.copy('[Nuevo Vacio] ' + nombreArchivo + currentDate); // Crea una nueva hoja de cálculo
    var idNuevoArchivo = nuevaHojaDeCalculo.getId();// Obtén la ID del archivo de la nueva hoja de cálculo
    var nuevoNombre = '[Nuevo Vacio] ' + nombreArchivo; // Cambia el nombre del archivo copiado
    DriveApp.getFileById(idNuevoArchivo).setName(nuevoNombre); // Puedes ajustar el nuevo nombre según tus necesidades
    var carpetaDestino = DriveApp.getFolderById('1yjigewfWWJTeOY2irxg8FOyVOc8OV6sI'); // Reemplaza 'ID_DE_LA_CARPETA' con la ID de la carpeta destino 10HcEi2RlaT1U_BwBsWzcEzW0JBJfQb8q //carpeta mia: 1NB8_H0vuuGaxXzn0n2Wi1FlBqDPYxc7e
    DriveApp.getFileById(idNuevoArchivo).moveTo(carpetaDestino); // Mueve el nuevo archivo a la carpeta de destino
    Logger.log('Copia de formato creada y guardada en la carpeta destino. Nombre del archivo: ' + nuevoNombre); // Registra el nombre del archivo en el registro

    /*g1 Y g1 FONDEO DE TARJETAS */
    var hojasDatosFT = [
      { origen: "ENTRECUENTAS G1", destino: "FONDEO DE TARJETAS", rango: "P93:U126", columnaInicio: 3 }//modificado 11/09/2024 C-H = COPIO
    ];

    hojasDatosFT.forEach(function (hoja) {
      var hojaOrigen = hojaDeCalculo.getSheetByName(hoja.origen);
      var hojaDestino = nuevaHojaDeCalculo.getSheetByName(hoja.destino);
      copiarYpegarDatos_FT12(hojaOrigen, hojaDestino, hoja.rango, hoja.columnaInicio);
    });



    limpiarCeldasEnHojas(nuevaHojaDeCalculo);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}


function limpiarCeldasEnHojas(nuevaHojaDeCalculo) {
  var hojas = [
    { nombre: "G1", rangos: ["D5:AM1536"] },
    {
      nombre: "ENTRECUENTAS G1", rangos: ["B4:H300","I3:N110", "P3:U50","P54:AG89", "P93:U126"]
    },
    {
      nombre: "HistorialEjecuciones", rangos: ["A1:E22"]
    }
  ];

  hojas.forEach(function (hoja) {
    var sheet = nuevaHojaDeCalculo.getSheetByName(hoja.nombre);
    hoja.rangos.forEach(function (rango) {
      sheet.getRange(rango).clearContent();
    });
  });
}


function copiarArchivosG1() { //saca a una copia de g2 y de ENTRECUENTAS ==funciona == 09/01/2025
  var hojaDeCalculo = SpreadsheetApp.getActiveSpreadsheet();
  var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  var nombreArchivo = hojaDeCalculo.getName();
  var hojasDatos = ["ENTRECUENTAS G1", "G1"];

  var carpetaBackup = DriveApp.getFolderById("1UCdPq3rYJQWbrkLTicxeEniBl59YuhFN");//id de la carpeta a depositar. //carpeta mia id:1kez8C5PfEDB4PHH0I6fEnMje-N76YCPX

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

/////////////////////////Temporar ////////////////
function ejemploFuncion() {//principal
  var ui = SpreadsheetApp.getUi();
  ui.alert("Función ejemploFuncion ejecutada correctamente.");
  try {
    // Lógica de tu función
    Logger.log("Ejecutando función ejemplo...");

    // Registro exitoso de la ejecución
    registrarEjecucion('ejemploFuncion', 'Éxito');
  } catch (error) {
    // Registro en caso de fallo
    registrarEjecucion('ejemploFuncion', 'Error: ' + error.message);
  }
}

function registrarEjecucion(funcionNombre, resultado) {
  var hojaHistorial = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HistorialEjecuciones'); // Obtén la hoja llamada 'HistorialEjecuciones'
  var ui = SpreadsheetApp.getUi(); //obtener la interfaz del usuario para mostrar alertas

  if (!hojaHistorial) {
    hojaHistorial = SpreadsheetApp.getActiveSpreadsheet().insertSheet('HistorialEjecuciones');
    hojaHistorial.appendRow(['Fecha', 'Hora', 'Función', 'Usuario', 'Resultado']);
  }

  var fechaActual = new Date();
  var fechaFormato = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Verificar si ya hay un registro de esta función en el día actual
  var datos = hojaHistorial.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var fechaEnHoja = datos[i][0]; // Fecha de la hoja
    // Si la fecha en la hoja no está formateada correctamente, intenta formatearla
    if (fechaEnHoja instanceof Date) {
      fechaEnHoja = Utilities.formatDate(fechaEnHoja, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    // Comparar la fecha formateada de la hoja con la fecha actual
    if (fechaEnHoja === fechaFormato && datos[i][2] === funcionNombre) {
      Logger.log("La función " + funcionNombre + " ya se ejecutó hoy.");
      ui.alert("La función '" + funcionNombre + "' ya ha sido registrada hoy."); // Mostrar alerta al usuario
      return; // Si ya existe un registro de esta función en el día actual, no hacemos nada
    }
  }

  // Si no existe un registro para hoy, se agrega uno nuevo
  var usuario = Session.getActiveUser().getEmail(); // Obtener el correo del usuario que ejecuta el script

  // Añadir un nuevo registro en la hoja de historial
  hojaHistorial.appendRow([
    fechaFormato, // Solo la fecha, no la hora
    Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'HH:mm:ss'), // La hora de ejecución
    funcionNombre,
    usuario || "Anónimo",
    resultado
  ]);
  
  copiarTemporarG1();
}
/////////////
function copiarTemporarG1() { //copia y elimina
  var libroOrigen = SpreadsheetApp.openById('18SOk6PCHpIxbL7oEfXK8MHnr8yzWGzJWNf_HYCmrmGk'); //temporal idOriginal=18SOk6PCHpIxbL7oEfXK8MHnr8yzWGzJWNf_HYCmrmGk
  var libroDestino = SpreadsheetApp.getActiveSpreadsheet(); 

  var hojaOrigen = libroOrigen.getSheetByName("SOLICITUD GASTOS TEMPORAL - CONCATENADO");
  var hojaDestino = libroDestino.getSheetByName("G1");

  var today = new Date();
  var fomateoToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');

 var datos = hojaOrigen.getRange("A:AJ").getValues(); // ✅ Ahora 34 columnas, en lugar de 33 //A:AH = A:AI

  var filasParaPegar = [];
  var filasParaEliminar = []; 

  for (var i = 0; i < datos.length; i++) {
    var dataFecha = datos[i][29]; //28 a 29 //fecha de pago

    if (dataFecha instanceof Date && !isNaN(dataFecha.getTime())) {
      var fomateoFecha = Utilities.formatDate(dataFecha, Session.getScriptTimeZone(), 'dd/MM/yy');

      if (fomateoFecha === fomateoToday) {
        if (datos[i][28] === "PAGADO" || datos[i][28] === "PAGADO Y COMPROBANTE EN CARPETA") { //27 a 28
          filasParaPegar.push(datos[i]); 
          filasParaEliminar.push(i + 1); 
        }
      }
    }
  }

  if (filasParaPegar.length > 0) {
    var inicioFila = 5;
    var inicioColumna = 4; // Columna D es la 4
    var maxFilas = 1536 - 5 + 1; 
    var maxColumnas = 36;  //34 a 35 == 36AM

    var filaFinalDestino = 1536; 

    var datosDestino = hojaDestino.getRange(inicioFila, inicioColumna, filaFinalDestino - inicioFila + 1, maxColumnas).getValues();

    var ultimaFilaDestino = inicioFila;

    for (var i = 0; i < datosDestino.length; i++) {
      var fila = datosDestino[i];
      if (fila.some(function (cell) { return cell !== "" && cell !== null; })) {
        ultimaFilaDestino = inicioFila + i + 1;
      }
    }

    var filaDestino = ultimaFilaDestino;

    if (filasParaPegar.length > maxFilas) {
      filasParaPegar = filasParaPegar.slice(0, maxFilas);
      Logger.log("Se truncaron los datos para ajustarse al rango permitido.");
    }

    // ✅ Ahora el número de columnas coincide con la hoja destino
    hojaDestino.getRange(filaDestino, inicioColumna, filasParaPegar.length, maxColumnas)
      .setValues(filasParaPegar);
    
    Logger.log(filasParaPegar.length + " filas copiadas en D5:AJ1536.");
  } else {
    Logger.log("No hay datos para copiar.");
  }

  if (filasParaEliminar.length > 0) {
    for (var j = filasParaEliminar.length - 1; j >= 0; j--) {
      hojaOrigen.deleteRow(filasParaEliminar[j]);
    }
    Logger.log(filasParaEliminar.length + " filas eliminadas.");
  } else {
    Logger.log("No hay datos para eliminar.");
  }
}

////
function copiarTemporalAlMaster() {//copiado y eliminado
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet(); // G1 = 5R
  var libroDestino = SpreadsheetApp.openById('1MMRqJ_9i-yUKYUxyPsiEx_NM5HH2Nr0-8oq4oqc5Ol8'); // Master idTesteoV1 = 1N12NZmKe0JjWuFVtww2C4Xm52E9XMZyRgx00vQXvRL0

  var hojaOrigen = libroOrigen.getSheetByName("G1");
  var hojaDestino = libroDestino.getSheetByName("ACUMULADO 2025");

  // Obtener la fecha actual formateada
  var today = new Date();
  var fomateoToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');

  // Obtener los valores de la hoja origen
  var datos = hojaOrigen.getRange("D5:AM1536").getValues();// de A:AO a A:AP

  // Preparar un arreglo para las filas que cumplen las condiciones
  var filasParaPegar = [];

  for (var i = 0; i < datos.length; i++) {
    var dataFecha = datos[i][29]; // Columna AB (índice 27) //28 a 29

    // Validar si el dato en la columna AB es una fecha válida
    if (dataFecha instanceof Date && !isNaN(dataFecha.getTime())) {
      var fomateoFecha = Utilities.formatDate(dataFecha, Session.getScriptTimeZone(), 'dd/MM/yy');

      // Verificar si coincide con la fecha de hoy
      if (fomateoFecha === fomateoToday) { //27 a 28
        // Verificar condiciones en la columna Z (índice 26)
        if (datos[i][28] === "PAGADO Y COMPROBANTE EN CARPETA" || datos[i][28] === "ALTA DE BENEFICIARIO" || datos[i][28] === "CANCELADO" || datos[i][28] === "EN PROCESO" || datos[i][28] === "PENDIENTE" || datos[i][28] === "RECHAZADO") {
          filasParaPegar.push(datos[i]); // Añadir fila para pegar
          //filasParaEliminar.push(i + 1); // Guardar el índice de la fila para eliminar (+1 porque es 1-based)
        }
      }
    }
  }

  // Pegar todas las filas que cumplen las condiciones en la hoja destino
  if (filasParaPegar.length > 0) {
    var ultimaFilaDestino = hojaDestino.getLastRow();
    hojaDestino.getRange(ultimaFilaDestino + 1, 1, filasParaPegar.length, filasParaPegar[0].length)
      .setValues(filasParaPegar);
    Logger.log(filasParaPegar.length + " filas copiadas a la hoja destino.");
  } else {
    Logger.log("No se encontraron filas que cumplan las condiciones para copiar.");
  }
}
