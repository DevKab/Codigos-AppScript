function ciclicosBoton(){//funciona 7/8/25
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var primerahoja =  libro.getSheets()[0];
  var nombreHoja = primerahoja.getName();

  if(nombreHoja === "S.Gastos CICLICOS INTERNO PS A0"){
    copiarCiclicosMaster("12gQU1l9FAKozTeGkhqdxbCUpFb-w_ddKUVDs3szNtSg", "S.Gastos CICLICOS INTERNO PS A0"); //1
  } else if(nombreHoja === "S.Gastos CICLICOS INTERNO PS A1"){
    copiarCiclicosMaster("12gQU1l9FAKozTeGkhqdxbCUpFb-w_ddKUVDs3szNtSg", "S.Gastos CICLICOS INTERNO PS A1");//2
  } else if(nombreHoja === "S.Gastos CICLICOS INTERNO PS A2"){
    copiarCiclicosMaster("1CalZsgEqEhWPJGloUGBSZwUXz9uPaVK2VfwbvRQuTms", "S.Gastos CICLICOS INTERNO PS A2"); //3
  } else if(nombreHoja === "S.Gastos CICLICOS INTERNO PS A3"){
    copiarCiclicosMaster("1PaQdKfVk51UiMNnKVS-Jo0YiDSs3mU-76_zweQy1M6c", "S.Gastos CICLICOS INTERNO PS A3");//4
  } else if(nombreHoja === "S.Gastos CICLICOS INTERNO PS A4"){
    copiarCiclicosMaster("1lOQ7p4H4pfqpADV5pDQBFKLv-_Jaf9aS6OBjwi0zGos", "S.Gastos CICLICOS INTERNO PS A4");//5
  } else if(nombreHoja === "S.Gastos CICLICOS INTERNO PS A5"){
    copiarCiclicosMaster("1ngHul195CohXo7eFB6lvDOhgNxAP9pwOgKnt27th8UI", "S.Gastos CICLICOS INTERNO PS A5");//6
  } else if(nombreHoja === "S.Gastos CICLICOS INTERNO PS A6"){
    copiarCiclicosMaster("1havjYfhnJ-Qe5DyDg0duLPAX7BN7veffhysscsG9jPc", "S.Gastos CICLICOS INTERNO PS A6");//7
  }
}

function copiarCiclicosMaster(link, nombreHoja) {//copiado y eliminado
  var libroOrigen = SpreadsheetApp.openById(link); // ciclicos
  var libroDestino = SpreadsheetApp.openById('1GuXfQkKrYbWzqHPXsKQj8hVZi1ndNQHLZZ9rytkafDU'); // Temporal

  var hojaOrigen = libroOrigen.getSheetByName(nombreHoja);
  var hojaDestino = libroDestino.getSheetByName("SOLICITUD GASTOS TEMPORAL - CONCATENADO");

  // Obtener la fecha actual formateada
  var today = new Date();
  var fomateoToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');

  // Obtener los valores de la hoja origen
  var datos = hojaOrigen.getRange("A:AP").getValues();// de A:AO a A:AP

  // Preparar un arreglo para las filas que cumplen las condiciones
  var filasParaPegar = [];


  for (var i = 0; i < datos.length; i++) {
    var dataFecha = datos[i][29]; // Columna AB (índice 27) //28 a 29 //fecha captura B

    // Validar si el dato en la columna AB es una fecha válida
    if (dataFecha instanceof Date && !isNaN(dataFecha.getTime())) {
      var fomateoFecha = Utilities.formatDate(dataFecha, Session.getScriptTimeZone(), 'dd/MM/yy');

      // Verificar si coincide con la fecha de hoy
      if (fomateoFecha === fomateoToday) { //27 a 28
        // Verificar condiciones en la columna Z (índice 26)
        if (datos[i][28] === "PAGADO Y COMPROBANTE EN CARPETA") {
          filasParaPegar.push(datos[i]); // Añadir fila para pegar
        }
      }
    }
  }

  // Pegar solo las columnas A:AJ de las filas que cumplen las condiciones en la hoja destino
  if (filasParaPegar.length > 0) {
      // Extraer solo las columnas A:AJ (índices 0 a 35)
      var filasA_AP = filasParaPegar.map(function(fila) {
          return fila.slice(0, 36);
      });
      var ultimaFilaDestino = hojaDestino.getLastRow();
      hojaDestino.getRange(ultimaFilaDestino + 1, 1, filasA_AP.length, filasA_AP[0].length)
          .setValues(filasA_AP);
      Logger.log(filasA_AP.length + " filas copiadas a la hoja destino (A:AP).");
  } else {
      Logger.log("No se encontraron filas que cumplan las condiciones para copiar.");
  }
}


