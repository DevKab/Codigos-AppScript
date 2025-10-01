function generarAnual() {//con la actualizacion//Anual
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS(Despacho)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");

  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  //var hojaDestino = libroOrigen.getSheetByName("hojaPrueba");
  
  var datos = hojaOrigen.getRange("A:AE").getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  var anioInicio = 2025, mesInicio = 8;  // septiembre 2025
  var anioFin = 2026, mesFin = 11;       // diciembre 2026


  var periodicidades = {
    "3ER LUNES DE JUNIO": function (anio) { return obtenerAnual(anio, 5, 14, 8); }, // Junio = 5
    "4TO LUNES DE ABRIL": function (anio) { return obtenerAnual(anio, 3, 21, 8); }, // Abril = 3
    "4TO LUNES DE AGOSTO": function (anio) { return obtenerAnual(anio, 7, 21, 8); }, // Agosto = 7
    "4TO LUNES DE DICIEMBRE": function (anio) { return obtenerAnual(anio, 11, 21, 8); }, // Diciembre = 11
    "4TO LUNES DE ENERO": function (anio) { return obtenerAnual(anio, 0, 21, 8); }, // Enero = 0
    "4TO LUNES DE FEBRERO": function (anio) { return obtenerAnual(anio, 1, 21, 8); }, // Febrero = 1
    "4TO LUNES DE JULIO": function (anio) { return obtenerAnual(anio, 6, 21, 8); }, // Julio = 6
    "4TO LUNES DE JUNIO": function (anio) { return obtenerAnual(anio, 5, 21, 8); }, // Junio = 5
    "4TO LUNES DE MARZO": function (anio) { return obtenerAnual(anio, 2, 21, 8); }, // Marzo = 2
    "4TO LUNES DE MAYO": function (anio) { return obtenerAnual(anio, 4, 21, 8); }, // Mayo = 4
    "4TO LUNES DE NOVIEMBRE": function (anio) { return obtenerAnual(anio, 10, 21, 8); }, // Noviembre = 10
    "4TO LUNES DE SEPTIEMBRE": function (anio) { return obtenerAnual(anio, 8, 21, 8); } // Septiembre = 8
  };

  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    // Tomar columnas C:AA (índices 2 a 26)
    var filaDatos = datos[i].slice(2, 28);

    // en vez de un solo año, recorre todos
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var fecha = funcion(anio);
      fecha = ajustarPorFestivoAnual(fecha);

      // Solo guardar si cae dentro del rango
      //if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 28)) {
      if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 29)) {

      
        var fechaFormateada = "'" + formatearFecha(fecha);  // <- apostrofe antes
        var nuevaFila = [fechaFormateada].concat(filaDatos);



        // asegura col AB = "NUEVO"
        //while (nuevaFila.length < 26) { 
        while (nuevaFila.length < 27) { 
          nuevaFila.push(""); 
        }
        nuevaFila.push("NUEVO");

        salida.push(nuevaFila);
        Logger.log("Periodicidad encontrada: " + periodicidad);
        Logger.log("Fecha generada: " + fecha);
      }
    }
  }

  // escribe a partir de col B
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida);
    
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
              .setNumberFormat("dd/MM/yyyy");
  }

}

function formatearFecha(fecha) {
  if (!fecha) return "";
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();

  // Asegura que día y mes tengan 2 dígitos
  var diaStr = (dia < 10 ? "0" : "") + dia;
  var mesStr = (mes < 10 ? "0" : "") + mes;

  return diaStr + "/" + mesStr + "/" + anio;
}


// Función para obtener el tercer lunes de un mes dado
function obtenerAnual(anio, mes, sumterCuar, dias) {
  var fecha = new Date(anio, mes, 1);//busca el primer dia del mes
  var diaSemana = fecha.getDay(); //sacamos el primer dia de la semana
  var diasHastaLunes = (dias - diaSemana) % 7; //cualcula cuantos dias hasta el lunes
  var tercerLunes = 1 + diasHastaLunes + sumterCuar; //suma 14 dias para, el tercer lunes del mes o 21 para cuato dia del mes
  return new Date(anio, mes, tercerLunes);
} 

//si cae en un dias festivo ara lo siguiente: 
//si cae lunes dias festivo que lo mueva para el martes de ese semana
//nuevo visto
function ajustarPorFestivoAnual(fecha) {
  var festivosFijos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  function obtenerTercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 Noviembre
    var primerDia = fecha.getDay(); // 0=Dom, 1=Lun...
    var primerLunes = primerDia === 1 ? 1 : (8 - primerDia);
    var tercerLunes = primerLunes + 14; // sumo 14 días (dos semanas más)
    return new Date(year, 10, tercerLunes);
  }

  function esFestivo(d) {
    // Revisa los festivos fijos
    var esFijo = festivosFijos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
    // Revisa si es tercer lunes de noviembre
    var tercerLunes = obtenerTercerLunesNoviembre(d.getFullYear());
    var esTercerLunesNov = d.getMonth() === 10 && d.getDate() === tercerLunes.getDate();
    return esFijo || esTercerLunesNov;
  }

  if (!esFestivo(fecha)) return fecha; // si no es festivo, regresa igual

  var dia = fecha.getDay(); // 0=Dom, 1=Lun...

  // Si es lunes festivo → mover al martes
  if (dia === 1) {
    var martes = new Date(fecha);
    martes.setDate(martes.getDate() + 1);
    if (martes.getDay() !== 0 && martes.getDay() !== 6 && !esFestivo(martes)) {
      return martes;
    }
  }

  return fecha; // si nada aplica, regresa la original
}

/////////////////////
function generarBimestrales() {
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS(Despacho)");
  var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");
  //var hojaDestino = libroOrigen.getSheetByName("hojaPrueba");
  

  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  
  

  var datos = hojaOrigen.getRange("A:AE").getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  var anioInicio = 2025, mesInicio = 8;  // septiembre 2025
  var anioFin = 2026, mesFin = 11;       // diciembre 2026

  // Meses bimestrales
  var tercerCuartoBiM = [0, 2, 4, 6, 8, 10]; //4TO LUNES / 3ER MIERCOLES DE ENE, MAR, MAY, JUL, SEP, NOV
  var tercerCuartoBiMF = [1, 3, 5, 7, 9, 11]; //4TO LUNES DE FEB / 3ER MIERCOLES, ABR, JUN, AGO, OCT, DIC

  var periodicidades = {

    "3ER MIERCOLES DE ENE, MAR, MAY, JUL, SEP, NOV": function (anio) { return obtenerBimestral(anio, tercerCuartoBiM, 14, 10); }, // Junio = 5
    "3ER MIERCOLES DE FEB, ABR, JUN, AGO, OCT, DIC": function (anio) { return obtenerBimestral(anio, tercerCuartoBiMF, 14, 10); }, // Junio = 5
    "4TO LUNES DE ENE, MAR, MAY, JUL, SEP, NOV": function (anio) { return obtenerBimestral(anio, tercerCuartoBiM, 21, 8); }, // Junio = 5
    "4TO LUNES DE FEB, ABR, JUN, AGO, OCT, DIC": function (anio) { return obtenerBimestral(anio, tercerCuartoBiMF, 21, 8); } // Junio = 5
  };

  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase();
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase();// gastos personales
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();//gastos Despacho
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    // Tomar columnas C:AA (índices 2 a 26)
    var filaDatos = datos[i].slice(2, 27);

    // recorrer años dentro del rango
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var fechas = funcion(anio); // arreglo de fechas bimestrales
      fechas = ajustarPorFestivoBime(fechas); // ajusta lunes/miércoles festivos


      fechas.forEach(function(fecha) {
        if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 28)) {
          var fechaFormateada = formatearFechaB(fecha);
          var nuevaFila = [fechaFormateada].concat(filaDatos);

          // asegura que llegue hasta col AB
          //while (nuevaFila.length < 26) {
          while (nuevaFila.length < 27) {
            nuevaFila.push("");
          }
          nuevaFila.push("NUEVO");

          salida.push(nuevaFila);
          Logger.log("Periodicidad encontrada: " + periodicidad);
          Logger.log("Fecha generada: " + fecha);
        }
      });
    }
  }


  // Escribe la salida en la hoja destino
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida); 
    //solo se cambio de 1 a 2 porque empieza la col 2 osea B
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
             .setNumberFormat("dd/MM/yyyy");
  }
}

function formatearFechaB(fecha) {
  if (!fecha) return "";
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();
  return dia + "/" + mes + "/" + anio;
}

// Esta función debe devolver un arreglo de fechas, una por cada mes del arreglo
function obtenerBimestral(anio, mesesArr, sumterCuar, dias) {
  var fechas = [];
  mesesArr.forEach(function (mes) {
    var fecha = obtenerBimes(anio, mes, sumterCuar, dias);
    fechas.push(fecha);
  });
  return fechas;
}

// Función para obtener el lunes/miércoles según parámetros
function obtenerBimes(anio, mes, sumterCuar, dias) {
  var fecha = new Date(anio, mes, 1);
  var diaSemana = fecha.getDay();
  var diasHastaDia = (dias - diaSemana) % 7;
  var diaFinal = 1 + diasHastaDia + sumterCuar;
  return new Date(anio, mes, diaFinal);
}

/*
Si cae en día festivo, ajustar la fecha:
Miercoles -> Martes de esa semana
Lunes -> Martes de esa semana
 los domingos y sabados
*/

//dias festivos
/*
si cae un dia festivo en miercoles que pase al mastes de esa semana
si cae lunes festivo que pase al mastes de esa semana
y si no pasa nada pasa
 */
/*function ajustarPorFestivoBime(fechas) {
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 10, dia: 20 }, // 20 Noviembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  function esFestivo(d) {
    return festivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  // Recorrer todas las fechas y ajustar
  return fechas.map(function(fecha) {
    if (!fecha) return fecha;

    if (!esFestivo(fecha)) return fecha; // si no es festivo, regresar igual

    var dia = fecha.getDay(); // 0=Dom, 1=Lun, 3=Mié

    // Miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    // Lunes festivo → martes
    if (dia === 1) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() + 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    return fecha; // si no aplica, regresar fecha original
  });
}*/
function ajustarPorFestivoBime(fechas) {
  var festivosFijos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  // --- calcular el 3er lunes de noviembre dinámico ---
  function obtenerTercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 Noviembre
    var primerDia = fecha.getDay();    // 0=Dom, 1=Lun...
    var primerLunes = primerDia === 1 ? 1 : (8 - primerDia);
    var tercerLunes = primerLunes + 14; // dos semanas más
    return new Date(year, 10, tercerLunes);
  }

  function esFestivo(d) {
    // primero los fijos
    var fijo = festivosFijos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
    // ahora el 3er lunes de noviembre
    var tercerLunes = obtenerTercerLunesNoviembre(d.getFullYear());
    var esTercerLunesNov = d.getMonth() === 10 && d.getDate() === tercerLunes.getDate();
    return fijo || esTercerLunesNov;
  }

  // recorrer todas las fechas
  return fechas.map(function(fecha) {
    if (!fecha) return fecha;
    if (!(fecha instanceof Date)) return fecha;

    if (!esFestivo(fecha)) return fecha; // si no es festivo, se queda igual

    var dia = fecha.getDay(); // 0=Dom, 1=Lun, 2=Mar, 3=Mié...

    // miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    // lunes festivo → martes
    if (dia === 1) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() + 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        return martes;
      }
    }

    return fecha; // si no aplica, devolver la original
  });
}

////
