//boton que esta con un activador
function ciclicosBoton(){//10/10/2025
  try {
    var hojasDatos = [
      { link: "1havjYfhnJ-Qe5DyDg0duLPAX7BN7veffhysscsG9jPc", nombreHojaD: "S.Gastos CICLICOS INTERNO PS A6", nombreHojaO: "Base de Datos Despacho"},
      { link: "1ngHul195CohXo7eFB6lvDOhgNxAP9pwOgKnt27th8UI", nombreHojaD: "S.Gastos CICLICOS INTERNO PS A5", nombreHojaO: "Base de Datos Despacho"},//para despacho  A5
      { link: "1lOQ7p4H4pfqpADV5pDQBFKLv-_Jaf9aS6OBjwi0zGos", nombreHojaD: "S.Gastos CICLICOS INTERNO PS A4", nombreHojaO: "Base de Datos Despacho"},
      { link: "1PaQdKfVk51UiMNnKVS-Jo0YiDSs3mU-76_zweQy1M6c", nombreHojaD: "S.Gastos CICLICOS INTERNO PS A3", nombreHojaO: "Base de Datos Despacho"},
      { link: "1CalZsgEqEhWPJGloUGBSZwUXz9uPaVK2VfwbvRQuTms", nombreHojaD: "S.Gastos CICLICOS INTERNO PS A2", nombreHojaO: "Base de Datos Despacho"},// es nominas.
      { link: "1HueYpVVHTSL6bJpBF8y_PEF-C5MGIt_o2wgPeYRJd7I", nombreHojaD: "S.Gastos CICLICOS INTERNO PS A1", nombreHojaO: "Base de Datos Personal"},//para personal A1
      { link: "1HueYpVVHTSL6bJpBF8y_PEF-C5MGIt_o2wgPeYRJd7I", nombreHojaD: "S.Gastos CICLICOS INTERNO PS A1", nombreHojaO: "Base de Datos Despacho"},//para Despacho A1
      { link: "18S6lqUMLJ07QB4QEWvh6Gppb7PSjEWnXPbhY57sbZhM", nombreHojaD: "S.Gastos CICLICOS INTERNO PS A0", nombreHojaO: "Base de Datos Despacho"}
    ];

    hojasDatos.forEach(function (hoja) {
      try {
        envioInfoCiclico(hoja.link, hoja.nombreHojaD, hoja.nombreHojaO);
      } catch (error) {
        Logger.log(`Error procesando hoja con link ${hoja.link} y nombre Origen ${hoja.nombreHojaD} y nombre de la hoja Destino ${hoja.nombreHojaO}: ${error.message}`);
      }
    });
  } catch (error) {
    Logger.log(`Error general en limk12Archivos: ${error.message}`);
  }
}

function envioInfoCiclico(link, nombreDestino, nombreOrigen) {
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName(nombreOrigen);
  //var hojaOrigen = libroOrigen.getSheetByName("Planeador Despacho");

  var datos = hojaOrigen.getRange("B:AC").getValues();

  var libroDestino = SpreadsheetApp.openById(link);
  var hojaDestino = libroDestino.getSheetByName(nombreDestino);

  var fecha = new Date();
  var fechaFormateada = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yy');

  var filasPegar = [];

  for (var i = 1; i < datos.length; i++) {
    var fechaOrigen = datos[i][0];
    if (fechaOrigen instanceof Date && !isNaN(fechaOrigen.getTime())) {
      var fechaFormateadaOrigen = Utilities.formatDate(fechaOrigen, Session.getScriptTimeZone(), 'dd/MM/yy');
      if (fechaFormateadaOrigen === fechaFormateada) {
        filasPegar.push(datos[i]);
      }
    }
  }

  var valoresColB = hojaDestino.getRange("B:B").getValues();
  var ultimaFilaColumnaB = 0;
  for (var i = valoresColB.length - 1; i >= 0; i--) {
    if (valoresColB[i][0] !== "" && valoresColB[i][0] !== null) {
      ultimaFilaColumnaB = i + 1;
      break;
    }
  }

  if (filasPegar.length > 0) {
    var filasFiltradas;

    // Dependiendo del nombre de la hoja aplicamos el filtro correcto
    if (nombreDestino === "S.Gastos CICLICOS INTERNO PS A0") {
      filasFiltradas = filasPegar.filter(function (filaArea) {
        return ["NATALIE_REYNA"].includes(filaArea[1]);
      });
    } 
    else if (nombreDestino === "S.Gastos CICLICOS INTERNO PS A1") {
      filasFiltradas = filasPegar.filter(function (filaArea) {
        return ["VALERIA_VARGAS"].includes(filaArea[1]);
      });
    } 
     else if ("S.Gastos CICLICOS INTERNO PS A2" === "S.Gastos CICLICOS INTERNO PS A2") {
      filasFiltradas = filasPegar.filter(function (filaArea) {
        //return ["VALERIA_VARGAS"].includes(filaArea[1]);
        return ["VALERIA_VARGAS"].includes(filaArea[1]) && ["DESPACHO"].includes(filaArea[3]);
      });
    }
    else if (nombreDestino === "S.Gastos CICLICOS INTERNO PS A3") {
      filasFiltradas = filasPegar.filter(function (filaArea) {
        return ["FATIMA_MARTINEZ"].includes(filaArea[1]);
      });
    } 
    else if (nombreDestino === "S.Gastos CICLICOS INTERNO PS A4") {
      filasFiltradas = filasPegar.filter(function (filaArea) {
        return ["NADIA_ELIZONDO"].includes(filaArea[1]);
      });
    } 
    else if (nombreDestino === "S.Gastos CICLICOS INTERNO PS A5") { //Seria gastos personales??
      filasFiltradas = filasPegar.filter(function (filaArea) {
        return ["NAYELI_LUNA"].includes(filaArea[1]);
      });
    } 
    else if (nombreDestino === "S.Gastos CICLICOS INTERNO PS A6") {
      filasFiltradas = filasPegar.filter(function (filaArea) {
        return ["FRIDA_PIÑA"].includes(filaArea[1]);
      });
    } else {
      filasFiltradas = filasPegar; // si no hay regla definida, manda todo
    }

    var valorG = filasFiltradas.map(function (filaArea) {
      return filaArea.slice(0, 28);
    });

    if (valorG.length > 0) {
      hojaDestino.getRange(ultimaFilaColumnaB + 1, 2, valorG.length, valorG[0].length).setValues(valorG);
      Logger.log("resultado: " + valorG.length+ " nombre de la hoja Des: " + nombreDestino);
    } else {
      Logger.log("no hay filas que cumplan la condición.");
    }
  }
}

/*codigo para sacar anual */
function generarAnual() {//con la actualizacion
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");

  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  var hojaDestino = libroOrigen.getSheetByName("planeador Despacho");
  
  var datos = hojaOrigen.getRange("A:AE").getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  var anioInicio = 2025, mesInicio = 8;  // septiembre 2025
  var anioFin = 2026, mesFin = 11;       // diciembre 2026


  var periodicidades = {
    "3ER LUNES DE JUNIO": function (anio) { return obtenerAnual(anio, 5, 14, 8); }, // Junio = 5
    "3ER MIERCOLES DE JULIO": function (anio) { return obtenerAnual(anio, 6, 14, 10); }, // Julio = 6
    //"3ER MIERCOLES DE JULIO": function (anio) { return obtenerAnual(anio, 6, 14, 8); }, // Julio = 6
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
      if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 28)) {
      //if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 29)) {

      
        var fechaFormateada = "'" + formatearFecha(fecha);  // <- apostrofe antes
        var nuevaFila = [fechaFormateada].concat(filaDatos);



        // asegura col AB = "NUEVO"
        while (nuevaFila.length < 26) { 
        //while (nuevaFila.length < 27) { 
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
                    //anio, 6, 14, 8
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

/*bimestral */
function generarBimestrales() {
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
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

/*eliminar para quitar informacion inesesaria */
function eliminarFilasDespacho() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("planeador Despacho");
  if (!hoja) {
    Logger.log("La hoja 'despacho' no existe.");
    return;
  }
  var numFilas = hoja.getLastRow();
  if (numFilas > 0) {
    hoja.deleteRows(2, numFilas - 1); // Elimina desde la fila 2 hasta la última (deja encabezados)
  }
}

/*mensual */
function mensual() { //funciona bien = no tocar
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");
  var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");

  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  

  var ultimaFilaOrigen = hojaOrigen.getLastRow();
  var ultimaColumnaOrigen = hojaOrigen.getLastColumn();
  var datos = hojaOrigen.getRange(1, 1, ultimaFilaOrigen, ultimaColumnaOrigen).getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  var anioInicio = 2025, mesInicio = 8;  // septiembre
  var anioFin = 2026, mesFin = 11;       // diciembre

  var periodicidades = {
    "1ER DIA HABIL DEL MES": primerDiaHabilDelMes,
    "1ER LUNES DE CADA MES": primerLunesDelMes,
    "1ER MIERCOLES DE CADA MES": primerMiercolesDelMes,
    "2DO LUNES DE CADA MES": segundoLunesDelMes,
    "2DO MIERCOLES DE CADA MES": segundoMiercolesDelMes,
    "3ER LUNES DE CADA MES": tercerLunesDelMes,
    "3ER MIERCOLES DE CADA MES": tercerMiercolesDelMes,
    "4TO LUNES DE CADA MES": cuartoLunesDelMes,
    "4TO VIERNES DE CADA MES": cuartoViernesDelMes,
    "ULTIMO VIERNES DEL MES": ultimoViernesDelMes
  };

  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase(); //gastos personales
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();//gastos despacho
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    //var filaDatos = datos[i].slice(2, 30); // columnas C:AD
    var filaDatos = 
    datos[i].slice(2, 27); // columnas C:AD

    // Reemplazar AB por "NUEVO" si está vacía o tiene cualquier valor
    filaDatos[26] = "NUEVO"; // posición 25 = columna AB

    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var mesInicial = (anio === anioInicio ? mesInicio : 0);
      var mesFinal = (anio === anioFin ? mesFin : 11);

      for (var mes = mesInicial; mes <= mesFinal; mes++) {
        var fecha = funcion(anio, mes);
        fecha = ajustarPorFestivo(fecha, periodicidad);

        var fechaFormateada = formatearFechaM(fecha);

        // Insertar la fecha en columna B
        var nuevaFila = [ , fechaFormateada].concat(filaDatos); // A vacío, B=fecha, C:AE = filaDatos
        salida.push(nuevaFila);
      }
    }
  }

  if (salida.length > 0) {
   // hojaDestino.getRange(ultimaFilaDestino + 1, 1, salida.length, salida[0].length).setValues(salida);
    hojaDestino.getRange(ultimaFilaDestino + 1, 1, salida.length, salida[0].length).setValues(salida);
  }
}



//
// -------- FUNCIONES DE PERIODICIDAD --------
//
function primerDiaHabilDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() === 0 || fecha.getDay() === 6) {
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function primerLunesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() !== 1) {
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function primerMiercolesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() !== 3) {
    fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function segundoLunesDelMes(anio, mes) {
  var fecha = primerLunesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 7);
  return fecha;
}

function segundoMiercolesDelMes(anio, mes) {
  var fecha = primerMiercolesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 7);
  return fecha;
}

function tercerLunesDelMes(anio, mes) {
  var fecha = primerLunesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 14);
  return fecha;
}

function tercerMiercolesDelMes(anio, mes) {
  var fecha = primerMiercolesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 14);
  return fecha;
}

function cuartoLunesDelMes(anio, mes) {
  var fecha = primerLunesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 21);
  return fecha;
}

function cuartoViernesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  var count = 0;
  while (count < 4) {
    if (fecha.getDay() === 5) count++;
    if (count < 4) fecha.setDate(fecha.getDate() + 1);
  }
  return fecha;
}

function ultimoViernesDelMes(anio, mes) {
  var fecha = new Date(anio, mes + 1, 0);
  while (fecha.getDay() !== 5) {
    fecha.setDate(fecha.getDate() - 1);
  }
  return fecha;
}

//
// --------- Ajuste por días festivos ---------
//si cae en un dias festivo ara lo siguiente: 
//si es miércoles el dia festivo, mover al martes de ese semana
//cuando sea esta periosidad 1ER DIA HABIL DEL MES y caiga en dia festivo solo mover un dia habil de esta misma semana, si no mover a la siguiente semana pero primer dia habil de esa semana
//si caae un dia lunes el dia festivo mover al dia martes
//si cae un viernes mover al miercoles de la semana A sepcion la periosidad, porque si tiene periosidad 1ER DIA HABIL DEL MES:
//buscar el primer dia habil, ejemplo dia destivo 01/05/82026 es viernes pero su periosidad trae 1ER DIA HABIL DEL MES, si es asi debe buscsra el primer dia habil como 04/05/26
function ajustarPorFestivo(fecha, periodicidad) {
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  function tercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 de noviembre
    var primerDia = fecha.getDay();
    var primerLunes = (primerDia === 1) ? 1 : ((8 - primerDia) % 7) + 1;
    return { mes: 10, dia: primerLunes + 14 }; // tercer lunes
  }

  // Agregar el tercer lunes de noviembre del año de la fecha
  festivos.push(tercerLunesNoviembre(fecha.getFullYear()));

  function esFestivo(d) {
    return festivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  if (!esFestivo(fecha)) return fecha;

  var dia = fecha.getDay();

  // Miércoles → martes
  if (dia === 3) {
    var martes = new Date(fecha);
    martes.setDate(fecha.getDate() - 1);
    if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) return martes;
  }

  // Lunes → martes
  if (dia === 1) {
    var martes = new Date(fecha);
    martes.setDate(fecha.getDate() + 1);
    if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) return martes;
  }

  // Viernes → miércoles (o reglas de "1ER DIA HABIL DEL MES")
  if (dia === 5) {
    if (periodicidad === "1ER DIA HABIL DEL MES") {
      var siguiente = new Date(fecha);
      do {
        siguiente.setDate(siguiente.getDate() + 1);
      } while ((siguiente.getDay() === 0 || siguiente.getDay() === 6 || esFestivo(siguiente)) &&
               siguiente.getMonth() === fecha.getMonth());

      if (siguiente.getMonth() !== fecha.getMonth()) {
        siguiente = new Date(fecha.getFullYear(), fecha.getMonth() + 1, 1);
        while (siguiente.getDay() === 0 || siguiente.getDay() === 6 || esFestivo(siguiente)) {
          siguiente.setDate(siguiente.getDate() + 1);
        }
      }
      return siguiente;
    } else {
      var miercoles = new Date(fecha);
      miercoles.setDate(fecha.getDate() - 2);
      if (!esFestivo(miercoles) && miercoles.getDay() !== 0 && miercoles.getDay() !== 6) return miercoles;
    }
  }

  // Reglas de "1ER DIA HABIL DEL MES"
  if (periodicidad === "1ER DIA HABIL DEL MES") {
    var siguiente = new Date(fecha);
    var intentos = 0;
    do {
      siguiente.setDate(siguiente.getDate() + 1);
      intentos++;
      if (siguiente.getDay() === 1 && intentos > 1) break;
    } while ((siguiente.getDay() === 0 || siguiente.getDay() === 6 || esFestivo(siguiente)) && intentos < 7);

    if (siguiente.getDay() === 0 || siguiente.getDay() === 6 || esFestivo(siguiente)) {
      var proxLunes = new Date(fecha);
      proxLunes.setDate(proxLunes.getDate() + (8 - proxLunes.getDay()) % 7);
      for (var i = 0; i < 5; i++) {
        var diaHabil = new Date(proxLunes);
        diaHabil.setDate(diaHabil.getDate() + i);
        if (diaHabil.getDay() !== 0 && diaHabil.getDay() !== 6 && !esFestivo(diaHabil)) return diaHabil;
      }
    } else {
      return siguiente;
    }
  }

  return fecha;
}



//
// ---- Formato de fecha
//
function formatearFechaM(fecha) {
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/*quincenal */
/*
  codigo que de la hoja horigen "S.Gastos CICLICOS INTERNO PS" sacaras la periosdad cuando sea en la columna 29 === DIAS 15 Y DIAS 30 y cuando encuentes este cadena vas a sacar la fecha de cada mes 15 dias: osea en enero 15 miercoles y 30 jueves pero adcesion de fedrero la primera quincena va ser el 15 y (si el 15 sale sabado o domingo moverlo al viernes de preferencia) como en febrero no hay 30 hay que moverlo al ultimo dia vigente del mes (viegente de dias lunes a viernes). y esas fechas se pegaran en la columna A de la hojas destino : "Fechas 2025" y despes de la columna A ira el rango de C:AA de la hoja origen corespondiente perioridad que salga "DIAS 15 Y DIAS 30", el rango a sacar de la hoja destino es "A:AD".
*/
function quincenal() {//funciona no mover
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libro.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  var hojaDestino = libro.getSheetByName("Planeador Despacho");
  //var hojaDestino = libro.getSheetByName("Base");

  //var hojaOrigen = libro.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libro.getSheetByName("Planeador Personal");
  //var hojaDestino = libro.getSheetByName("hojaPrueba");

  var datos = hojaOrigen.getRange("A:AE").getValues(); // Incluye todo hasta col 30
  var salida = [];

  // Rango de fechas personalizable
  var anioInicio = 2025, mesInicio = 8;  // Septiembre 2025 (0=Enero)
  var anioFin = 2026, mesFin = 11;       // Diciembre 2026

  for (var i = 1; i < datos.length; i++) { // desde la fila 2
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();

    if (periodicidad === "DIAS 15 Y DIAS 30") {
      //var filaDatos = datos[i].slice(2, 27); // columnas C:AA
      var filaDatos = datos[i].slice(2, 28); // columnas C:AA

      for (var anio = anioInicio; anio <= anioFin; anio++) {
        var mesStart = (anio === anioInicio) ? mesInicio : 0;
        var mesEnd = (anio === anioFin) ? mesFin : 11;

        for (var mes = mesStart; mes <= mesEnd; mes++) {
          // ---- Primera quincena ----
          var fecha1 = ajustarSiFinDeSemana(new Date(anio, mes, 15));
          var nuevaFila1 = [formatearFechaQ(fecha1)].concat(filaDatos);
          nuevaFila1 = ajustarPorFestivoQuin(nuevaFila1); // ajusta lunes/miércoles festivos


          // Completa hasta col AB y agrega "NUEVO"
          //while (nuevaFila1.length < 26) nuevaFila1.push("");
          while (nuevaFila1.length < 26) nuevaFila1.push("");
          nuevaFila1.push("NUEVO");
          salida.push(nuevaFila1);

          // ---- Segunda quincena ----
          var fecha2;
          if (mes === 1) { // Febrero
            fecha2 = ultimoDiaHabilDelMes(anio, mes);
          } else {
            fecha2 = ajustarSiFinDeSemana(new Date(anio, mes, 30));
          }

          var nuevaFila2 = [formatearFechaQ(fecha2)].concat(filaDatos);
          nuevaFila2 = ajustarPorFestivoQuin(nuevaFila2); // ajusta lunes/miércoles festivos
          //while (nuevaFila2.length < 26) nuevaFila2.push("");
          while (nuevaFila2.length < 26) nuevaFila2.push("");
          nuevaFila2.push("NUEVO");
          salida.push(nuevaFila2);// salida agregar mes que salga con numero

          Logger.log("Periodicidad encontrada: " + periodicidad);
          Logger.log("Fechas generadas: " + fecha1 + ", " + fecha2);
        }
      }
    }
  }

  // Pegar en hoja destino desde columna B
 // Pegar en hoja destino desde columna B
  if (salida.length > 0) {
    var ultimaFilaDestino = hojaDestino.getLastRow();
    var numCols = salida[0].length; // columnas reales en la fila
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, numCols).setValues(salida);

    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
              .setNumberFormat("dd/MM/yyyy");
  }

}

function ajustarSiFinDeSemana(fecha) {
  var diaSemana = fecha.getDay(); // 0=Domingo, 6=Sábado
  if (diaSemana === 0) fecha.setDate(fecha.getDate() - 2);
  else if (diaSemana === 6) fecha.setDate(fecha.getDate() - 1);
  return fecha;
}

function ultimoDiaHabilDelMes(anio, mes) {
  var fecha = new Date(anio, mes + 1, 0); // último día del mes
  while (fecha.getDay() === 0 || fecha.getDay() === 6) {
    fecha.setDate(fecha.getDate() - 1);
  }
  return fecha;
}

function formatearFechaQ(fecha) {
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();
  return dia + "/" + mes + "/" + anio;
}

/*function ajustarPorFestivoQuin(fila) {
  // fila[0] es la fecha en formato "d/m/yyyy"
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
    if (!(d instanceof Date)) {
      var partes = d.split("/");
      d = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
    }
    return festivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  // Solo ajusta la fecha de la columna 0
  var partes = fila[0].split("/");
  var fecha = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));

  if (esFestivo(fecha)) {
    var dia = fecha.getDay();
    // Miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        fila[0] = formatearFechaQ(martes);
        return fila;
      }
    }
    // Lunes festivo → martes
    if (dia === 1) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() + 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        fila[0] = formatearFechaQ(martes);
        return fila;
      }
    }
  }
  // Si no aplica ajuste, regresa igual
  return fila;
}*/

function ajustarPorFestivoQuin(fila) {
  // fila[0] es la fecha en formato "d/m/yyyy"
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  // Calcular tercer lunes de noviembre para un año dado
  function tercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 de noviembre
    var primerDia = fecha.getDay();    // 0=Dom,...,6=Sab
    // calcular primer lunes
    var primerLunes = (primerDia === 1) ? 1 : ((8 - primerDia) % 7) + 1;
    // sumar 14 días para llegar al tercer lunes
    return { mes: 10, dia: primerLunes + 14 };
  }

  function esFestivo(d) {
    var year = d.getFullYear();
    // agregar el tercer lunes de noviembre dinámico
    var festivoMovil = tercerLunesNoviembre(year);
    var todosFestivos = festivos.concat(festivoMovil);

    return todosFestivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  // convertir la fecha de la fila a objeto Date
  var partes = fila[0].split("/");
  var fecha = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));

  if (esFestivo(fecha)) {
    var dia = fecha.getDay();
    // Miércoles festivo → martes
    if (dia === 3) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() - 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        fila[0] = formatearFechaQ(martes);
        return fila;
      }
    }
    // Lunes festivo → martes
    if (dia === 1) {
      var martes = new Date(fecha);
      martes.setDate(fecha.getDate() + 1);
      if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) {
        fila[0] = formatearFechaQ(martes);
        return fila;
      }
    }
  }

  return fila;
}

/*semanal */
/*
codigo que de la hoja horigen "S.Gastos CICLICOS INTERNO PS" sacaras la periosdad cuando sea en la columna 29 === "CADA LUNES || CADA VIERNES ||
DIARIO" y cuando encuentes este cadena vas a sacar laS fechaS: cada lunes de cada mes y viernes y diario sin contar sabado o domigos y esas fechas se pegaran en la columna A de la hojas destino : "Fechas 2025" y despes de la columna A ira el rango de C:AA de la hoja origen corespondiente perioridad que salga "CADA LUNES || CADA VIERNES ||
DIARIO", el rango a sacar de la hoja destino es "A:AD".
/*
    Extrae de la hoja origen "S.Gastos CICLICOS INTERNO PS" las filas donde la columna 29 (AC) sea "CADA LUNES", "CADA VIERNES" o "DIARIO".
    Para cada coincidencia, genera todas las fechas de 2025 según la periodicidad:
        - "CADA LUNES": todos los lunes del año (sin sábados ni domingos)
        - "CADA VIERNES": todos los viernes del año (sin sábados ni domingos)
        - "DIARIO": todos los días del año excepto sábados y domingos
    Por cada fecha, pega en la hoja destino "Fechas 2025" en la columna A la fecha, y de la B a la AA los datos de la fila origen (C:AA).
    El rango de salida es "A:AD".
*/
function generarSemanal() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libro.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  var hojaDestino = libro.getSheetByName("Planeador Despacho");
  //var hojaDestino = libro.getSheetByName("hojaPrueba");

  //var hojaOrigen = libro.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libro.getSheetByName("Planeador Personal");

  var datos = hojaOrigen.getRange("A:AE").getValues();
  var ultimaFilaDestino = hojaDestino.getLastRow();

  // Rango de fechas personalizable
  var anioInicio = 2025, mesInicio = 8;  // Septiembre 2025 (0=Enero)
  var anioFin = 2026, mesFin = 11;       // Diciembre 2026
  /*var fechaInicio = new Date(anioInicio, mesInicio, 1);
  var fechaFin = new Date(anioFin, mesFin + 1, 0);*/

  var salida = [];

  for (var i = 1; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase();
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();
    if (periodicidad !== "CADA LUNES" && periodicidad !== "CADA VIERNES" && periodicidad !== "DIARIO") continue;
  
    //var filaDatos = datos[i].slice(2, 27);
    var filaDatos = datos[i].slice(2, 28);
  
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      // Determina el mes de inicio y fin para el primer y último año
      var mesIni = (anio === anioInicio) ? mesInicio : 0;
      var mesFinLoop = (anio === anioFin) ? mesFin : 11;
  
      for (var mes = mesIni; mes <= mesFinLoop; mes++) {
        var fechaIniMes = new Date(anio, mes, 1);
        var fechaFinMes = new Date(anio, mes + 1, 0);
  
        var fechas = [];
        if (periodicidad === "CADA LUNES") fechas = obtenerDiasPorSemanaRango(fechaIniMes, fechaFinMes, 1);
        else if (periodicidad === "CADA VIERNES") fechas = obtenerDiasPorSemanaRango(fechaIniMes, fechaFinMes, 5);
        else if (periodicidad === "DIARIO") fechas = obtenerDiasHabilesRango(fechaIniMes, fechaFinMes);
  
        fechas.forEach(function(fecha) {
          fecha = ajustarPorFestivoSem(fecha);
          var fila = [formatearFechaS(fecha)].concat(filaDatos);
          fila.push("NUEVO");
          salida.push(fila);
          Logger.log("Periodicidad: " + periodicidad + " | Fecha generada: " + fecha);
        });
      }
    }
  }

  // Escribir en hoja destino desde la columna B
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida);
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
             .setNumberFormat("dd/MM/yyyy");
  }
}

// Días hábiles dentro de un rango
function obtenerDiasHabilesRango(fechaInicio, fechaFin) {
  var fechas = [];
  var f = new Date(fechaInicio);
  while (f <= fechaFin) {
    if (f.getDay() !== 0 && f.getDay() !== 6) fechas.push(new Date(f));
    f.setDate(f.getDate() + 1);
  }
  return fechas;
}

// Días de semana (lunes=1, viernes=5) dentro de un rango
function obtenerDiasPorSemanaRango(fechaInicio, fechaFin, diaSemana) {
  var fechas = [];
  var f = new Date(fechaInicio);
  while (f <= fechaFin) {
    if (f.getDay() === diaSemana) fechas.push(new Date(f));
    f.setDate(f.getDate() + 1);
  }
  return fechas;
}

// Formato de fecha
function formatearFechaS(fecha) {
  //return fecha.getDate() + "/" + (fecha.getMonth() + 1) + "/" + fecha.getFullYear();
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}


/*
si cae lune dia festivo se pasa a martes de eesa semana
si cae viernes dia festivo se pasa a miercoles de esa semana
si cae diario seria los siguientes:
  si cae lunes dia festivo que lo mueva para el martes de ese semana
  si cae martes dia festivo que lo mueva para el miercoles de esa semana
  si cae miercoles dia festivo se pasa a martes de esa semana
  si cae jueves dia festivo que lo mueva para el miercoles de esa semana
  si cae viernes dia festivo que lo mueva para el miercoles de esa semana
*/

function ajustarPorFestivoSem(fecha) {
  // Festivos fijos
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  // Calcular tercer lunes de noviembre para el año dado
  function tercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 de noviembre
    var primerDia = fecha.getDay();    // 0=Dom,...,6=Sab
    // calcular primer lunes
    var primerLunes = (primerDia === 1) ? 1 : ((8 - primerDia) % 7) + 1;
    // sumar 14 días para llegar al tercer lunes
    return { mes: 10, dia: primerLunes + 14 };
  }

  function esFestivo(d) {
    var year = d.getFullYear();
    var festivoMovil = tercerLunesNoviembre(year);
    var todosFestivos = festivos.concat(festivoMovil);

    return todosFestivos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);
  }

  if (!fecha) return fecha;

  var dia = fecha.getDay(); // 0=Dom, 1=Lun,...,5=Vie

  if (!esFestivo(fecha)) return fecha;

  // Lunes festivo → martes
  if (dia === 1) {
    var martes = new Date(fecha);
    martes.setDate(fecha.getDate() + 1);
    if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) return martes;
  }

  // Martes festivo → miércoles
  if (dia === 2) {
    var miercoles = new Date(fecha);
    miercoles.setDate(fecha.getDate() + 1);
    if (!esFestivo(miercoles) && miercoles.getDay() !== 0 && miercoles.getDay() !== 6) return miercoles;
  }

  // Miércoles festivo → martes
  if (dia === 3) {
    var martes = new Date(fecha);
    martes.setDate(fecha.getDate() - 1);
    if (!esFestivo(martes) && martes.getDay() !== 0 && martes.getDay() !== 6) return martes;
  }

  // Jueves festivo → miércoles
  if (dia === 4) {
    var miercoles = new Date(fecha);
    miercoles.setDate(fecha.getDate() - 1);
    if (!esFestivo(miercoles) && miercoles.getDay() !== 0 && miercoles.getDay() !== 6) return miercoles;
  }

  // Viernes festivo → miércoles
  if (dia === 5) {
    var miercoles = new Date(fecha);
    miercoles.setDate(fecha.getDate() - 2);
    if (!esFestivo(miercoles) && miercoles.getDay() !== 0 && miercoles.getDay() !== 6) return miercoles;
  }

  // Si no aplica ninguna regla, devuelve la fecha original
  return fecha;
}

/*trimestral */
function generarTrimestral() {
  var libroOrigen = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = libroOrigen.getSheetByName("S.Gastos CICLICOS INTERNO PS");
  var hojaDestino = libroOrigen.getSheetByName("Planeador Despacho");
  //var hojaDestino = libroOrigen.getSheetByName("hojaPrueba");
  
  //var hojaOrigen = libroOrigen.getSheetByName("Copia de S.Gastos CICLICOS INTERNO PS(Personal)");
  //var hojaDestino = libroOrigen.getSheetByName("Planeador Personal");
  
  var datos = hojaOrigen.getRange("A:AE").getValues();

  var ultimaFilaDestino = hojaDestino.getLastRow();

  // Rango de fechas personalizable
  var anioInicio = 2025, mesInicio = 8;  // Septiembre 2025 (0=Enero)
  var anioFin = 2026, mesFin = 11;       // Diciembre 2026

  // Meses trimestrales
  var unoLu = [0, 4, 8]; // 1ER LUNES DE ENE, MAYO Y SEP
  var segundoMier = [2, 5, 8, 11]; // 2DO MIERCOLES DE MAR, JUN, SEP, DIC


  var periodicidades = {
    "1ER LUNES DE ENE, MAYO Y SEP": function(anio) { 
      return obtenerTrimestral(anio, unoLu, "1ER_LUNES"); 
    },
    "2DO MIERCOLES DE MAR, JUN, SEP, DIC": function(anio) { 
      return obtenerTrimestral(anio, segundoMier, "2DO_MIERCOLES"); 
    }
  };


  var salida = [];

  for (var i = 5; i < datos.length; i++) {
    //var periodicidad = (datos[i][29] || "").toString().trim().toUpperCase();
    var periodicidad = (datos[i][30] || "").toString().trim().toUpperCase();
    var funcion = periodicidades[periodicidad];
    if (!funcion) continue;

    // Tomar columnas C:AA (índices 2 a 26)
    var filaDatos = datos[i].slice(2, 27);
    //var filaDatos = datos[i].slice(2, 28);

    // Recorrer los años dentro del rango
    for (var anio = anioInicio; anio <= anioFin; anio++) {
      var fechas = funcion(anio); // Devuelve un arreglo de fechas trimestrales
      fechas = ajustarPorFestivoTrime(fechas); // ajusta lunes/miércoles festivos

      fechas.forEach(function(fecha) {
        // Solo fechas dentro del rango
        if (fecha >= new Date(anioInicio, mesInicio, 1) && fecha <= new Date(anioFin, mesFin, 28)) {
          var fechaFormateada = formatearFechaT(fecha);
          var nuevaFila = [fechaFormateada].concat(filaDatos);

          // Asegura que llegue hasta la columna AB
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

  // Escribe la salida en la hoja destino desde la columna B
  if (salida.length > 0) {
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, salida[0].length).setValues(salida);
    // ✅ Formatear la columna de fechas (col B)
    hojaDestino.getRange(ultimaFilaDestino + 1, 2, salida.length, 1)
             .setNumberFormat("dd/MM/yyyy");
  }
}

// Formato de fecha
function formatearFechaT(fecha) {
  if (!fecha) return "";
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1;
  var anio = fecha.getFullYear();
  return dia + "/" + mes + "/" + anio;
}

// Funciones para calcular lunes/miércoles
function primerLunesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() !== 1) fecha.setDate(fecha.getDate() + 1);
  return fecha;
}

function primerMiercolesDelMes(anio, mes) {
  var fecha = new Date(anio, mes, 1);
  while (fecha.getDay() !== 3) fecha.setDate(fecha.getDate() + 1);
  return fecha;
}

function segundoLunesDelMes(anio, mes) {//no se ocupa
  var fecha = primerLunesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 7);
  return fecha;
}

function segundoMiercolesDelMes(anio, mes) {
  var fecha = primerMiercolesDelMes(anio, mes);
  fecha.setDate(fecha.getDate() + 7);
  return fecha;
}

// Devuelve fechas trimestrales correctas
function obtenerTrimestral(anio, mesesArr, tipo) {
  var fechas = [];
  mesesArr.forEach(function(mes) {
    if (tipo === "1ER_LUNES") {
      fechas.push(primerLunesDelMes(anio, mes));
    }
    if (tipo === "2DO_MIERCOLES") {
      fechas.push(segundoMiercolesDelMes(anio, mes));
    }
  });
  return fechas;
}

//dias festivos
/*
si cae un dia festivo en miercoles que pase al mastes de esa semana
si cae lunes festivo que pase al mastes de esa semana
y si no pasa nada pasa
 */
/*function ajustarPorFestivoTrime(fechas) {
  var festivos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 10, dia: 20 }, // 20 Noviembre
    { mes: 11, dia: 12 },  // 12 Diciembre
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
function ajustarPorFestivoTrime(fechas) {
  var festivosFijos = [
    { mes: 0, dia: 1 },   // 1 Enero
    { mes: 1, dia: 5 },   // 5 Febrero
    { mes: 2, dia: 21 },  // 21 Marzo
    { mes: 4, dia: 1 },   // 1 Mayo
    { mes: 8, dia: 16 },  // 16 Septiembre
    { mes: 11, dia: 12 }, // 12 Diciembre
    { mes: 11, dia: 25 }  // 25 Diciembre
  ];

  // 👉 calcular dinámicamente el 3er lunes de noviembre
  function obtenerTercerLunesNoviembre(year) {
    var fecha = new Date(year, 10, 1); // 1 Noviembre
    var primerDia = fecha.getDay();    // 0=Dom, 1=Lun...
    var primerLunes = primerDia === 1 ? 1 : (8 - primerDia);
    var tercerLunes = primerLunes + 14; // tercer lunes = primer lunes + 14 días
    return new Date(year, 10, tercerLunes);
  }

  function esFestivo(d) {
    // festivos fijos
    var esFijo = festivosFijos.some(f => d.getMonth() === f.mes && d.getDate() === f.dia);

    // tercer lunes de noviembre
    var tercerLunes = obtenerTercerLunesNoviembre(d.getFullYear());
    var esTercerLunes = d.getMonth() === 10 && d.getDate() === tercerLunes.getDate();

    return esFijo || esTercerLunes;
  }

  // Recorrer todas las fechas y ajustar
  return fechas.map(function (fecha) {
    if (!fecha) return fecha;

    if (!esFestivo(fecha)) return fecha; // si no es festivo, regresar igual

    var dia = fecha.getDay(); // 0=Dom, 1=Lun, 2=Mar, 3=Mié...

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
}
