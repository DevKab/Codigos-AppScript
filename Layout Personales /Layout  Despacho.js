function onOpen() { //05/09/2025
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('📑 | Layout')
    .addItem('1. Borrar Layout  | 📄', 'metodoEliminarV02')
    .addToUi();
}

function metodoHerenciaGastosDespacho(){//para la herencia
  layoutMasterV1("1Am2CQKrYwHX1nYYnC81N4oPHn4wFeV1RPieLt3GsJcM", "SOLICITUD GASTOS TEMPORAL - CONCATENADO");
  SpreadsheetApp.flush(); // Fuerza la escritura de los cambios /*esto afuerza que suelte el pegado del 003 para que lo lleve en el excel */
}

function layoutMasterV1(libroOrigenLink, hojaOrigenNombre) {
  var libroOrigen = SpreadsheetApp.openById(libroOrigenLink); //Master concentrado
  var libroDestino = SpreadsheetApp.openById("1B-Pp9g1vvp7_NpaI3eFf-OC-MsuIwr0Wo-GQduvZfkw"); //Layout V3

   var hojaOrigen = libroOrigen.getSheetByName(hojaOrigenNombre);
  var hojaDestino =  libroDestino.getSheetByName("Layout");

  //sacamos los datos de la hoja destino
  //var dataOrigen = hojaOrigen.getRange("R:AD").getValues(); //solo las columnas para el layaout
  var dataOrigen = hojaOrigen.getRange("A:AC").getValues(); //solo las columnas para el layaout AC

  //arreglo para agregar las filas filtradas
  var filasPegar = [];

  //interamos para el bucle
  for(var i = 0; i < dataOrigen.length; i++){

    //verificar si esta paado el gasto
    if(dataOrigen[i][28] === "EN PROCESO"){//11
      if(dataOrigen[i][17] === "TRANSFERENCIA"|| dataOrigen[i][17] === "TARJETA DE CREDITO"){//0
        if(dataOrigen[i][18] === "NACIONAL"){//18
            var tipoTarjeta = "";
            if(dataOrigen[i][17] === "TRANSFERENCIA"){
              if( dataOrigen[i][20]=== "AMERICAN EXPRESS"){
                  tipoTarjeta = "TC";
              }
              else{
                  tipoTarjeta = "TD";
              }
                
            }else if(dataOrigen[i][17] === "TARJETA DE CREDITO"){
                tipoTarjeta = "TC";
            }
            
            //validar que no sea bancoopel ni famsa(sin importar mayusculas/minusculas)
            var banco = dataOrigen[i][20] ? dataOrigen[i][20].toString().toLowerCase().trim() : "";
            if(banco === "bancoopel" || banco === "famsa" || banco === "") continue;// Salta esta fila//3


            // Validar y limpiar número de tarjeta en la col. CLABE DESTINO
            if (dataOrigen[i][21] && dataOrigen[i][21].toString().trim() !== "") {
              // Paso 1: Quitar espacios
              const limpia = dataOrigen[i][21].replace(/\s+/g, '');//4

              // Paso 2: Buscar todos los bloques de dígitos de 15, 16 o 18 caracteres
              const bloques = limpia.match(/\d{15,18}/g);

              let tarjetasLimpia = "";

              if (bloques && bloques.length > 0) {
                // Recorre todos los bloques encontrados y guarda el primero que cumpla la condición
                for (let b = 0; b < bloques.length; b++) {
                  const num = bloques[b];
                  if (
                    num.length === 18 ||
                    num.length === 16 ||
                    (num.length === 15 && (num.startsWith("34") || num.startsWith("37")))
                  ) {
                    tarjetasLimpia = num;
                    break; // Solo toma el primero válido
                  }
                }
                Logger.log("Tarjeta valida: " + tarjetasLimpia + " tipo de tarjeta " + tipoTarjeta);
                if (!tarjetasLimpia) continue; // Si no encontró ninguna válida, salta la fila
              } else {
                // Si no hay bloques, intenta limpiar todo y validar
                const soloNumeros = limpia.replace(/\D/g, "");
                if (
                  soloNumeros.length === 18 ||
                  soloNumeros.length === 16 ||
                  (soloNumeros.length === 15 && (soloNumeros.startsWith("34") || soloNumeros.startsWith("37")))
                ) {
                  tarjetasLimpia = soloNumeros;
                  Logger.log("Tarjeta valida: " + tarjetasLimpia + " tipo de tarjeta " + tipoTarjeta);
                } else {
                  continue; // Si no cumple, salta la fila
                }
              }
              //}

              //Titular
              // Validar que AE no tenga Ñ, . , o espacios al final /* Para limpiar todo el valor si contiene caracteres no permitidos, asigna "". */
              //cambia "ñ" a "n": dataOrigen[i][5].toString().replace(/[\ñ]/gi, 'n') para mayusculas toLocaleUpperCase()
              //var titular = dataOrigen[i][5] ? dataOrigen[i][5].toString().replace(/[\ñ]/gi, 'n').replace(/\./g, "").replace(/,/g, "").replace(/\s+$/, "").toLocaleUpperCase() : ""; 
              // Elimina ñ, ., , y espacios al final
                var titular = dataOrigen[i][22]
                ? dataOrigen[i][22]
                  .toString()
                  .normalize("NFD")                // Quita tildes
                  .replace(/[\u0300-\u036f]/g, "") // Quita los signos diacríticos (tildes)
                  .replace(/[\ñ]/gi, 'n')    // reemplaza ñ por n
                  .replace(/[\d\W_]+/g, " ") // elimina números y caracteres no alfabéticos, deja espacios
                  .replace(/\s+$/, "")       // elimina espacios al final
                  .replace(/\s{2,}/g, " ")   // elimina espacios dobles
                  .toLocaleUpperCase()
                  .replace(/\b(BBVA|CLAVE BANCARIA|CLAVE|AFIRME|BANORTE|AMERICAN EXPRESS|AZTECA|BANAMEX|BANREGIO|SANTANDER|SCOTIABANK|HSBC|N\/A|CI BANCO)\b/g, "") // elimina palabras prohibidas
                  .replace(/\s{2,}/g, " ")   // elimina espacios dobles generados por el replace anterior
                  .trim()
                : "";
                  /* quitar BBVA,  clave bancaria, clave, AFIRME, BANORTE, AMERICAN EXPRESS, AZTECA, BANORTE, BANAMEX,BANREGIO, SANTANDER, SCOTIABANK, HSBC, N/A, CI BANCO,*///5

                if(titular === "") continue;// Salta esta fila vacia

                //var importe = dataOrigen[i][23] ? dataOrigen[i][23].toString().replace(/[-,]/g, "") : ""; //Importe, sin giones 6
                var importe = dataOrigen[i][23] ? dataOrigen[i][23].toString().replace(/[,]/g, "") : "";

                // Convertir a número y formatear como moneda MXN
                var importeNum = parseFloat(importe);
                if (!isNaN(importeNum)) {
                  importe = importeNum.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' });
                } else {
                  importe = "";
                }
                if(importe === "") continue;// Salta esta fila vacia

                var comentarioEmpresa = dataOrigen[i][19] 
                ? dataOrigen[i][19]
                .toString()
                .replace(/\s+$/, "")       // elimina espacios al final
                .replace(/\s{2,}/g, " ")   // elimina espacios dobles
                .trim() : ""; //aqui estoy

                if(comentarioEmpresa === "") continue;//salta la columna vacia

                // Validar empresa permitida
                var empresasValidas = new Set([
                  "2GA", "9/16", "ACCEROX", "ACEROMEX", "ADMAS", "AFB", "ALFA88", "ALFAREY", "ALFASEG", "ALFSTAR","ALGATICA", "ALGORITT", "ALIMSA", "ALLEN", "ALLFOOD", "ALMAR", "ALORA", "ALQUINCO", "AM&CE", "APB","AQR", "ARAUCCO", "ARBOK", "ARCE", "ARYBE", "ASPEN", "ATRIO", "AURIMETAL", "AVENTIA", "AXELIA", "AZYTEC", "BACKCOM", "BALAY", "BANDI", "BASSALTO", "BE&M", "BEJUCO", "BEMACK", "BERCKER", "BERETH", "BERTE", "BETRUCK","BEXTUS", "BICTTO", "BIDANTA", "BOGANT", "BOWITT", "BQ7", "BRIXCO", "BRIXMAN", "BROCCA", "BROSISSA", "BRUNCH", "CALTIGA", "CAOBA", "CAVANNA", "CETEC", "CIMENTIA", "CLEAN & SHINE", "CLEAN&CO", "CLEAN&SHINE", "CLEANMEX",
                  "CLEANPLACE", "CMP", "COMPUMAS", "COMPUTECH", "CONCRETOMEX", "CONSTRUCTURE", "CRM", "CRONEK", "CRT", "CWO","CYGNUS", "DAEGU", "DALAC", "DELCO", "DELLOW", "DELORIA", "DENTRUCK", "DEYMA", "DIXEN", "DRAWEN", "DRISCOLL","DURLINK", "EDIFIK", "ELTE", "ENDERCO", "EQ. DEL NORTE", "EVORA", "F&B", "FELDER", "FELER", "FERPREX","FERRECSA", "FIDELIS", "FISCASE", "FIVE STAR", "FORTCALL", "FORTEM", "FORTEX", "FRESNO", "FUSION", "G 10","GABAM", "GADISSA", "GAMALTA", "GASTEL", "GATRE", "GAYU", "GEMSE", "GENNOA", "GENOX", "GEOTERRA", "GERONNA","GESTIONA", "GLOBAL", "GNQ", "GOBBI", "GOLIA", "GP", "GRABUM", "GRAMEX", "GRAVLER", "GRAVMARK", "INDIGO", "INMOBILIARIA BROSISSA", "INOSTAR", "INTER TRUCKING", "INTERFOOD", "INTERPREX", "K11", "KABBA", "KADYL",
                  "KAPAM", "KARSE", "KATPRO", "LA NATIVA", "LANDECK", "LANN", "LATIMEX", "LAWRE", "LINE 123", "LIVETT","LOGIXEN", "LOGTEK", "LOWTT", "LUSOL", "LUXA", "LUXO", "LYON", "MABBO", "MADACSA", "MADERERIA", "MANON","MARGAL", "MARTE", "MATERIALES", "MATINSA", "MDM", "MEDALLO", "MEDICMAS", "MEDICSA", "MEGSA", "MOB Y EQ","MOBIMAX", "MONAVI", "MONTRED", "MOVED", "MQ", "MURETT", "MUTTANT", "MUTTED", "OCCINORTE", "ODESSA","OFIMEX", "OLENKA", "ORIOON", "OSTO", "OZMA", "PAPELERIA REAL", "PBS", "PITFULL", "PLASTIQ", "PROACTEC","PROFIX", "PROMEC", "PROSCAR", "PROSTEEL", "PROXTEC", "QTR", "QUALITTY", "QUANTTIC", "RADXO", "RCC", "REGIO EXPRESS", "RIU", "RIVAS", "RODRETT", "RODYKEY", "ROGERS", "RUBRAK", "SAGGE", "SEGGUSA", "SEMAX","SOLIXTIC", "TACTIK", "TARGET", "TECMAX", "TECNOFIX", "TENDERMAX", "TENZA", "TERRA4", "TERRANOVA","TESORERÍA CAPITEL 1004", "TESORERÍA CAPITEL 905", "TESORERÍA FUNDADORES", "TESORERÍA GUADALAJARA","TESORERÍA PLAYA", "TFG", "TGR", "TIERRA FUERTE", "TITAN", "TOMSON", "TORREXA", "TOSCAR", "TOWSON","TREVIA", "TRF", "TRIVENTTO", "TRUCKFULL", "TUXO", "UB41", "URBAN52", "URBANITMO", "VANTO", "VECTRA",
                  "VESSEL", "VIAYA", "VIEYRA", "VIGANT", "XENO", "ZAC", "N/A"
                ]);

                if (!empresasValidas.has(comentarioEmpresa)) continue; // Empresa no permitida


              //Guardar los datos limpios en la fila original
              
              dataOrigen[i][17] = tipoTarjeta; //detalle de pago
              dataOrigen[i][19] = comentarioEmpresa; //comentarios de entrega Col.T
              dataOrigen[i][21] = tarjetasLimpia;//numero de tarjeta
              dataOrigen[i][22] = titular;
              dataOrigen[i][23] = importe;
              
              filasPegar.push(dataOrigen[i]); //Añadiendo fila para pegar
                    
            }
        }
      }
    }
  }
    if (filasPegar.length > 0) {
      var ultimaFilaDestino = hojaDestino.getLastRow();
      var startRow = ultimaFilaDestino + 1;

      // Agrupar por identificador y sumar importes según la condición solicitada
      var importesPorId = {};

      // Primero, agrupa todas las filas por identificador y guarda los importes originales
      for (var j = 0; j < filasPegar.length; j++) {
        var fila = filasPegar[j];
        var identificador = fila[0];
        var banco = fila[20];
        var tipo = fila[17];
        var comenEmpresa = fila[19];
        var clabe = fila[21];
        var titular = fila[22];
        var importeStr = fila[23];
        var importeNum = parseFloat(importeStr.toString().replace(/[^0-9.-]+/g, "")) || 0; // Quitar formato de moneda para sumar

        if (!importesPorId[identificador]) {
          importesPorId[identificador] = [];
        }
        importesPorId[identificador].push({
          banco,
          tipo,
          comenEmpresa,
          clabe,
          titular,
          importeNum
        });
      }

      // Ahora, procesa cada identificador según la lógica solicitada
      var datosParaPegar = [];
      var idx = 1;
      for (var key in importesPorId) {
        // Si no hay identificador (columna A vacía), no se pasa
        if (!key || key.trim() === "") continue;

        var items = importesPorId[key];
        var sumaNegativos = 0;
        var sumaPositivos = 0;
        var tienePositivo = false;

        // Suma negativos y positivos
        for (var k = 0; k < items.length; k++) {
          var imp = items[k].importeNum;
          if (imp < 0) sumaNegativos += imp;
          if (imp > 0) {
            sumaPositivos += imp;
            tienePositivo = true;
          }
        }

        var totalImporte = 0;
        if (tienePositivo) {
          // Si hay positivos, suma todos los negativos y positivos, pero la porción negativa se resta
          totalImporte = sumaPositivos + sumaNegativos;
        } else {
          // Si no hay positivos, suma todos los negativos (como está el código original)
          totalImporte = sumaNegativos;
        }

        // Usa los datos del primer elemento para los demás campos
        var item = items[0];

        // 👇 Aplica valor absoluto al total para quitar el signo negativo
        totalImporte = Math.abs(totalImporte);//modificsacion 15/08/2025

        var importeFormateado = totalImporte.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' });

        datosParaPegar.push([
          idx++,
          item.banco,
          item.tipo,
          item.comenEmpresa, // COMENTARIOS DE ENTREGA
          item.clabe,
          item.titular,
          key,                // ← Identificador en la columna G
          importeFormateado
        ]);
      }

      hojaDestino.getRange(startRow, 5, datosParaPegar.length, 1).setNumberFormat("@");
      hojaDestino.getRange(startRow, 1, datosParaPegar.length, datosParaPegar[0].length).setValues(datosParaPegar);


      Logger.log(`${datosParaPegar.length} filas pegadas en hojaDestino.`);
  } else {
    Logger.log("No se encontraron filas con la fecha de hoy.");
  }
}

//eliminar los datos del Layout
function metodoEliminarV02(){
    var libroDestino = SpreadsheetApp.openById("1B-Pp9g1vvp7_NpaI3eFf-OC-MsuIwr0Wo-GQduvZfkw"); //Layout
    var hojaDestino = libroDestino.getSheetByName("Layout");

    var ultimaFila = hojaDestino.getLastRow();
    if (ultimaFila >= 2) {
        // Desde A2:H[ultimaFila]
        hojaDestino.getRange(2, 1, ultimaFila - 1, 8).clearContent(); // Desde A2:E[ultimaFila] //fila, columna, filaUltima, ColumnaFinal
        Logger.log("Contenido eliminado de A2:G" + ultimaFila);
    }
}
