function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('📑 | Layout')
    .addItem('1. Borrar Layout  | 📄', 'metodoEliminarV02')
    .addToUi();
}

function metodoHerenciaGastosDespacho(){//para la herencia
  layoutMasterV1("1wEwNRBi24ezui460nSdsHpnsYGe3A7zGs3C5cBt0_bM", "SOLICITUD GASTOS TEMPORAL - CONCATENADO");
  SpreadsheetApp.flush(); // Fuerza la escritura de los cambios /*esto afuerza que suelte el pegado del 003 para que lo lleve en el excel */
}

function layoutMasterV1(libroOrigenLink, hojaOrigenNombre) {
  var libroOrigen = SpreadsheetApp.openById(libroOrigenLink); //Master concentrado
  var libroDestino = SpreadsheetApp.openById("1b2vIve0yzxHBL5ty7Kn59cmwJ40Wa_FdQwaSZnILOgM"); //Layout V3

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

              var importe = dataOrigen[i][23] ? dataOrigen[i][23].toString().replace(/[-,]/g, "") : ""; //Importe, sin giones 6

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

    // Agrupar por identificador y sumar importes
    var agrupados = {};
    for (var j = 0; j < filasPegar.length; j++) {
        var fila = filasPegar[j];
        var identificador = fila[0];
        var banco = fila[20];
        var tipo = fila[17];
        var comenEmpresa = fila[19];
        var clabe = fila[21];
        var titular = fila[22];
        var importeStr = fila[23];
        // Quitar formato de moneda para sumar
        var importeNum = parseFloat(importeStr.toString().replace(/[^0-9.-]+/g,"")) || 0;

        if (!agrupados[identificador]) {
            agrupados[identificador] = {
                identificador,
                banco,
                tipo,
                comenEmpresa,
                clabe,
                titular,
                importe: 0
            };
        }
        agrupados[identificador].importe += importeNum;
    }

    // Preparar los datos a pegar con la estructura requerida
    var datosParaPegar = [];
    var idx = 1;
    for (var key in agrupados) {
        var item = agrupados[key];
        // Formatear el importe como moneda MXN
        var importeFormateado = item.importe.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' });
        datosParaPegar.push([
            idx++,           // ID consecutivo
            item.banco,      // Banco
            item.tipo,       // Tipo TD o TC
            item.comenEmpresa, // COMENTARIOS DE ENTREGA
            item.clabe,      // CLABE
            item.titular,    // Titular
            importeFormateado // Importe sumado y formateado
        ]);
    }


    // Formatea columna C (índice 4) como texto (solo la columna de CLABE)
    hojaDestino.getRange(startRow, 5, datosParaPegar.length, 1).setNumberFormat("@");
   // hojaDestino.getRange(startRow, 4, datosParaPegar.length, 1).setNumberFormat("@");

    // Ahora sí: Pega todos los datos
    hojaDestino.getRange(startRow, 1, datosParaPegar.length, datosParaPegar[0].length).setValues(datosParaPegar);

    Logger.log(`${datosParaPegar.length} filas pegadas en hojaDestino.`);
  } else {
    Logger.log("No se encontraron filas con la fecha de hoy.");
  }
}

//eliminar los datos del Layout
function metodoEliminarV02(){ //MODIFICADO
    var libroDestino = SpreadsheetApp.openById("1b2vIve0yzxHBL5ty7Kn59cmwJ40Wa_FdQwaSZnILOgM"); //Layout
    var hojaDestino = libroDestino.getSheetByName("Layout");

    var ultimaFila = hojaDestino.getLastRow();
    if (ultimaFila > 0) {
        hojaDestino.getRange(2, 1, ultimaFila, 7).clearContent(); // Desde A2:E[ultimaFila] //fila, columna, filaUltima, ColumnaFinal
        Logger.log("Contenido eliminado de A2:G" + ultimaFila);
    }
}
