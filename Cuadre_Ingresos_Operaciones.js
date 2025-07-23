function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(' ➠ Transferir Datos')
        .addItem('Iniciar Envío', 'enviarInfo')
        .addItem('Envio Provisionadas', 'enviarProv')
        .addItem('Check test', 'btns')
        // .addItem('Insertar Fecha', 'fecha')

    .addToUi();
      OperFact.mostrarMensaje();
}

//////////////////////////////

function onEdit(e) {
  var celdaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  var fila = celdaActiva.getRow() //fila activa que se bloquea
  var columna = celdaActiva.getColumn() //fila activa que se bloquea
  var hoja = celdaActiva.getSheet();
  var valor = celdaActiva.getValue();

  const correoEditor = Session.getActiveUser().getEmail();
  const fechaModificacion = new Date();

  if (hoja.getName() === "DOCUMENTOS FACTURACION" && columna === 26) {
    (valor)?hoja.getRange(fila, columna+1, 1, 2).setValues([[correoEditor, fechaModificacion]]):hoja.getRange(fila, columna+1, 1, 2).clearContent();
  }

  if (hoja.getName() === "MOVS" && columna === 19) {
    (valor)?hoja.getRange(fila, columna+1, 1, 2).setValues([[correoEditor, fechaModificacion]]):hoja.getRange(fila, columna+1, 1, 2).clearContent();
  }
}

//////////////////////////////

function verificarFacturas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMovs = ss.getSheetByName("MOVS");
  const hojaDocs = ss.getSheetByName("DOCUMENTOS FACTURACION");

  if (!hojaMovs || !hojaDocs) {
    Logger.log("Asegúrate de que ambas hojas existan: MOVS y DOCUMENTOS FACTURACION");
    return;
  }

  // Obtener datos de las hojas
  const datosMovs = hojaMovs.getDataRange().getValues();
  const datosDocs = hojaDocs.getDataRange().getValues();

  const resultados = []; // Almacena los resultados

  // Recorrer las filas de MOVS (empezar en la fila 2 para omitir encabezados)
  for (let i = 1; i < datosMovs.length; i++) {
    try {
      Logger.log(`--- Procesando MOVS fila ${i + 1} ---`); // Seguimiento de progreso

      const mesaMovs = String(datosMovs[i][1]).trim(); // Columna B
      const promotorMovs = String(datosMovs[i][4]).trim(); // Columna E
      const idMovs = String(datosMovs[i][5]).trim(); // Columna F
      const empresaMovs = String(datosMovs[i][8]).trim(); // Columna I
      const totalMovs = parseFloat(datosMovs[i][11]); // Columna L (Total)

      if (!totalMovs) {
        Logger.log(`MOVS fila ${i + 1}: Sin total, saltando.`);
        continue; // Saltar si el total está vacío
      }

      let sumaFacturas = 0;
      const filasRelacionadas = [];

      // Buscar coincidencias en DOCUMENTOS FACTURACION
      for (let j = 1; j < datosDocs.length; j++) {
        const mesaDocs = String(datosDocs[j][18]).trim(); // Columna S
        const promotorDocs = String(datosDocs[j][17]).trim(); // Columna R
        const vendedorDocs = String(datosDocs[j][11]).trim(); // Columna L
        const empresaDocs = String(datosDocs[j][1]).trim(); // Columna B
        const factura = parseFloat(datosDocs[j][5]); // Columna F (Total Factura)

        // Comparar columnas especificadas (sin fecha)
        if (
          mesaMovs === mesaDocs &&
          promotorMovs === promotorDocs &&
          idMovs === vendedorDocs &&
          empresaMovs === empresaDocs
        ) {
          sumaFacturas += factura || 0; // Sumar factura si existe
          filasRelacionadas.push(j + 1); // Guardar fila (en formato humano)

          // Imprimir solo coincidencias
          Logger.log(
            `DOCUMENTOS fila ${j + 1}: Mesa=${mesaDocs}, Promotor=${promotorDocs}, Vendedor=${vendedorDocs}, Empresa=${empresaDocs}, Factura=${factura}`
          );
        }
      }

      // Comparar la suma de facturas con el total de la hoja MOVS
      if (sumaFacturas === totalMovs) {
        resultados.push(
          `Movimiento fila ${i + 1}: Facturas encontradas en las filas ${filasRelacionadas.join(", ")}. Total MOVS: ${totalMovs}, Suma Facturas: ${sumaFacturas}`
        );
      } else if (filasRelacionadas.length > 0) {
        resultados.push(
          `Movimiento fila ${i + 1}: Facturas encontradas en las filas ${filasRelacionadas.join(", ")}, pero los totales no coinciden. Total MOVS: ${totalMovs}, Suma Facturas: ${sumaFacturas}`
        );
      } else {
        Logger.log(`MOVS fila ${i + 1}: No se encontraron coincidencias en DOCUMENTOS.`);
      }
    } catch (error) {
      Logger.log(`Error procesando la fila ${i + 1} de MOVS: ${error.message}`);
    }
  }

  // Escribir los resultados en el log
  if (resultados.length > 0) {
    Logger.log("Resultados:\n" + resultados.join("\n"));
  } else {
    Logger.log("No se encontraron coincidencias.");
  }
}

//////////////////////////////

function enviarInfo() {
   // Verifica si ya se ha enviado la informacion
  if(SpreadsheetApp.getActive().getSheetByName("DOCUMENTOS FACTURACION").getRange("AF1").getValue() != "V"){
    OperFact.enviarInfo();
     // En la celda 'AE1' inserta el valor de V para marcar que ya ha sido enviada la informacion
    SpreadsheetApp.getActive().getSheetByName("DOCUMENTOS FACTURACION").getRange("AF1").setValue("V");
    SpreadsheetApp.getUi().alert("Informacion Enviada Exitosamente.");
  }else{SpreadsheetApp.getUi().alert("Ya se habia enviado la informacion anteriormente.")};
}

//////////////////////////////

function enviarProv() {
  if(SpreadsheetApp.getActive().getSheetByName("MOVS").getRange("X1").getValue() != "P"){
    OperFact.enviarProv();
     // En la celda 'AE1' inserta el valor de V para marcar que ya ha sido enviada la informacion
    SpreadsheetApp.getActive().getSheetByName("MOVS").getRange("X1").setValue("P");
    SpreadsheetApp.getUi().alert("Provisionadas, Dolares y Canceladas Enviadas Exitosamente.");
  }else{SpreadsheetApp.getUi().alert("Ya se habia enviado la informacion anteriormente.")};
}

//////////////////////////////

function fecha(){
  var rangeDF = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DOCUMENTOS FACTURACION").getRange("AE2");
  var rangeM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MOVS").getRange("W2");
  rangeDF.setValue(DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getDateCreated());
  rangeM.setValue(DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getDateCreated());
}

//////////////////////////////

function btns(){
  OperFact.upDocFact2();
  OperFact.upMovs2();
}
