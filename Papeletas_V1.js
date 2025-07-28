  // ID's de archivos 5R y 10R
const AR_5RID = "";
const AR_10RID = "";

  // Nombre de Pestaña de Papeletas
const SH_PAPELETAS = "Prueba";

//////////////////////////////

function papeletasInfoDir(SSID, SHEET) {
  const hoja = SpreadsheetApp.openById(SSID).getSheetByName(SHEET);
  const ultimaFila = hoja.getLastRow();
  const cantidadFilas = ultimaFila - 5;
  const datos = hoja.getRange(6, 1, cantidadFilas, hoja.getLastColumn()).getValues();
  const datosFiltrados = datos.filter((fila, index) => {
    return (fila[17] == "EFECTIVO" && fila[28] == "NUEVO");
  });
  var identificador = 0;    // A
  var fechaCaptura = 1;     // B
  var quienSolEmail = 2;    // C
  var dptoSol = 3;          // D
  var areaCliente = 6;      // G
  var desc = 15;            // P
  var formaPago = 17;       // R
  var comentarios = 19;     // T
  var titularContact = 22;  // W
  var monto = 23;           // X
  const agrupados = {};
  const email = {
    "ANGEL_PULIDO": "consultoria.administrativa@kabzo.org"
  };

  datosFiltrados.forEach(fila => {
    const id = fila[identificador];
    const nuevaDesc = fila[desc]?.trim() || "";
    const nuevoTexto = fila[comentarios]?.trim() || "";
    if (!agrupados[id]) {
      agrupados[id] = {
        identificador: id,
        asesor: "GASTOS",
        fechaCaptura: fila[fechaCaptura],
        envio: "DIRECCION",
        pr: "GASTOS",
        quienSolEmail: email[fila[quienSolEmail]],
        dptoSol: fila[dptoSol],
        areaCliente: fila[areaCliente],
        vacio: "",
        descs: new Set(nuevaDesc ? [nuevaDesc] : []),
        textos: new Set(nuevoTexto ? [nuevoTexto] : []),
        formaPago: fila[formaPago],
        titularContact: fila[titularContact],
        monto: (parseFloat(fila[monto])*-1) || 0
      };
      } else {
        agrupados[id].monto += (parseFloat(fila[monto]) * -1) || 0;
        if (nuevaDesc) agrupados[id].descs.add(nuevaDesc);
        if (nuevoTexto) agrupados[id].textos.add(nuevoTexto);
      }
  });

  const salida = Object.values(agrupados).map(obj => [
    obj.asesor,
    obj.fechaCaptura,
    obj.envio,
    obj.pr,
    obj.quienSolEmail,
    obj.areaCliente,
    obj.titularContact,
    obj.vacio,
    obj.vacio,
    [...obj.descs, ...obj.textos].join(" // "),
    obj.vacio,
    Math.round(obj.monto)
  ]);
  if (salida.length === 0) {
    SpreadsheetApp.getUi().alert("No hay datos para escribir.");
    return;
  }
  const hojaDestino = SpreadsheetApp.openById(SSID).getSheetByName(SH_PAPELETAS);
  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert("La hoja de Papeletas no existe.");
  }
  const columnasPorBloque = 14; // 14 columnas: B (2) hasta AO (41) en saltos de 3
  const saltoFilas = 25;        // Cada bloque nuevo baja 25 filas
  const columnasEspaciadas = 3; // Espacio entre columnas

  salida.forEach((colData, colIndex) => {
    const bloque = Math.floor(colIndex / columnasPorBloque);
    const posicionEnBloque = colIndex % columnasPorBloque;

    const filaDestino = 1 + bloque * saltoFilas;
    const columnaDestino = 2 + posicionEnBloque * columnasEspaciadas;

    hojaDestino.getRange(filaDestino, columnaDestino, colData.length, 1).setValues(
      colData.map(valor => [valor])
    );
  });
}

//////////////////////////////

function customTranspose(matrix) {
  return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));
}

//////////////////////////////

function eraseColumns(SSID){
  const hojaDestino = SpreadsheetApp.openById(SSID).getSheetByName(SH_PAPELETAS);
  for(i=0;i<14;i++){
    hojaDestino.getRange(1, (2+i*3),200,1).clearContent();
  }
}
