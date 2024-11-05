function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(' ➠ Transferir Datos')
    .addItem('Iniciar Envío', 'enviarInfo')
    .addToUi();
}


function enviarInfo() {
  const archivosDestino = {
    "ANDRES": "1HiAhAx5nK8TB41_rZAQgdQRtkOWFt7ZxnQFZlhhNJBg",
    "CARSO": "1jfRJV9sg8paf8zPP-EL_wyt9wz2wRI_Ms_sEqRii_LU",
    "MARTIN": "1ObmQDydnOI-phK5u5T-xKq_jbXIf8ONgbobiLGw3TqM",
    "MIKE": "1TeOn9z_Q9LDbI6Aa1OAnkGMShgEXw7VO9laqY8U7HUc",
    "REYNALDO": "1xpUld3tyQ6D6LIeJkimaDcktq2MqGi-ZueUvpJh6uls",
    "VLADIMIR": "1F57_4tFJ-n-6rQmcGZ0YC9jrUvtGoJnD606nappJ4VY"
  };

  const hojaOrigen = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const datos = hojaOrigen.getDataRange().getValues();
  const encabezado = datos[0]; // Guardamos el encabezado
  const datosFiltrados = {}; // Objeto para agrupar las filas por destino

  // Agrupamos las filas según el valor de la columna S
  datos.slice(1).forEach(fila => { // Ignora el encabezado
    const nombre = fila[18]; // Columna S (índice 18)
    if (archivosDestino[nombre]) {
      if (!datosFiltrados[nombre]) {
        datosFiltrados[nombre] = [encabezado]; // Agrega el encabezado para cada destino
      }
      datosFiltrados[nombre].push(fila); // Agrega la fila al grupo correspondiente
    }
  });

  // Enviar datos agrupados a cada archivo de destino
  Object.keys(datosFiltrados).forEach(nombre => {
    const ssDestino = SpreadsheetApp.openById(archivosDestino[nombre]);
    const hojaDestino = ssDestino.getSheets()[0];
    hojaDestino.clear(); // Limpia la hoja antes de pegar los datos nuevos
    hojaDestino.getRange(1, 1, datosFiltrados[nombre].length, datosFiltrados[nombre][0].length).setValues(datosFiltrados[nombre]);
  });

  SpreadsheetApp.getUi().alert('Datos enviados, favor de verificar.');
}
