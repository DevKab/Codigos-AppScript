const SSID = SpreadsheetApp.getActiveSpreadsheet().getId();

function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu("📑 CORREO");
  //menu.addSeparator();
  menu.addItem("📄 HACER CORTE DEL DÍA", "copiarOtroLibro");
  //menu.addSeparator();
  menu.addToUi();
}

function copiarOtroLibro() {
  const libroOrigen = SpreadsheetApp.openById(SSID); // hoja donde se copia la informacion
  const hojaOrigen = libroOrigen.getSheetByName("DB"); // nombre hoja donde se copia

  const libroDestino = SpreadsheetApp.openById(SSID); // id donde se va pegar
  const hojaDestino = libroDestino.getSheetByName("CONCENTRADO 2025"); // hoja donde se va pegar

  const ultimaFila = hojaOrigen.getLastRow();
  if (ultimaFila < 2) return; // Si no hay datos, termina

  // Lee todos los datos desde la fila 2
  const datos = hojaOrigen.getRange(2, 1, ultimaFila - 1, hojaOrigen.getLastColumn()).getValues();

  // Filtra solo las filas donde la primera columna tenga contenido
  const datosFiltrados = datos.filter(fila => fila[0] !== "" && fila[0] !== null);

  if (datosFiltrados.length === 0) return; // Si no hay filas con datos, termina

  // Inserta filas en la hoja destino antes de la fila 3
  hojaDestino.insertRowsBefore(3, datosFiltrados.length);

  // Pega los datos filtrados
  hojaDestino
    .getRange(3, 1, datosFiltrados.length, hojaOrigen.getLastColumn())
    .setValues(datosFiltrados);
}

function CopiaDiasLaborales() {
  const dia = new Date();
  const dias = [
    'domingo',
    'lunes',
    'martes',
    'miércoles',
    'jueves',
    'viernes',
    'sábado'
  ];
  const numeroDia = new Date(dia).getDay();
  const nombreDia = dias[numeroDia];

  if (nombreDia == "sábado" || nombreDia == "domingo") {

  } else {
    copiarOtroLibro();
  }
}
