// CORREO PROPIETARIO "ejemplo@correo.com" - {NOMBRE MESA}
function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu("📑 Reporte Correo");
  menu.addSeparator();
  menu.addItem("🔄| Actualizar datos ", "getData");
  menu.addSeparator();
  menu.addToUi();
}

//////////////////////////////

function getData() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
  const buscafecha = Utilities.formatDate(new Date(), "GMT-6", "yyyy-M-dd");
  const before = Utilities.formatDate(new Date(), "GMT-6", "yyyy-12-31")
  var query = "after:" + buscafecha + " before:" + before + "label:inbox";
  const cadenas = GmailApp.search(query);
  hoja.getRange("A1:H1").setBackgroundColor('#46bdc6').setFontColor('#ffffff')
  .setValues([["REMITENTE","ASUNTO","MENSAJE","FECHA","HORA","NOMBRE","MESA","TIEMPO"]]);
  cadenas.forEach(cadena => {
    const asunto = cadena.getFirstMessageSubject();
    const correo = cadena.getMessages()[0];
    const cuerpo = correo.getPlainBody();
    const remitente = correo.getFrom().split("<")[1];
    const corto = remitente;
    const fechahoy = Utilities.formatDate(new Date(), "GMT-6", "yyyy-M-dd");
    const fechaor = Utilities.formatDate(correo.getDate(), "GMT-6", "yyyy-M-dd");
    if (fechaor == fechahoy) {
      var date = Utilities.formatDate(correo.getDate(), "GMT-6", "yyyy-M-dd");
      var horaoriginal = Utilities.formatDate(correo.getDate(), "GMT-6", "h a");
    } else {
      var date = Utilities.formatDate(cadena.getLastMessageDate(), "GMT-6", "yyyy-M-dd");
      var horaoriginal = Utilities.formatDate(cadena.getLastMessageDate(), "GMT-6", "h a");
    }
    GmailApp.markMessagesRead(cadena.getMessages());
    const lectura = cadena.isUnread();
    const correos = [
      {CORREOS PROMOTORES}
    ];
    if (correos.includes(corto)) {
      const correoNombres = {
        {CORREOS CON NOMBRES DE PROMOTORES}
      }
      var maniana = ["12 AM","1 AM","2 AM","3 AM","4 AM","5 AM","6 AM","7 AM","8 AM","9 AM","10 AM","11 AM"];
      var mediodia = ["12 PM","1 PM","2 PM","3 PM"];
      var tarde = ["4 PM","5 PM","6 PM","7 PM"];
      var fuera = ["8 PM","9 PM","10 PM","11 PM"];
      var tiempo;
      (maniana.includes(horaoriginal))?tiempo="MAÑANA":0;
      (mediodia.includes(horaoriginal))?tiempo="MEDIO DIA":0;
      (tarde.includes(horaoriginal))?tiempo="TARDE":0;
      (fuera.includes(horaoriginal))?tiempo="FUERA DE HORARIO":0;
      if (date == buscafecha && lectura) {
        hoja.appendRow([corto, asunto, cuerpo, date, horaoriginal, correoNombres[corto], {NOMBRE DE MESA}, tiempo]);
      }
    }
  })
}

//////////////////////////////

function DelTable() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
  hoja.getRange("A2:J").clearContent();
}

//////////////////////////////

function CopiaDiasLaborales() {
  const dia = new Date();
  const dias = {
    0:'domingo',
    1:'lunes',
    2:'martes',
    3:'miércoles',
    4:'jueves',
    5:'viernes',
    6:'sábado'
  };
  const numeroDia = new Date(dia).getDay();
  const nombreDia = dias[numeroDia];

  if(nombreDia != "sábado" && nombreDia != "domingo") getData();
}
