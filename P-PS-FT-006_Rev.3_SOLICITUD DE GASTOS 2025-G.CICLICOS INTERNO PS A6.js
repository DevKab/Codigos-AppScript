/*esta conectado con una herencia del archivo Temporal*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var mensaje = "Recuerda que esta plantilla contiene listas anidadas y recibe información de otros archivos:"
    + "\n- 🚫 No agregar o quitar columnas y filas."
    + "\n- 🚫 No alterar fórmulas."
    + "\n- 🚫 No modificar la posición de las tablas o el rango."
    + "\n- ✅ Para un uso adecuado del archivo consulta tu instrucción de trabajo P-PS-IT-002_ SOLICITUD DE GASTOS DESPACHO DIRECCIÓN SOLICITANTE"
    + "\n- ☎︎ Contacta a 'Optimización' para realizar modificaciones. V11";
  ui.alert(mensaje);
  
    ui.createMenu('📅 | Backup')
      .addItem('1. Envio de informacion Concentrado  | 📄', 'master')
      .addToUi();
}

function master(){
  CiclicoMaster.llamado();
  CiclicoMaster.bloqueo();
}

function bloqueo(){ //activador
 CiclicoMaster.bloqueo()
}

function onEdit(e){ //funciona
    var hojaActiva = e.source.getActiveSheet();
    var nombreHoja = hojaActiva.getName();
  
    // Verificar si la hoja activa es "SOLICITUD GASTOS TEMPORAL - CONCATENADO"
    if (nombreHoja !== "S.Gastos CICLICOS INTERNO PS A6") return;
  
    var rangoEditado = e.range;
    var filaEditada = rangoEditado.getRow();
    var columnaEditada = rangoEditado.getColumn();
  
    const columnaR = 18; // Columna AC 29 (FORMA DE PAGO) 
    const columnaAC = 29; // Columna AC 29 (estatus)
    const columnaAD = 30; // Columna AD 30 (fecha de pago)
    const columnaAE = 31; // Columna AD 30 (LINK COMPROBANTE/FOLIO)
    const columnaAJ = 36; // Columna AJ 36 mes de pago
  
    // Si la edición ocurrió en la columna AE
    if (columnaEditada === columnaAE) {
        var valorCeldaR = hojaActiva.getRange(filaEditada, columnaR).getValue(); // Obtener lA FORMA DE PAGO
        var valorCeldaAC = hojaActiva.getRange(filaEditada, columnaAC).getValue(); // Obtener el estatus
        var valorCeldaAD = hojaActiva.getRange(filaEditada, columnaAD).getValue(); // Obtener la fecha de pago
        var valorCeldaAJ = hojaActiva.getRange(filaEditada, columnaAJ).getValue(); // Obtener el mes de pago
  
        // Si el FOMATO DE PAGO es "CHEQUES" o "TARJETA" o "CARGO_BANCARIO" o "SALDOS"
        if (valorCeldaR === "CARGO_BANCARIO" || valorCeldaR === "CHEQUE CERTIFICADO" || valorCeldaR === "CHEQUE DE CAJA" || valorCeldaR === "CHEQUE" || valorCeldaR === "EFECTIVO" 
            || valorCeldaR === "EFECTIVO_DOLARES" ||  valorCeldaR === "ORDEN DE PAGO" ||  valorCeldaR === "PAGO_CIE" || valorCeldaR === "PAGO_EN_LINEA" ||  valorCeldaR === "SALDOS"
            || valorCeldaR === "TARJETA DE CREDITO"|| valorCeldaR === "TARJETA DE DEBITO" || valorCeldaR === "TRANSFERENCIA" ) {
            // Si la columna AD o AJ está vacía
           if (valorCeldaAC === "EN PROCESO") {
              if (valorCeldaAD === "" || valorCeldaAD === null || valorCeldaAJ === "" || valorCeldaAJ === null) {
                    // Obtener el mes actual
                    var today = new Date();
                    var meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
                    var mesActual = meses[today.getMonth()]; // Obtener el mes actual en formato de texto
                    var fomateoToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');
                    var status = "PAGADO Y COMPROBANTE EN CARPETA";
      
                    // Escribir la fecha de hoy en la columna AD (fecha de pago) y el mes en la columna AJ
                    hojaActiva.getRange(filaEditada, columnaAC).setValue(status);
                    hojaActiva.getRange(filaEditada, columnaAD).setValue(fomateoToday);
                    hojaActiva.getRange(filaEditada, columnaAJ).setValue(mesActual);
                    Logger.log("Fecha de pago y mes actualizados en fila " + filaEditada);
              }
           }
        }
        
    }
    // Obtener el valor de la celda AC para esta fila
    var valorCeldaACCheck = hojaActiva.getRange(filaEditada, columnaAC).getValue();
    if (columnaEditada === columnaAC && (valorCeldaACCheck === "CANCELADO" || valorCeldaACCheck === "RECHAZADO")) {
        // Obtener la fecha actual
        var today = new Date();
        var fomateoToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');

        // Escribir la fecha de hoy en la columna AD (fecha de pago)
        hojaActiva.getRange(filaEditada, columnaAD).setValue(fomateoToday);
        // Limpiar el valor de la columna AJ (mes de pago)
        hojaActiva.getRange(filaEditada, columnaAJ).setValue("");
        Logger.log("Fecha de pago actualizada y mes limpiado en fila " + filaEditada);
    }
}
