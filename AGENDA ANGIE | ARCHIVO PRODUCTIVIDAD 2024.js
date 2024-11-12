function onEdit(e) {
  var celdaActiva = e.range;
  var filaActiva = celdaActiva.getRow();
  var hojaActiva = celdaActiva.getSheet();

  var colDeseada = 7; // columna G
  var filaNodeseada = 1;
  var hojaDeseada = "PRODUCTIVIDAD"; // cambie el nombre a la hoja que desee

  //31/07/2024
  var hojaEva = "EVA"; // cambie el nombre a la hoja que desee
  var hojaAngie = "ANGIE"; // cambie el nombre a la hoja que desee
  var hojaAngel = "ANGEL"; // cambie el nombre a la hoja que desee

  if (celdaActiva.getColumn() == colDeseada && celdaActiva.getRow() != filaNodeseada && hojaActiva.getName() == hojaDeseada) {
    var email = Session.getActiveUser().getEmail();
    var targetSheet = hojaActiva.getParent().getSheetByName(hojaDeseada);
    targetSheet.getRange(filaActiva, 3).setValue(email); // insertar el correo en la columna B de la fila activa
  }else if(celdaActiva.getColumn() == colDeseada && celdaActiva.getRow() != filaNodeseada && hojaActiva.getName() == hojaEva){
    var email = Session.getActiveUser().getEmail();
    var targetSheet = hojaActiva.getParent().getSheetByName(hojaEva);
    targetSheet.getRange(filaActiva, 3).setValue(email); // insertar el correo en la columna B de la fila activa
  }else if(celdaActiva.getColumn() == colDeseada && celdaActiva.getRow() != filaNodeseada && hojaActiva.getName() == hojaAngie){
    var email = Session.getActiveUser().getEmail();
    var targetSheet = hojaActiva.getParent().getSheetByName(hojaAngie);
    targetSheet.getRange(filaActiva, 3).setValue(email); // insertar el correo en la columna B de la fila activa
  }else if(celdaActiva.getColumn() == colDeseada && celdaActiva.getRow() != filaNodeseada && hojaActiva.getName() == hojaAngel){
    var email = Session.getActiveUser().getEmail();
    var targetSheet = hojaActiva.getParent().getSheetByName(hojaAngel);
    targetSheet.getRange(filaActiva, 3).setValue(email); // insertar el correo en la columna B de la fila activa
  }

  updateTimestamp(e) //Esta función actualiza una marca de tiempo en una celda cuando se edita una celda específica.
}

function BloqueaUsuariosA_I() {
  var celdaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  var filaActiva = celdaActiva.getRow();
  var hojaActiva = celdaActiva.getSheet();

  var colDeseada = 10;
  var filaNodeseada = 13;
  var nombrehojadeseada = "PRODUCTIVIDAD";

  if (celdaActiva.getColumn() == colDeseada && celdaActiva.getRow() != filaNodeseada && celdaActiva.getSheet().getName() == nombrehojadeseada) {
    var prot = hojaActiva.getRange(filaActiva, 1, 1, 9).protect().setDescription("BLOCK a-i");
    var propietario = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
    var editoresActuales = prot.getEditors();

    // Eliminar todos los editores excepto el propietario
    editoresActuales.forEach(function(editor) {
      if (editor != propietario) {
        prot.removeEditor(editor);
      }
    });
    if (prot.canDomainEdit()) {
      prot.setDomainEdit(false);
    }
  }
}


function updateTimestamp(e) {
    var sheet = e.source.getActiveSheet();
    var editedRange = e.range;
    var editedColumn = editedRange.getColumn();
    
    // Especifica la columna que quieres monitorear (por ejemplo, columna 9 para la columna I)
    var monitoredColumn = 9; // Columna I
    
    // Verifica si la columna editada es la que estamos monitoreando
    if (editedColumn == monitoredColumn) {
        var row = editedRange.getRow();
        var timestampCell = sheet.getRange(row, editedColumn + 1); // La celda a la derecha de la editada
        timestampCell.setValue(new Date());
    }
}

