function formulaBuscarBDir11() { // Activador a las 11am
  try {
    var hojasDatos = [
      { link: "1bndgKZFumEHaC9Ys8fGH3baMQz5wG4k-tswOyD1kAvg", nombreHoja: "S.Gastos Dir ANGIE" }, 
      { link: "1rJeu91j_SUVWCstdWDZzDd1yvx6U2oxO5FfS59XQAlg", nombreHoja: "S.Gastos Dir ALE" },
      { link: "100188c6jGaZsiKQ6gt1A7C4UcnyhDo5Y9mRwJV4E6sM", nombreHoja: "S.Gastos Dir PERLA" },
      { link: "14ie2zXXSMPmv2QPJbvLNPFdkNUyBMcAtakJxG2XPmXM", nombreHoja: "S.Gastos Dir RRHH" },//4
      { link: "1Ut23kn0z7VxnF5x2oJlojdBG1Mk49i91NNHf30uj3tM", nombreHoja: "S.Gastos Dir COBRANZA" },
      { link: "1msYek234jOIuZlLblqzYi67g67-IBgTp-cgtBLzP1Ko", nombreHoja: "S.Gastos Dir GABRIELA" },
      { link: "1URDyv9CLQQOMIwSniLmr-BnRq8z8UEgCxIs1Qr8VPh4", nombreHoja: "S.Gastos Dir YESSICA" },
      { link: "1qRXOJnoioNAD2mAITx9If1HDfNx0EfciooWQTkem2a8", nombreHoja: "S.Gastos Dir AZAEL" },
      { link: "1GAs64cBqewWRml9Ut3Q5zlFwaad0UUw5ZMxOKddyR7Q", nombreHoja: "S.Gastos Dir CARLOS" },
      { link: "1sRWZQOkyqOE46JhESDwe6EPMcgHVag4GVLB64NT8NF8", nombreHoja: "S.Gastos PROYECTOS" }
    ];

    hojasDatos.forEach(function (hoja) {
      try {
        copia10DirProCiclico(hoja.link, hoja.nombreHoja);
      } catch (error) {
        Logger.log(`Error procesando hoja con link ${hoja.link} y nombre ${hoja.nombreHoja}: ${error.message}`);
      }
    });
  } catch (error) {
    Logger.log(`Error general en limk12Archivos: ${error.message}`);
  }
}

function copia10DirProCiclico(link, nombreHoja){
  /*vectorTemporal */
  var libroOrigen = SpreadsheetApp.openById('18SOk6PCHpIxbL7oEfXK8MHnr8yzWGzJWNf_HYCmrmGk'); // Temporar idTesteo=1Rp7b_2B4IsvhesThR_4bXNRBuq21B8rSaA6g4AKe7Fk
  var hojaOrigen = libroOrigen.getSheetByName("SOLICITUD GASTOS TEMPORAL - CONCATENADO");
  var ultimaFilaOrigen = hojaOrigen.getLastRow();
  var ultimaColOrigen = hojaOrigen.getLastColumn();

  if (ultimaFilaOrigen === 0 || ultimaColOrigen < 42) {
    Logger.log("Hoja origen vacía o con menos de 42 columnas.");
    return;
  }

  var datoOrigenTemporal = hojaOrigen.getRange(1, 1, ultimaFilaOrigen, 42).getValues();

  var libroDestino = SpreadsheetApp.openById(link);
  var hojaDestino = libroDestino.getSheetByName(nombreHoja);
  var ultimaFilaDestino = hojaDestino.getLastRow();
  var ultimaColDestino = hojaDestino.getLastColumn();

  if (ultimaFilaDestino === 0 || ultimaColDestino < 42) {
    Logger.log("Hoja destino vacía o con menos de 42 columnas.");
    return;
  }

  var datosDestino10Dir = hojaDestino.getRange(1, 1, ultimaFilaDestino, 42).getValues();

  Logger.log("Nombre de la hoja destino: " + nombreHoja);

  for(var i = 0; i < datosDestino10Dir.length; i++){
    if(datosDestino10Dir[i][28] === "EN PROCESO"){
      for(var j = 0; j < datoOrigenTemporal.length; j++){
        if(datoOrigenTemporal[j][0] && datosDestino10Dir[i][0] && datoOrigenTemporal[j][0] === datosDestino10Dir[i][0]){
          var datosAC_AJ = datoOrigenTemporal[j].slice(28,36);
          hojaDestino.getRange(i+1, 29, 1, datosAC_AJ.length).setValues([datosAC_AJ]);
        }
      }
    }
  }
}
