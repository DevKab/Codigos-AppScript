function formulaBuscarBDir11() { // Activador a las 11am
  try {
    var hojasDatos = [
      { link: "1c1uWh2SGymRdv33kG3S4XGt-UDCAujs6gWRCblJzR0I", nombreHoja: "S.Gastos Dir ANGIE" }, //1 //LOS DEMAS ME FALTAN
      { link: "13eP0FC42Sglo6cr6kCUnd9jDUFRyDOqEmbfXrqhigv4", nombreHoja: "S.Gastos Dir ALE" },//3
      { link: "13RHjR0oWDoOIHiU3mtuR6rDLgQ2RRK4tjE1xVusTACU", nombreHoja: "S.Gastos Dir PERLA" },//4
      { link: "1kM7LCRFoOJReKUd0qqd0RVo6YB03ls1de2rHDsQ7XEg", nombreHoja: "S.Gastos Dir RRHH" },//5
      { link: "12K_YCoHRCit9iLrbAGuInQKuJGFI8Xinmv4cSCCGcXA", nombreHoja: "S.Gastos Dir COBRANZA" },//6
      { link: "1JMgv8TqIwD2LoEAg6KTYu8IDpCnc20CkFC1vS8CQ86s", nombreHoja: "S.Gastos Dir GABRIELA" },//7
      { link: "1dKi59V6zPD4q-GuKDvRX0XJo4IDsWa4QnWcO9vu3b2g", nombreHoja: "S.Gastos Dir YESSICA" },//8
      { link: "1NAIUKXsHgeNU-d0rUlR-9xGF6kEp-3geLJYAbCnVauw", nombreHoja: "S.Gastos Dir AZAEL" },//9
      { link: "1ut-WpFn3Zch55b7_uHsbUF0fv646o71GYt9gpnhqL9Q", nombreHoja: "S.Gastos Dir CARLOS" },//10
      { link: "1C78qxfla-rxrNRUwCiLwD5oY8vtlaByNJGDXp1O3nLM", nombreHoja: "S.Gastos PROYECTOS" }//12 //MODIFICADO
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
  var libroOrigen = SpreadsheetApp.openById('1mPIaW4vio2y5lM29AgnlSl_Y10ZkajZvqpIW6CibRmg'); // Temporar idTesteo=1Rp7b_2B4IsvhesThR_4bXNRBuq21B8rSaA6g4AKe7Fk
  var hojaOrigen = libroOrigen.getSheetByName("SOLICITUD GASTOS TEMPORAL - CONCATENADO");
  
  // Obtener los valores de la hoja origen
  var datoOrigenTemporal = hojaOrigen.getRange("A:AP").getValues();// de A:AO a A:AP


  /*optener los datos de 10 dir / pro / ciclico */
  var libroDestino = SpreadsheetApp.openById(link); // Master idTesteoV1 = 1N12NZmKe0JjWuFVtww2C4Xm52E9XMZyRgx00vQXvRL0 = 10Dir
  var hojaDestino = libroDestino.getSheetByName(nombreHoja);

  // Obtener los valores de la hoja origen
  var datosDestino10Dir = hojaDestino.getRange("A:AP").getValues();// de A:AO a A:AP //10Dir
  Logger.log("Nombre de la hoja destino: " + nombreHoja);

  for(var i = 0; i < datosDestino10Dir.length; i++){ //dir
    if(datosDestino10Dir[i][28] === "EN PROCESO"){//dir
      for(var j = 0; j < datoOrigenTemporal.length; j++){
        if(datoOrigenTemporal[j][0] && datosDestino10Dir[i][0] && datoOrigenTemporal[j][0] === datosDestino10Dir[i][0]){
          
            //copiar del temporal AC-AJ 
            var datosAC_AJ = datoOrigenTemporal[j].slice(28,36); 
            //pegar 10Dir
            hojaDestino.getRange(i+1, 29, 1, datosAC_AJ.length).setValues([datosAC_AJ]);//pegar en la hojas correspondiente.
            
          
        }
      }
    }
  }
}
