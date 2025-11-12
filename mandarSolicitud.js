const SSID_A2 = "1zc6xunIz8J3B52QVu5sXWYkkxV6VhCq_SKJO_vBgpwM";

function mandarSolicitudBoton(){
  mandarSolicitud();
  // rebajes();
  // delTable();
}

function mandarSolicitud() {
  var nominaHoja = SpreadsheetApp.openById(SSID).getSheetByName(TABLAS_SHEET);

  var nominaSemanal = (nominaHoja.getRange(12,1,158,11).getValues())
    .map(fila => 
      [fila[0], fila[1], fila[2], fila[3], fila[4]]
    ).slice(1)
  var nomSemFiltro = nominaSemanal.filter(fila => fila[0] !== "" && fila[0] !== null);

  var nominaBonos = (nominaHoja.getRange(12,1,158,11).getValues())
    .map(fila => 
      [fila[6], fila[7], fila[8], fila[9], fila[10]]
    ).slice(1);
  var nomBonosFiltro = nominaBonos.filter(fila => fila[0] !== "" && fila[0] !== null);

  const filtroInferiores = (nominaHoja.getRange(1001,1,600,25).getValues())
    .filter(fila => 
      fila[2] !== "" && fila[2] !== null && fila[2] !== "QUIEN SOLICITA"
    );

  const nominaInferior = filtroInferiores
    .filter(fila => 
      fila[8] == "NOMINA"
    );
  const bonosBInferior = filtroInferiores
    .filter(fila => 
      fila[8] == "BONO LEALTAD" || fila[8] == "BONO DESPENSA" || fila[8] == "BONO TRANSPORTE"
    );
  const kpiInferior = filtroInferiores
    .filter(fila => 
      fila[8] == "BONO MENSUAL"
    );
  const ingeMesInferior = filtroInferiores
    .filter(fila => 
      fila[8] == "EMPLEADO DEL MES"
    );
  var solicitudInferior = nominaInferior.map(fila => fila.slice());

  // Logger.log(nomSemFiltro);
  // return

  var nominaSuperior = [[]];
  if(nomBonosFiltro.length !== 0 && nomBonosFiltro[0][1] == "BONO LEALTAD"){
    nominaSuperior = [...nomSemFiltro, ...nomBonosFiltro];
    solicitudInferior = [...nominaInferior, ...bonosBInferior];
    Logger.log(nominaSuperior);
  } else if(nomBonosFiltro.length !== 0 && nomBonosFiltro[0][1] == "BONO MENSUAL"){
    nominaSuperior = [...nomSemFiltro, ...nomBonosFiltro];
    solicitudInferior = [...nominaInferior, ...kpiInferior];
    Logger.log(nominaSuperior);
  } else if(nomBonosFiltro.length !== 0 && nomBonosFiltro[0][1] == "EMPLEADO DEL MES"){
    nominaSuperior = [...nomSemFiltro, ...nomBonosFiltro];
    solicitudInferior = [...nominaInferior, ...ingeMesInferior];
    Logger.log(nominaSuperior);
  }

    // const arr3 = arr1.map((fila, i) => [...fila, ...(arr2[i] || [])]); Concatenar columnas
  var filasSolicitud = solicitudInferior.map((fila, i) => [...fila, ...(nominaSuperior[i] || [])]);

  // Logger.log(filasSolicitud);
  // return

  var sheet12 = SpreadsheetApp.openById(SSID_A2).getSheetByName("S.Gastos CICLICOS INTERNO PS A2");
  var lastRow12 = sheet12.getRange(1,3).getDataRegion().getLastRow()+1;
  var value = (sheet12.getRange(lastRow12-1,1).getValue()).substring(8,12);
  var num = value*1;

  var numStr = "";  // GENERACION DE IDENTIFICADOR
  if (num>=0&&num<9){
    num = (value.substring(3)*1)+1;
    numStr = `D000`+num;
  } else if (num>=9&&num<99){
    num = (value.substring(2)*1)+1;
    numStr = `D00`+num;
  } else if (num>=99&&num<999){
    num = (value.substring(1)*1)+1;
    numStr = `D0`+num;
  } else if (num>=999&&num<9999){
    num = (value.substring(0)*1)+1;
    numStr = `D`+num;
  }

  // Logger.log(numStr);
  // return

  var solicitudNomina = filasSolicitud
    .map(fila => 
      [`PS-A2-S${numStr}D`,  // IDENTIFICACION
      (new Date()).toISOString().substring(0,10),  // FECHA CAPTURA
      fila[2],          // QUIEN SOLICITA
      fila[3],          // DPTO SOLICITANTE
      fila[4],          // USO
      fila[5],          // PERIODICIDAD
      fila[6],          // ÁREA DONDE APLICA
      fila[7],          // CATEGORIA
      fila[8],          // SUBCATEGORIA
      "N/A",            // DETALLE
      fila[10],         // USUARIO FINAL
      fila[27],         // CANTIDAD (TABLAS SUPERIORES)
      fila[12],         // PRECIO UNITARIO
      fila[13],         // MARCA
      fila[14],         // PROVEEDOR
      fila[15],         // DESCRIPCION
      fila[16],         // CATEGORÍA GASTOS
      "TRANSFERENCIA",  // FORMA DE PAGO
      "NACIONAL",       // DETALLE DE PAGO
      "N/A",            // COMENTARIOS DE ENTREGA
      fila[20],         // DESTINO
      fila[21],         // CUENTA_CLABE
      fila[22],         // TITULAR
      fila[29],         // MONTO / IMPORTE (TABLAS SUPERIORES)
      fila[24]          // AUTORIZO
    ])
  sheet12.getRange(lastRow12,1,solicitudNomina.length, solicitudNomina[0].length).setValues(solicitudNomina);
}

function rebajes(){
  var rebajesSemana = nominaHoja.getRange("N3").getValue();
  var rebajesTotalViejo = nominaHoja.getRange("P3").getValue();
  var rebajesTotalNuevo = rebajesTotalViejo+rebajesSemana;
  nominaHoja.setValue(rebajesTotalNuevo);
}
