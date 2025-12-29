// const SSID_A2 = ""; // REAL
const SSID_A2 = SpreadsheetApp.getActiveSpreadsheet().getId(); // PRUEBA
const ID_NOMINAS = SpreadsheetApp.getActiveSpreadsheet().getId(); // Archivo A2
const ID_DIRECTORIO = `1NZBsJOLjnP6aojinaPUaLMnliDHYnNqVJKMYq8VhTJE`; // V0.2
const ID_MASTER_GASTOS = `178M33EaTbv6rT6CA2XkA_csJlMoBI9Ej3s1T_7hq0no`;

function mandarSolicitudBoton(){
  if(mandarSolicitud()){
    rebajes();
    delTable();
    return
  }
  
}

function mandarSolicitud() {
  var nominaHoja = SpreadsheetApp.openById(SSID).getSheetByName(TABLAS_SHEET);
  var nominaSupCompleta = nominaHoja.getRange(12,1,200,16).getValues();
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`Formato Nomina Ejemplo`);
  if (hoja.getRange(`A13`).getValue() == `` || hoja.getRange(`A13`).getValue() == null){
    SpreadsheetApp.getActiveSpreadsheet().toast(`🔴NO HAY DATOS POR MANDAR.🔴`);
    return false
  }
  if ((hoja.getRange(`G13:G300`).getValues().filter(fila => fila[0] === `NOMBRE`)).flat().length === 1 ){
    SpreadsheetApp.getActiveSpreadsheet().toast(`🟡SELECCIONA UN EMPLEADO DEL MES.🟡`);
    return false
  }
  var fechaDinamica = new Date();

  (hoja.getRange(`R5`).getValue())?fechaDinamica = hoja.getRange(`R2`).getValue():0;

  var nominaSemanal = (nominaSupCompleta).filter(fila => 
    fila[0] !== "" && fila[0] !== null && fila[3] !== `HORAS EXTRAS`)
    .map(fila => 
      [fila[0], fila[1], fila[2], `NOMINA`, fila[4]] 
    ).slice(1).filter(fila => 
    fila[0] !== "" && fila[0] !== null && fila[3] !== `HORAS EXTRAS`);
  
  var horasExtra = (nominaSupCompleta).slice(1).filter(fila =>   //  Cambiar el intervalo a la tercer tabla
    fila[12] !== "" && fila[12] !== null && fila[12] !== `AL DARLE CLICK A "MANDAR SOLICITUD" ESTA CONFIRMANDO QUE LOS DATOS SON CORRECTOS.`)
    .map(fila => 
      [fila[12], fila[13], fila[14], `HORAS EXTRAS`, fila[15]] 
    );

  var nominaBonos = (nominaSupCompleta)
    .map(fila => 
      [fila[6],fila[8],1,fila[7], fila[10]]
    ).slice(1).filter(fila => 
    fila[0] !== "" && fila[0] !== null && fila[4] !== 0);

  var solicitudSuperior = [...nominaSemanal, ...horasExtra, ...nominaBonos];

  var datos = nominaHoja.getRange(1002,1,95,29).getValues()
    .filter(fila => fila[10] != `` && fila[10] != null);
  var datosObj = personaObject(datos);
  var today = Utilities.formatDate(fechaDinamica,Session.getScriptTimeZone(), `dd/MM/yyyy`);

  // Logger.log(JSON.stringify(datosObj));
  // return;
  const a2Sheet = SpreadsheetApp.openById(ID_NOMINAS).getSheetByName(`S.Gastos CICLICOS INTERNO PS A2`);
  var consecutivo = a2Sheet.getRange(`A6:A`).getValues();
    // .filter(fila => typeof fila[0] === `string` && fila[0].startsWith(`${numArea}-${archivo}-${numEmpleado}-${numSubcatego}`)).flat()).length+1;

  //   var datosObjStr = JSON.stringify(datosObj);
  // Logger.log(`
  // ${datosObjStr}
  // `);
  // return false

  solicitudSuperior = solicitudSuperior.map(fila => [
    generarIdentificador( // IDENTIFICADOR
      datosObj[fila[0]].AREA_APLICA,
      datosObj[fila[0]].CATEGORIA,
      fila[3],
      fila[0],
      consecutivo) || `SIN DATOS`,
    today, // FECHA CAPTURA
    datosObj[fila[0]].QUIEN_SOL || `SIN DATOS`, // QUIEN SOLICITA
    datosObj[fila[0]].DPTO_SOL || `SIN DATOS`, // DPTO SOLICITANTE
    datosObj[fila[0]].USO || `SIN DATOS`, // USO
    ((fila[3]==`NOMINA`)?`SEMANAL`:(fila[3]==`HORAS EXTRAS`)?`ÚNICO`:`MENSUAL`) || `SIN DATOS`, // PERIODICIDAD
    datosObj[fila[0]].AREA_APLICA || `SIN DATOS`, // ÁREA DONDE APLICA
    datosObj[fila[0]].CATEGORIA || `SIN DATOS`, // CATEGORIA
    fila[3] || `SIN DATOS`, // SUBCATEGORIA
    datosObj[fila[0]].DETALLE || `SIN DATOS`, // DETALLE
    fila[0] || `SIN DATOS`, // USUARIO FINAL
    fila[2] || `SIN DATOS`, // CANTIDAD
    (-1*fila[1]) || `SIN DATOS`, // PRECIO UNITARIO
    `N/A`, // MARCA
    `N/A`, // PROVEEDOR
    (fila[3]!=`NOMINA`&&fila[3]!=`HORAS EXTRAS`)?mesNomina(fila[3],today):semanaDelMesNominaSemanal(fila[3],today) || `SIN DATOS`, // DESCRIPCION
    `SERVICIO`, // CATEGORÍA GASTOS
    `TRANSFERENCIA`, // FORMA DE PAGO
    `NACIONAL`, // DETALLE DE PAGO
    `N/A`, // COMENTARIOS DE ENTREGA
    datosObj[fila[0]].DESTINO || `SIN DATOS`, // DESTINO
    datosObj[fila[0]].CUENTA_CLABE || `SIN DATOS`, // CUENTA_CLABE
    datosObj[fila[0]].TITULAR || `SIN DATOS`, // TITULAR
    (-1*fila[4]) || `SIN DATOS`, // MONTO / IMPORTE
    `AZAEL_RANGEL`,
    `N/A`,
    `SIN TICKET`,
    `N/A`
  ])
  

  var sheet13 = SpreadsheetApp.openById(SSID_A2).getSheetByName(`S.Gastos CICLICOS INTERNO PS A2`);
  var lastRow13 = (sheet13.getRange(`C1:C`).getValues().filter(fila => fila[0]!="").flat()).length+3;

  // Logger.log(`Arreglo concat:
  // ${solicitudSuperior}`);

    //  Insertar datos en A2
  sheet13.getRange(lastRow13,1,solicitudSuperior.length,solicitudSuperior[0].length).setValues(solicitudSuperior);
  // return false
  return true
}


function semanaDelMesNominaSemanal(subcatego,fechaStr) {
  const [dia, mes, anio] = fechaStr.split('/').map(n => parseInt(n, 10));
  const fecha = new Date(anio, mes - 1, dia);
  const PRIMER_DIA_SEMANA = 5;
  function obtenerViernesSemana(fecha) {
    const d = new Date(fecha);
    const day = d.getDay();
    const diff = (PRIMER_DIA_SEMANA - day + 7) % 7;
    d.setDate(d.getDate() + diff);
    return d;
  }
  const viernesSemana = obtenerViernesSemana(fecha);
  const mesViernes = viernesSemana.getMonth();
  const anioViernes = viernesSemana.getFullYear();
  const inicioMes = new Date(anioViernes, mesViernes, 1);
  const offset = (inicioMes.getDay() - PRIMER_DIA_SEMANA + 7) % 7;
  const numeroDia = viernesSemana.getDate();
  const semana = Math.floor((numeroDia + offset - 1) / 7);
  var numSemana = ``;
  switch (semana){
    case 1: numSemana = `1RA`; break;
    case 2: numSemana = `2DA`; break;
    case 3: numSemana = `3RA`; break;
    case 4: numSemana = `4TA`; break;
    default: numSemana = `5TA`; break;
  }
  const meses = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
  ];

  var mesNombre = meses[mesViernes] || ``;
  return (`${subcatego} ${numSemana} SEMANA ${mesNombre} ${anioViernes}`);
}

function mesNomina(subcatego,fechaStr) {
  const [dia, mes, anio] = fechaStr.split('/').map(n => parseInt(n, 10));
  const fecha = new Date(anio, mes - 1, dia);
  const PRIMER_DIA_SEMANA = 5;
  function obtenerViernesSemana(fecha) {
    const d = new Date(fecha);
    const day = d.getDay();
    const diff = (PRIMER_DIA_SEMANA - day + 7) % 7;
    d.setDate(d.getDate() + diff);
    return d;
  }
  const viernesSemana = obtenerViernesSemana(fecha);
  const mesViernes = viernesSemana.getMonth();
  const anioViernes = viernesSemana.getFullYear();
  const meses = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
  ];

  var mesNombre = meses[mesViernes] || ``;
  return (`${subcatego} ${mesNombre} ${anioViernes}`);
}

//////////////////////////////

function rebajes(){
  const nominaHoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rebajesSemana = (nominaHoja.getRange("M3").getValue())*1;
  var rebajesTotalViejo = nominaHoja.getRange("O3").getValue()*1;
  var rebajesTotalNuevo = (rebajesTotalViejo+rebajesSemana)*1;
  // Logger.log(rebajesTotalNuevo);
  nominaHoja.getRange(`O3`).setValue(rebajesTotalNuevo);
}

//////////////////////////////

function generarIdentificador(area,categoria,subcatego,nombre,consecutivo){
// function generarIdentificador(){
  //  const ID_DIRECTORIO = `1NZBsJOLjnP6aojinaPUaLMnliDHYnNqVJKMYq8VhTJE`; // V0.2
  //  const ID_MASTER_GASTOS = `178M33EaTbv6rT6CA2XkA_csJlMoBI9Ej3s1T_7hq0no`;
  //  const ID_NOMINAS = SpreadsheetApp.getActiveSpreadsheet().getId(); // Archivo A2
  //  const subcatego = `NOMINA`;
  //  const nombre = `SARAI BELLO ALBARRAN`;
  //  const area = `PROYECTOS`;
  //  const categoria = `NOMINAS`;
  // try{
   var archivo;
   (categoria == `NOMINAS`)?archivo = `A2`:0; // SWITCH para obtener que archivo es con la categoria
   const directorioSheet = SpreadsheetApp.openById(ID_DIRECTORIO).getSheetByName(`RESTRUCTURACION`);
   const masterGastosSheet = SpreadsheetApp.openById(ID_MASTER_GASTOS).getSheetByName(`DIR-CAT-SUBCAT`);
   const directorioArreglo = directorioSheet.getRange(`B1:C`).getValues()
    .filter(fila => fila[1] != `` && fila[1] != null).map(fila => [fila[1], fila[0]]);
   const masterGArreglo = masterGastosSheet.getRange(`D1:G`).getValues()
    .filter(fila => fila[1] != `` && fila[1] != null).map(fila => [fila[0], fila[3]]);
   const areasArreglo = masterGastosSheet.getRange(`P1:Q`).getValues()
    .filter(fila => fila[1] != `` && fila[1] != null);
  const directorioObjeto = arrayToObject(directorioArreglo);
  const masterGObjeto = arrayToObject(masterGArreglo);
  const areasObjeto = arrayToObject(areasArreglo);
  var numEmpleado = cerosAntes(directorioObjeto[nombre]);
  (subcatego==`EMPLEADO DEL MES`)?subcatego=`BONO GRATIFICACION`:0;
  var numSubcatego = masterGObjeto[subcatego];
  var numArea = areasObjeto[area];

  //  Logger.log(`Nombre: ${nombre}`);

   consecutivo = consecutivo.filter(fila => typeof fila[0] === `string` && fila[0].startsWith(`${numArea}-${archivo}-${numEmpleado}-${numSubcatego}`)).flat().length+1;
    const folio = seisCerosAntes(consecutivo);

  //   Logger.log(`
  //   Folio: ${folio}
  //   Consecutivo: ${consecutivo}`);

  // Logger.log(`
  // Nombre: ${nombre}
  // Folio: ${folio}
  // `);

  // Logger.log(`
  // ${numArea}-${archivo}-${numEmpleado}-${numSubcatego}-${folio}
  // `);

  if(archivo === undefined || 
  numArea === undefined || 
  numEmpleado === undefined || 
  numSubcatego === undefined || 
  folio === undefined){
    return`Identificador Invalido`;
  }
  return`${numArea}-${archivo}-${numEmpleado}-${numSubcatego}-${folio}`;
  // } catch (err) {
      Logger.log('Error al Generar Identificador: ' + err.message);
      Logger.log(`
      Area: ${area}
      Categoria: ${categoria}
      Subcategoria: ${subcatego}
      Nombre: ${nombre}
      Consecutivo: ${consecutivo.length}
      `);
      Logger.log(`
      Numero de Area: ${numArea}
      Archivo: ${archivo}
      Numero de Empleado: ${numEmpleado}
      Subcategoria Abreviada: ${numSubcatego}
      `);
      // Folio: ${folio}
      // Logger.log(`ID: ${numArea}-${archivo}-${numEmpleado}-${numSubcatego}-${folio}`);
      // Logger.log(`ID: ${numArea}-${archivo}-${numEmpleado}-${numSubcatego}`);
      // SpreadsheetApp.getActiveSpreadsheet().toast('Error al Generar Identificador: ' + err.message)
    // }
}

//////////////////////////////

function arrayToObject(data) {
  return data.reduce((acc, row) => {
    const clave = row[0];
    const valor = row[1];
    acc[clave] = valor;
    return acc;
  }, {});
}

//////////////////////////////

function personaObject(data) {
  return data.reduce((acc, row) => {
    const nombre = row[10];
    acc[nombre] = {
      QUIEN_SOL: row[2],
      DPTO_SOL: row[3],
      USO: row[4],
      AREA_APLICA: row[6],
      CATEGORIA: row[7],
      DETALLE: row[9],
      DESTINO: row[20],
      CUENTA_CLABE: row[21],
      TITULAR: row[22]
    };
    return acc;
  }, {});
}


//////////////////////////////

function cerosAntes(numero){
  try{
    numeroStr = JSON.stringify(numero);
    switch (numeroStr.length){
    case 0: numeroStr = `0000`+numeroStr; break;
    case 1: numeroStr = `000`+numeroStr; break;
    case 2: numeroStr = `00`+numeroStr; break;
    case 3: numeroStr = `0`+numeroStr; break;
    case 4: numeroStr = numeroStr; break;
    default: numeroStr = `0000`;
    }
  }catch(err){
    return undefined;
  }
  return numeroStr;
}

//////////////////////////////

function seisCerosAntes(numero){
    numeroStr = JSON.stringify(numero);
    switch (numeroStr.length){
    case 0: numeroStr = `000000`+numeroStr; break;
    case 1: numeroStr = `00000`+numeroStr; break;
    case 2: numeroStr = `0000`+numeroStr; break;
    case 3: numeroStr = `000`+numeroStr; break;
    case 4: numeroStr = `00`+numeroStr; break;
    case 5: numeroStr = `0`+numeroStr; break;
    case 6: numeroStr = numeroStr; break;
    default: numeroStr = `000000`;
  }
  return numeroStr;
}
