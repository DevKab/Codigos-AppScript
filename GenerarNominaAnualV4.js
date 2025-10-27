function generarNominaAnual() { // Actualizado: Intervalos correctos y sin copiar formato
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = ss.getSheetByName("Formato Nomina Ejemplo");
  const datos = hojaOrigen.getRange("A1001:AC1500").getValues(); // A1000:A1500
  const datosOriginales = datos.slice(  0,  12);       // A1001:AC1100
  const datosBonoL      = datos.slice(100, 112);       // A1101:AC1200
  const datosBonoD      = datos.slice(200, 212);       // A1201:AC1300
  const datosBonoT      = datos.slice(300, 312);       // A1301:AC1400
  const datosInge       = datos.slice(400, 412);       // A1401:AC1500

  const encabezado = datosOriginales[0];
  const plantilla = datosOriginales.slice(1).filter(fila => fila.some(cell => cell !== ""));
  const plantillaBonoD = datosBonoD.slice(1).filter(fila => fila.some(cell => cell !== ""));
  const plantillaBonoL = datosBonoL.slice(1).filter(fila => fila.some(cell => cell !== ""));
  const plantillaBonoT = datosBonoT.slice(1).filter(fila => fila.some(cell => cell !== ""));
  const plantillaInge = datosInge.slice(1).filter(fila => fila.some(cell => cell !== ""));

  const nombreNuevaHoja = "Nómina 2025 v4";
  let hojaNueva = ss.getSheetByName(nombreNuevaHoja);
  if (hojaNueva) ss.deleteSheet(hojaNueva);

  hojaNueva = ss.insertSheet(nombreNuevaHoja);
  hojaNueva.appendRow(encabezado);

  // Cambio clave: Usar jueves en lugar de miércoles
  const primerJueves = getFirstThursday(2025); // 👈 Función actualizada
  let todasLasFilas = [];

  for (let i = 0; i < 52; i++) {
    const fechaSemana = new Date(primerJueves);
    fechaSemana.setDate(fechaSemana.getDate() + i * 7);

    const semana = getWeekNumber(fechaSemana);
    const mes = fechaSemana.getMonth() + 1;
    const año = fechaSemana.getFullYear();
    const fJ = new Date(getForthThursdayOfMonth(2025,mes-1));
    const tJ = new Date(getThirdThursdayOfMonth(2025,mes-1));

    plantilla.forEach(fila => {
      const nuevaFila = [...fila];
      nuevaFila[1]  = fechaSemana;  // Columna B  (fecha)
      nuevaFila[26] = semana;       // Columna AA (semana)
      nuevaFila[27] = mes;          // Columna AB (mes)
      nuevaFila[28] = año;          // Columna AC (año)
      todasLasFilas.push(nuevaFila);
    });
    if(getWeekNumber(tJ) == semana){ // Inge del Mes
      plantillaInge.forEach(filaInge => {
        const nuevaFilaInge = [...filaInge];
        nuevaFilaInge[1]  = fechaSemana;  // Columna B  (fecha)
        nuevaFilaInge[26] = semana;       // Columna AA (semana)
        nuevaFilaInge[27] = mes;          // Columna AB (mes)
        nuevaFilaInge[28] = año;          // Columna AC (año)
        todasLasFilas.push(nuevaFilaInge);
      });
    }
    if(getWeekNumber(fJ) == semana){ // Bono DESPENSA (apartir del 6to mes)
      plantillaBonoD.forEach(filaBono => {
        const nuevaFilaBono = [...filaBono];
        nuevaFilaBono[1]  = fechaSemana;  // Columna B  (fecha)
        nuevaFilaBono[26] = semana;       // Columna AA (semana)
        nuevaFilaBono[27] = mes;          // Columna AB (mes)
        nuevaFilaBono[28] = año;          // Columna AC (año)
        let diff = monthDiff(new Date(nuevaFilaBono[25]),fechaSemana)
        if(diff>=6){
          todasLasFilas.push(nuevaFilaBono);
        } else {
          console.log(`${nuevaFilaBono[10]} = ${diff}`);
          nuevaFilaBono[23] = 0;
          todasLasFilas.push(nuevaFilaBono);
        }
      });
    }
    if(getWeekNumber(fJ) == semana){ // Bono LEALTAD (apartir del 12vo mes)
      plantillaBonoL.forEach(filaBono => {
        const nuevaFilaBono = [...filaBono];
        nuevaFilaBono[1]  = fechaSemana;  // Columna B  (fecha)
        nuevaFilaBono[26] = semana;       // Columna AA (semana)
        nuevaFilaBono[27] = mes;          // Columna AB (mes)
        nuevaFilaBono[28] = año;          // Columna AC (año)
        let diff = monthDiff(new Date(nuevaFilaBono[25]),fechaSemana)
        if(diff>=12){
          todasLasFilas.push(nuevaFilaBono);
        } else {
          console.log(`${nuevaFilaBono[10]} = ${diff}`);
          nuevaFilaBono[23] = 0;
          todasLasFilas.push(nuevaFilaBono);
        }
      });
    }
    if(getWeekNumber(fJ) == semana){ // Bono TRASLADO (apartir del 12vo mes)
      plantillaBonoT.forEach(filaBono => {
        const nuevaFilaBono = [...filaBono];
        nuevaFilaBono[1]  = fechaSemana;  // Columna B  (fecha)
        nuevaFilaBono[26] = semana;       // Columna AA (semana)
        nuevaFilaBono[27] = mes;          // Columna AB (mes)
        nuevaFilaBono[28] = año;          // Columna AC (año)
        let diff = monthDiff(new Date(nuevaFilaBono[25]),fechaSemana)
        if(diff>=12){
          todasLasFilas.push(nuevaFilaBono);
        } else {
          console.log(`${nuevaFilaBono[10]} = ${diff}`);
          nuevaFilaBono[23] = 0;
          todasLasFilas.push(nuevaFilaBono);
        }
      });
    }
  }

  // Escribe todos los datos
  if (todasLasFilas.length > 0) {
    hojaNueva.getRange(2, 1, todasLasFilas.length, todasLasFilas[0].length).setValues(todasLasFilas);
  }
  SpreadsheetApp.flush(); // Refresca la hoja para asegurarse que se haya generado correctamente
  Logger.log(`¡Nómina 2025 generada con ${plantilla.length} empleados (jueves)!`); // Mensaje actualizado
}
