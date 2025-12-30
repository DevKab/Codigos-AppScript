function validarDatos(archivo, numArea, numEmpleado, numSubcatego, folio) {
  const variables = {
    archivo,
    numArea,
    numEmpleado,
    numSubcatego,
    folio
  };
  const faltantes = Object.entries(variables)
    .filter(([_, valor]) => valor === undefined)
    .map(([nombre]) => nombre);
  if (faltantes.length > 0) {
    return `Identificador inválido. Falta: ${faltantes.join(', ')}`;
  }
  return`${numArea}-${archivo}-${numEmpleado}-${numSubcatego}-${folio}`;
}
