/**
 * Devuelve los enlaces en el interior del 
 * intervalo de celdas que se pasa como parámetro
 * @param {A1:A10} intervalo Intervalo de datos.
 * 
 * ¡Solo soporta referencias a celdas o intervalos en formato A1 (nada de composición matricial)!
 * Esta versión extrae el URL aunque solo afecte a parte del texto (solo devuelve el 1er enlace encontrado)
 * 
 * @return Intervalo de URLs
 * @customfunction
 */
function OBTENERENLACE(intervalo) {

  const hdc = SpreadsheetApp.getActiveSheet();
  const referencia = hdc.getActiveCell().getFormula().match(/\((.+)\)/)[1];

  if (referencia.includes('{')) throw('Especificación de rango no soportada.');

  if (intervalo.map) {
    // Procesar intervalo
    const rtvs = hdc.getRange(referencia).getRichTextValues(); // ningún elemento será null, aunque la celda no contenga texto
    return rtvs.map(rtvFila => rtvFila.map(rtvCelda => rtvCelda.getRuns().find(run => run.getLinkUrl())?.getLinkUrl()));

  } else {
    // Procesar celda única
    const runs = hdc.getRange(referencia).getRichTextValue(); // devuelve null si la celda no contiene texto
    return runs?.getRuns().find(run => run.getLinkUrl())?.getLinkUrl();
  }

}

/**
 * Devuelve los enlaces en el interior de
 * la celda que se pasa como parámetro
 * 
 * Esta versión solo extrae el url si el enlace afecta a todo el texto de la celda
 * 
 * @param {A1} intervalo Intervalo de datos.
 * @return Intervalo de URLs
 * 
 */
function OBTENERENLACE1(intervalo) {

  const hdc = SpreadsheetApp.getActiveSheet();
  const referencia = hdc.getActiveCell().getFormula().match(/\((.+)\)/)[1]; 
  const url = hdc.getRange(referencia).getRichTextValue().getLinkUrl();
  return url;
}

/**
 * Devuelve los enlaces en el interior de
 * la celda que se pasa como parámetro
 * @param {A1} intervalo Intervalo de datos.
 * 
 * Esta versión extrae el URL aunque solo afecte a parte del texto (solo devuelve el 1º)
 * 
 * @return Intervalo de URLs
 * 
 */
function OBTENERENLACE2(intervalo) {

  const hdc = SpreadsheetApp.getActiveSheet();
  const referencia = hdc.getActiveCell().getFormula().match(/\((.+)\)/)[1]; 
  const runs = hdc.getRange(referencia).getRichTextValue().getRuns();
  return runs.find(run => run.getLinkUrl()).getLinkUrl();

}