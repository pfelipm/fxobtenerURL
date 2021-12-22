/**
 * Devuelve los URL de los enlaces encontrados en el interior de las celdas del intervalo que se pasa como parámetro.
 * 
 * Admite:
 *  (+) Celdas en las que se ha utilizado la función HIPERENLACE().
 *  (+) Celdas en las que se han generado 1 o varios enlaces usando Insertar → Enlace o el botón 🔗.
 * 
 * Limitaciones:
 *  (-) Solo soporta referencias a celdas o intervalos en formato A1 estricto (nada de composición matricial).
 *  (-) Si el contenido de la celda es un número, su enlace no será recuperado (basta con aplicarle formato de texto para que sí lo sea).
 * 
 * @param {A1:A10}      intervalo   Intervalo de datos.
 * @param {VERDADERO}   todos       VERDADERO si se desean extraer todos los URL, FALSO si se omite.
 * @param {";"}         separador   Secuencia de caracteres a utilizar para separar los URL extraídos de una misma celda, ", " si se omite.
 *
 * @return Intervalo de URLs, como cadenas de texto.
 * 
 * @customfunction
 */
function OBTENERENLACES(intervalo, todos = false, separador = ', ') {

  const hdc = SpreadsheetApp.getActiveSheet();

  // Truco: se obtiene el primer parámetro (literal de la referencia al intervalo) parseando el valor de la celda que contiene esta fórmula
  const referencia = hdc.getActiveCell().getFormula().match(/\(([A-Za-z0-9:]+)/)?.[1];
  if (!referencia) throw 'Especificación de rango no soportada.';

  if (intervalo.map) {
    // [A] Procesar intervalo de celdas
    const rtvs = hdc.getRange(referencia).getRichTextValues();
    // getRichTextValues() siempre devuelve un [[]], pero todos sus elementos son de tipo getRichTextValue, aún cuando la celda no contiene texto
    // https://twitter.com/pfelipm/status/1459949065089789954
    if (todos) {
      return rtvs.map(rtvFila => rtvFila.map(rtvCelda => rtvCelda.getRuns().filter(run => run.getLinkUrl()).map(run => run.getLinkUrl()).join(separador)));
    } else {
      return rtvs.map(rtvFila => rtvFila.map(rtvCelda => rtvCelda.getRuns().find(run => run.getLinkUrl())?.getLinkUrl()));
    }
  } else {
    // [B] Procesar celda única
    const rtv = hdc.getRange(referencia).getRichTextValue();
    // getRichTextValue() devuelve null cuando la celda no contiene texto, por esa razón se usa rtv?.
    if (todos) {
      return rtv?.getRuns().filter(run => run.getLinkUrl()).map(run => run.getLinkUrl()).join(separador);
    } else {
      return rtv?.getRuns().find(run => run.getLinkUrl())?.getLinkUrl();
    }
  }

}