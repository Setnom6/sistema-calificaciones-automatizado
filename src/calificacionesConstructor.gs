/**
 * calificacionesConstructor.gs
 * Construye/actualiza la hoja calificacionesN.
 * Para el formato llama a funciones de formatter.gs
 *
 * Exporta: buildCalificaciones(n, alumnos, instrumentos, claveToColor)
 */

function buildCalificaciones(n, alumnos, instrumentos, claveToColor) {
  const ss = SpreadsheetApp.getActive();
  const hojaCalifName = "calificaciones" + n;

  // ---------- PREPARAR SHEET CALIFICACIONES ----------
  let sheetCalif = ss.getSheetByName(hojaCalifName);
  if (sheetCalif) {
    // romper merges previos
    sheetCalif.getRange(1, 1, sheetCalif.getMaxRows(), sheetCalif.getMaxColumns()).breakApart();
  } else {
    sheetCalif = ss.insertSheet(hojaCalifName);
  }

  // ---------- FILA DE HEADERS ----------
  const califHeadersRow1 = ["Alumno"];
  const califHeadersRow2 = ["Alumno"];
  const headerGroupSizes = [];

  instrumentos.forEach(inst => {
    const k = inst.criterios.length;
    if (k === 0) return; // ignorar instrumentos sin criterios

    for (let j = 0; j < k; j++) {
      if (j === 0) califHeadersRow1.push(inst.nombre);
      else califHeadersRow1.push("");
      califHeadersRow2.push(inst.criterios[j]);
    }
    califHeadersRow1.push(inst.nombre);
    califHeadersRow2.push("Media");
    headerGroupSizes.push(k + 1);
  });

  const totalCols = califHeadersRow2.length;
  const totalRows = 2 + alumnos.length;

  // ---------- CREAR TEMPORAL ----------
  let temp = ss.getSheetByName("TEMP_CALIF");
  if (temp) ss.deleteSheet(temp);
  temp = ss.insertSheet("TEMP_CALIF");
  temp.clear();

  // headers y alumnos
  temp.getRange(1, 1, 1, totalCols).setValues([califHeadersRow1]);
  temp.getRange(2, 1, 1, totalCols).setValues([califHeadersRow2]);
  if (alumnos.length > 0) temp.getRange(3, 1, alumnos.length, 1).setValues(alumnos);
  if (alumnos.length > 0 && totalCols > 1) {
    temp.getRange(3, 2, alumnos.length, totalCols - 1)
        .setValues(Array.from({ length: alumnos.length }, () => Array(totalCols - 1).fill("")));
  }

  // ---------- MAPEO DATOS ANTIGUOS ----------
  let oldData = null;
  let oldHeadersRow1 = null;
  let oldHeadersRow2 = null;
  let oldColumnsList = [];
  let oldRowByAlumno = {};

  if (sheetCalif) {
    const lastRow = sheetCalif.getLastRow();
    const lastCol = sheetCalif.getLastColumn();
    if (lastRow >= 2 && lastCol >= 1) {
      oldData = sheetCalif.getRange(1, 1, lastRow, lastCol).getValues();
      oldHeadersRow1 = oldData[0].map(x => x ? x.toString().trim() : "");
      oldHeadersRow2 = oldData[1].map(x => x ? x.toString().trim() : "");
      for (let c = 0; c < oldHeadersRow2.length; c++) {
        oldColumnsList.push({ colIndex: c, h1: oldHeadersRow1[c] || "", h2: oldHeadersRow2[c] || "" });
      }
      for (let r = 2; r < oldData.length; r++) {
        const val = (oldData[r][0] || "").toString().trim();
        if (val !== "") oldRowByAlumno[val] = r;
      }
    }
  }

  // ---------- CONSTRUIR BLOQUES ANTIGUOS ----------
  const oldBlocks = {}; // instrumento -> {startCol, endCol, columns, data}
  if (oldColumnsList.length > 0) {
    let ptr = 1;
    while (ptr < oldColumnsList.length) {
      const instName = oldHeadersRow1[ptr];
      if (!instName) { ptr++; continue; }

      let start = ptr;
      let end = ptr;
      while (end + 1 < oldColumnsList.length && (oldHeadersRow1[end + 1] === "" || oldHeadersRow1[end + 1] === instName)) {
        end++;
      }

      const cols = [];
      for (let c = start; c <= end; c++) {
        const clave = oldHeadersRow2[c];
        if (clave && clave !== "Media") cols.push({ clave, colIndex: c });
      }

      const blockData = [];
      for (let r = 2; r < oldData.length; r++) {
        blockData.push(oldData[r].slice(start, end + 1));
      }

      oldBlocks[instName] = { startCol: start, endCol: end, columns: cols, data: blockData };
      ptr = end + 1;
    }
  }

  // ---------- COPIAR DATOS ANTIGUOS ----------
  let colPtrNew = 2;
  instrumentos.forEach(inst => {
    const blockSize = inst.criterios.length + 1;
    const oldBlock = oldBlocks[inst.nombre];

    if (oldBlock) {
      inst.criterios.forEach((clave, idx) => {
        const colNew = colPtrNew + idx;
        const oldColObj = oldBlock.columns.find(c => c.clave === clave);
        if (oldColObj) {
          const oldCol = oldColObj.colIndex;
          for (let i = 0; i < alumnos.length; i++) {
            const alumnoNombre = alumnos[i][0];
            const oldRowIndex = oldRowByAlumno[alumnoNombre];
            if (oldRowIndex !== undefined) {
              const val = oldData[oldRowIndex][oldCol];
              temp.getRange(3 + i, colNew).setValue(val);
            }
          }
        }
      });
    }
    colPtrNew += blockSize;
  });

  // ---------- REEMPLAZAR SHEET CALIFICACIONES ----------
  const tempValues = temp.getDataRange().getValues();
  const tempNumRows = temp.getLastRow();
  const tempNumCols = temp.getLastColumn();
  // ---------- DEDUPLICAR FILAS DE ALUMNOS EN TEMPORAL ----------
  // Comparamos toda la fila (todas las columnas) para eliminar duplicados exactos
  if (tempNumRows > 2) {
    const seenRowSigs = new Set();
    const uniqueDataRows = [];
    let duplicatesCount = 0;

    // tempValues[0] y tempValues[1] son las dos filas de cabecera; copiarlas
    const newTempValues = [ tempValues[0], tempValues[1] ];

    for (let r = 2; r < tempValues.length; r++) {
      const row = tempValues[r];
      // crear firma de fila con trim para evitar diferencias por espacios
      const sig = row.map(c => (c === null || c === undefined) ? "" : c.toString().trim()).join("|");
      if (!seenRowSigs.has(sig)) {
        seenRowSigs.add(sig);
        uniqueDataRows.push(row);
        newTempValues.push(row);
      } else {
        duplicatesCount++;
      }
    }

    if (duplicatesCount > 0) {
      Logger.log(`buildCalificaciones: eliminados ${duplicatesCount} duplicados exactos en TEMP_CALIF`);
    }

    // Reemplazar tempValues y actualizar conteos
    // Nota: no modificamos la hoja TEMP_CALIF; trabajamos con newTempValues y actualizamos variables
    tempValues.length = 0;
    Array.prototype.push.apply(tempValues, newTempValues);
    // actualizar contadores locales
    // tempNumRows y tempNumCols se obtendrán más abajo desde temp o desde tempValues
  }

  sheetCalif.clear({ contentsOnly: false, formatOnly: false });
  if (sheetCalif.getMaxColumns() < tempNumCols)
    sheetCalif.insertColumnsAfter(sheetCalif.getMaxColumns(), tempNumCols - sheetCalif.getMaxColumns());
  if (sheetCalif.getMaxRows() < tempNumRows)
    sheetCalif.insertRowsAfter(sheetCalif.getMaxRows(), tempNumRows - sheetCalif.getMaxRows());

  sheetCalif.getRange(1, 1, tempNumRows, tempNumCols).setValues(tempValues);

  // ---------- LIMPIAR COLUMNAS SOBRANTES ----------
  if (sheetCalif.getMaxColumns() > totalCols) {
    sheetCalif.getRange(1, totalCols + 1, sheetCalif.getMaxRows(), sheetCalif.getMaxColumns() - totalCols).clearContent().clearFormat();
    sheetCalif.deleteColumns(totalCols + 1, sheetCalif.getMaxColumns() - totalCols);
  }


  // ---------- MERGES ----------
  sheetCalif.getRange(1, 1, 2, 1).merge(); // Alumno
  let ptr = 2;
  instrumentos.forEach(inst => {
    const size = inst.criterios.length + 1;
    if (size > 1) sheetCalif.getRange(1, ptr, 1, size).merge();
    ptr += size;
  });

  // ---------- FORMATO ----------
  sheetCalif.getRange(1, 1, 2, tempNumCols).setHorizontalAlignment("center")
             .setVerticalAlignment("middle").setFontWeight("bold");
  sheetCalif.getRange(2, 1, 1, tempNumCols).setHorizontalAlignment("left");

  // ---------- COLORES ----------
  for (let c = 1; c <= tempNumCols; c++) {
    const clave = sheetCalif.getRange(2, c).getValue()?.toString().trim() || "";
    if (clave && claveToColor[clave]) {
      const color = claveToColor[clave];
      sheetCalif.getRange(2, c).setBackground(color);
      if (alumnos.length > 0) sheetCalif.getRange(3, c, alumnos.length, 1).setBackground(color);
    } else {
      sheetCalif.getRange(1, c, sheetCalif.getLastRow(), 1).setBackground(null);
    }
  }

  // ---------- FORMULAS MEDIA ----------
  colPtrNew = 2;
  instrumentos.forEach(inst => {
    const criterios = inst.criterios;
    const blockSize = criterios.length + 1;
    const colMedia = colPtrNew + criterios.length;
    for (let i = 0; i < alumnos.length; i++) {
      const row = 3 + i;
      const refs = criterios.map((_, idx) => columnToLetter(colPtrNew + idx) + row);
      const formula = `=IFERROR(AVERAGE(${refs.join(";")}); "")`;
      sheetCalif.getRange(row, colMedia).setFormula(formula);
    }
    colPtrNew += blockSize;
  });

  // ---------- VALIDACION Y FORMATO CONDICIONAL PARA NOTAS ----------
  if (alumnos.length > 0) {
    // rango de notas: filas 3 a 3+alumnos.length-1, columnas 2 hasta penúltima (sin Media)
    colPtrNew = 2;
    instrumentos.forEach(inst => {
      const criterios = inst.criterios;
      const blockSize = criterios.length + 1; // +1 para Media
      const notasRange = sheetCalif.getRange(3, colPtrNew, alumnos.length, criterios.length);

      // Validacion: solo numeros entre 0 y 10
      const rule = SpreadsheetApp.newDataValidation()
                    .requireNumberBetween(0, 10)
                    .setAllowInvalid(true) // deja ingresar valores fuera, los marcamos en rojo
                    .build();
      notasRange.setDataValidation(rule);

      // Formato condicional: color rojo si valor <0 o >10
      const redRule = SpreadsheetApp.newConditionalFormatRule()
                        .whenNumberLessThan(0)
                        .setBackground("red")
                        .setRanges([notasRange])
                        .build();
      const redRule2 = SpreadsheetApp.newConditionalFormatRule()
                        .whenNumberGreaterThan(10)
                        .setBackground("red")
                        .setRanges([notasRange])
                        .build();
      const rules = sheetCalif.getConditionalFormatRules();
      sheetCalif.setConditionalFormatRules(rules.concat([redRule, redRule2]));

      colPtrNew += blockSize;
    });
  }


  // ---------- ANCHOS ----------
  let maxNameLen = 10;
  alumnos.forEach(a => { if (a[0] && a[0].length > maxNameLen) maxNameLen = a[0].length; });
  sheetCalif.setColumnWidth(1, Math.max(200, Math.min(800, Math.round(maxNameLen * 7))));
  colPtrNew = 2;
  instrumentos.forEach(inst => {
    const blockSize = inst.criterios.length + 1;
    const perCol = Math.max(90, Math.min(420, Math.ceil(Math.max(inst.nombre.length * 10, 80 * blockSize) / blockSize)));
    for (let k = 0; k < blockSize; k++) sheetCalif.setColumnWidth(colPtrNew + k, perCol);
    colPtrNew += blockSize;
  });

  // ---------- BORDES ----------
  const rowsToBorder = 2 + alumnos.length;
  formatter_applyVerticalInstrumentBorders(sheetCalif, instrumentos, rowsToBorder);

  // ---------- FORMATO DECIMAL ----------
  if (alumnos.length > 0) sheetCalif.getRange(3, 2, alumnos.length, tempNumCols - 1).setNumberFormat("0.00");

  // ---------- CONGELAR ALUMNOS ----------
  sheetCalif.setFrozenColumns(1);

  // ---------- ELIMINAR TEMPORAL ----------
  ss.deleteSheet(temp);

  return { sheetCalif, alumnos };
}