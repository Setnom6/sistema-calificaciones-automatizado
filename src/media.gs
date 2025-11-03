/***************************************************
 * media.gs
 * Calcula las medias y muestra trimestres en columnas
 * Modularizado: createMediaForAlumno crea la hoja 'media' para un alumno
 * y devuelve las referencias A1 de las medias trimestrales.
 ***************************************************/

/**
 * Crea/actualiza la hoja de media para UN alumno.
 * - ss: Spreadsheet
 * - al: { nombre, primerApellido, segundoApellido }
 * - criterios: array de criterios
 * - instrumentosPorTrim: { trim1:[], trim2:[], trim3:[] }
 *
 * Devuelve: { trim1: "B10", trim2: "E10", trim3: "H10", sheetName: "<base>_media" }
 */
function createMediaForAlumno(ss, al, criterios, instrumentosPorTrim) {
  const baseName = sanitizeSheetName(al.primerApellido + "_" + al.nombre);
  const sheetName = baseName + "_media";

  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(makeUniqueSheetName(ss, sheetName));
  } else {
    sh.clear();
  }

  // --- Comportamiento equivalente al anterior buildMedias ---
  const competencias = [...new Set(criterios.map(c => c.competencia))];
  const colorMap = {};
  competencias.forEach(comp => {
    const c = criterios.find(x => x.competencia === comp);
    colorMap[comp] = c && c.color ? c.color : "#ffffff";
  });

  // Nueva fila superior con nombre completo
  const fullName = `${al.nombre} ${al.primerApellido} ${al.segundoApellido || ""}`.trim();
  sh.getRange(1, 1).setValue(fullName)
    .setFontWeight("bold")
    .setFontSize(12);
  sh.getRange(1, 1, 1, 12).merge();
  sh.getRange(1, 1).setHorizontalAlignment("center").setBackground("#e8e8e8");

  const bloques = ["1er Trimestre", "2º Trimestre", "3er Trimestre"];
  const anchoBloque = 4;

  const mapCeldasLocal = {};

  // Necesitamos funciones auxiliares para localizar datos en desgloses (emulan al buildMedias)
  function dataStartRowForTrim(t) {
    const claves = ["trim1", "trim2", "trim3"];
    let offset = 3;
    for (let i = 0; i < t; i++) {
      const rowsPerTrim = (instrumentosPorTrim[claves[i]] || []).length;
      offset += rowsPerTrim + 2;
    }
    return offset + 1;
  }
  function dataEndRowForTrim(t) {
    const claves = ["trim1", "trim2", "trim3"];
    const start = dataStartRowForTrim(t);
    const rowsPerTrim = (instrumentosPorTrim[claves[t]] || []).length;
    return start + rowsPerTrim - 1;
  }

  bloques.forEach((bloque, idxTrim) => {
    const colOffset = 1 + idxTrim * anchoBloque;
    sh.getRange(2, colOffset).setValue(bloque).setFontWeight("bold").setFontSize(12);

    competencias.forEach((comp, iComp) => {
      const criteriosComp = criterios.filter(c => c.competencia === comp);

      const validParts = criteriosComp.map(c => {
        const idxGlobal = criterios.indexOf(c);
        const colIdxDesglose = idxGlobal + 2;
        const startRow = dataStartRowForTrim(idxTrim);
        const endRow = dataEndRowForTrim(idxTrim);
        const colLetter = columnLetter(colIdxDesglose);
        const rangeA1 = `'${baseName}_desglose'!${colLetter}${startRow}:${colLetter}${endRow}`;
        return `FILTER(${rangeA1},(${rangeA1}>=0)*(${rangeA1}<=10))`;
      });

      let formulaComp = `""`;
      if (validParts.length > 0) {
        formulaComp = `=IFERROR(AVERAGE({${validParts.join(",")}}),"")`;
      }

      const nameCell = sh.getRange(iComp + 3, colOffset);
      const notaCell = sh.getRange(iComp + 3, colOffset + 1);
      nameCell.setValue(comp);
      notaCell.setFormula(formulaComp).setNumberFormat("0.00").setBackground(colorMap[comp]);

      const helperCell = sh.getRange(iComp + 3, colOffset + 2);
      const countParts = criteriosComp.map(c => {
        const idxGlobal = criterios.indexOf(c);
        const colIdxDesglose = idxGlobal + 2;
        const startRow = dataStartRowForTrim(idxTrim);
        const endRow = dataEndRowForTrim(idxTrim);
        const colLetter = columnLetter(colIdxDesglose);
        const rangeA1 = `'${baseName}_desglose'!${colLetter}${startRow}:${colLetter}${endRow}`;
        return `IF(COUNTIFS(${rangeA1},">=0",${rangeA1},"<=10")>0,1,0)`;
      });
      helperCell.setFormula("=" + countParts.join("+")).setNumberFormat("0");

      const helperA1 = helperCell.getA1Notation();
      const totalCount = criteriosComp.length;
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${helperA1}<${totalCount}`)
        .setFontColor("#FF0000")
        .setRanges([notaCell])
        .build();
      const rules = sh.getConditionalFormatRules();
      rules.push(rule);
      sh.setConditionalFormatRules(rules);
    });

    const firstCompRow = 3;
    const lastCompRow = competencias.length + 2;
    const promedioRow = competencias.length + 4;
    const promedioCell = sh.getRange(promedioRow, colOffset + 1);
    const colLetterMed = columnLetter(colOffset + 1);
    promedioCell.setFormula(`=IFERROR(AVERAGE(${colLetterMed}${firstCompRow}:${colLetterMed}${lastCompRow}),"")`)
      .setFontWeight("bold").setFontSize(12).setNumberFormat("0.00");

    mapCeldasLocal["trim" + (idxTrim + 1)] = promedioCell.getA1Notation();
  });

  // Ajuste visual
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(5, 180);
  sh.setColumnWidth(6, 90);
  sh.setColumnWidth(9, 180);
  sh.setColumnWidth(10, 90);

  // Devolvemos la referencia a las celdas de medias (A1) para su uso externo
  return { sheetName: sheetName, trim1: mapCeldasLocal.trim1, trim2: mapCeldasLocal.trim2, trim3: mapCeldasLocal.trim3 };
}

/**
 * buildMedias: mantiene la misma firma que antes (ss, alumnos, criterios, instrumentosPorTrim)
 * pero delega en createMediaForAlumno para cada alumno. Devuelve el mapCeldas al final
 * (para que populateGeneral siga funcionando).
 */
function buildMedias(ss, alumnos, criterios, instrumentosPorTrim) {
  const mapCeldas = {};

  alumnos.forEach(al => {
    try {
      const res = createMediaForAlumno(ss, al, criterios, instrumentosPorTrim);
      const baseName = sanitizeSheetName(al.primerApellido + "_" + al.nombre);
      mapCeldas[baseName] = {
        trim1: res.trim1,
        trim2: res.trim2,
        trim3: res.trim3
      };
    } catch (e) {
      Logger.log("Error creando media para " + al.nombre + " " + al.primerApellido + ": " + e);
    }
  });

  return mapCeldas;
}

/**
 * Función auxiliar pública para obtener las referencias A1 de las medias
 * de un alumno que ya tiene hoja *_media creada.
 * - ss: Spreadsheet
 * - al: { nombre, primerApellido }
 * Devuelve { trim1, trim2, trim3 } con A1 o nulls si no encuentra.
 */
function getMediaRefsForAlumno(ss, al) {
  const baseName = sanitizeSheetName(al.primerApellido + "_" + al.nombre);
  const sh = ss.getSheetByName(baseName + "_media");
  if (!sh) return null;

  // Intentamos localizar las celdas tal y como createMediaForAlumno las deja:
  // - las medias trimestrales están en la columna (colOffset + 1) en la fila promedioRow (competencias.length + 4)
  // Ya que no conocemos 'criterios' aquí, la forma simple es buscar la primera celda con formato numérico y no vacía en
  // cada bloque. Pero para mantener robustez, vamos a replicar el cálculo de offsets leyendo la hoja:
  // asumimos que las fórmulas de promedio están en filas con valor numérico (formato 0.00). Buscamos 3 promedios en la hoja.
  const vals = sh.getDataRange().getValues();
  const formulas = sh.getDataRange().getFormulas();
  const found = [];

  for (let r = 0; r < vals.length && found.length < 3; r++) {
    for (let c = 0; c < vals[0].length && found.length < 3; c++) {
      // consideramos celda con fórmula que contiene AVERAGE(...) como candidata
      const f = formulas[r][c];
      if (f && f.toString().toUpperCase().indexOf("AVERAGE(") !== -1) {
        const a1 = sh.getRange(r + 1, c + 1).getA1Notation();
        // evitamos duplicados
        if (!found.includes(a1)) found.push(a1);
      }
    }
  }

  // Si no conseguimos por fórmulas, devolvemos null
  if (found.length < 3) {
    return null;
  }

  return { trim1: found[0], trim2: found[1], trim3: found[2] };
}
