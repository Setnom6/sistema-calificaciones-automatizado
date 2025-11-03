/***************************************************
 * actualizarInstrumentos.gs
 * Añade nuevos instrumentos en los desgloses si aparecen
 * en el sheet "instrumentos" de la carpeta de clase.
 ***************************************************/

function actualizarInstrumentos() {
  const CLASS_FOLDER_NAME_INSTRUMENTOS = "claseEjemplo"; // ← editar aquí según la clase
  
  try {
    // Buscar spreadsheet de calificaciones existente
    const parentFolder = getProjectFolder();
    const files = parentFolder.getFilesByName("calificaciones_" + CLASS_FOLDER_NAME_INSTRUMENTOS);
    if (!files.hasNext()) {
      throw new Error("No se encontró el archivo de calificaciones. Ejecuta buildProjectClase primero.");
    }
    const ss = SpreadsheetApp.open(files.next());
    const sheetGeneral = ss.getSheetByName("general");
    if (!sheetGeneral) {
      throw new Error("Falta la hoja 'general'. El proyecto debe construirse antes.");
    }

    // Leer instrumentos actualizados
    const classFolder = findClassFolderByName(CLASS_FOLDER_NAME_INSTRUMENTOS);
    const fileInstrumentos = findFileInFolder(classFolder, "instrumentos");
    const nuevosInstrumentos = parseInstrumentosFromSheet(fileInstrumentos);

    // Detectar diferencias e insertar columnas nuevas en desgloses
    const hojas = ss.getSheets().filter(sh => sh.getName().endsWith("_desglose"));
    hojas.forEach(sh => {
      const data = sh.getDataRange().getValues();

      // Buscar dónde empieza cada bloque de trimestre
      const bloques = ["1er Trimestre", "2º Trimestre", "3er Trimestre"];
      bloques.forEach((bloque, idxTrim) => {
        const claveTrim = idxTrim === 0 ? "trim1" : idxTrim === 1 ? "trim2" : "trim3";
        const filaInicio = data.findIndex(r => r[0] === bloque);
        if (filaInicio < 0) return;

        // Determinar filas con instrumentos (hasta una fila vacía o siguiente bloque)
        let fila = filaInicio + 1;
        const instrumentosActuales = [];
        while (fila < data.length && data[fila][0] && !bloques.includes(data[fila][0])) {
          instrumentosActuales.push(data[fila][0]);
          fila++;
        }

        // Calcular qué instrumentos faltan
        const nuevos = nuevosInstrumentos[claveTrim].filter(
          ins => ins && !instrumentosActuales.includes(ins)
        );
        if (nuevos.length === 0) return;

        // Insertar filas nuevas antes de la fila vacía final del bloque
        const filaInsercion = fila; // justo antes de la vacía o del siguiente bloque
        nuevos.forEach(nuevoInstrumento => {
          sh.insertRowBefore(filaInsercion);
          sh.getRange(filaInsercion, 1).setValue(nuevoInstrumento);
          const ultimaCol = sh.getLastColumn();
          const criteriosCount = ultimaCol - 3;
          const primeraCol = 2;
          const ultimaColCriterios = 1 + criteriosCount;
          const rangeCriterios = sh.getRange(filaInsercion, primeraCol, 1, criteriosCount).getA1Notation();
          const formulaMedia = `=IF(COUNTA(${rangeCriterios})>0,AVERAGEIF(${rangeCriterios},"<>"),"")`;
          sh.getRange(filaInsercion, ultimaCol - 1).setFormula(formulaMedia).setNumberFormat("0.00");
        });
      });
    });

    SpreadsheetApp.flush();
    Logger.log("Instrumentos actualizados correctamente.");
  } catch (e) {
    Logger.log("Error en actualizarInstrumentos: " + e);
    SpreadsheetApp.getUi().alert(e.message || e);
  }
}
