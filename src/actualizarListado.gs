/***************************************************
 * actualizarListado.gs (modularizado y agnóstico)
 * Añade nuevos alumnos detectados en el sheet 'listado'
 * al spreadsheet de calificaciones existente, creando
 * su hoja de desglose y su hoja de media, e insertando
 * la fila en 'general' en orden alfabético por primer apellido.
 ***************************************************/
function actualizarListado() {
  const CLASS_FOLDER_NAME_LISTADO = "claseEjemplo"; // ← editar si hace falta

  try {
    // === 1) Abrir archivo calificaciones existente ===
    const parentFolder = getProjectFolder();
    const files = parentFolder.getFilesByName("calificaciones_" + CLASS_FOLDER_NAME_LISTADO);
    if (!files.hasNext()) throw new Error("No se encontró el archivo de calificaciones. Ejecuta buildProjectClase primero.");

    const ss = SpreadsheetApp.open(files.next());
    const sheetGeneral = ss.getSheetByName("general");
    if (!sheetGeneral) throw new Error("Falta la hoja 'general'. El proyecto debe construirse antes.");

    // === 2) Leer listado actual desde carpeta de clase (usando parseListadoRows) ===
    const classFolder = findClassFolderByName(CLASS_FOLDER_NAME_LISTADO);
    const fileListado = findFileInFolder(classFolder, "listado");
    const ssListado = SpreadsheetApp.openById(fileListado.getId());
    const alumnosListado = parseListadoRows(ssListado.getSheets()[0].getDataRange().getValues())
      .map(a => ({
        nombre: (a.nombre || "").trim(),
        primerApellido: (a.primerApellido || "").trim(),
        segundoApellido: (a.segundoApellido || "").trim()
      }))
      .sort((a, b) => a.primerApellido.localeCompare(b.primerApellido));

    // === 3) Leer criterios e instrumentos (usando funciones auxiliares) ===
    const fileCriterios = findFileInFolder(classFolder, "criteriosDeEvaluacion");
    const fileInstrumentos = findFileInFolder(classFolder, "instrumentos");
    const ssCriterios = SpreadsheetApp.openById(fileCriterios.getId());
    const criterios = parseCriteriosRows(ssCriterios.getSheets()[0].getDataRange().getValues(), ssCriterios.getSheets()[0]);
    const instrumentos = parseInstrumentosFromSheet(fileInstrumentos);

    // === 4) Leer alumnos ya existentes en 'general' ===
    const dataGeneral = sheetGeneral.getDataRange().getValues();
    const existentes = dataGeneral.slice(1).map(r => ({
      nombre: (r[0] || "").toString().trim(),
      primerApellido: (r[1] || "").toString().trim(),
      segundoApellido: (r[2] || "").toString().trim()
    })).filter(a => a.primerApellido && a.nombre);

    // === 5) Determinar alumnos nuevos ===
    const key = a => (a.primerApellido + "||" + a.nombre).toLowerCase();
    const existentesSet = new Set(existentes.map(key));
    const nuevos = alumnosListado.filter(a => !existentesSet.has(key(a)));

    if (nuevos.length === 0) {
      Logger.log("No se han detectado alumnos nuevos.");
      return;
    }

    // === 6) Crear desglose y media para cada nuevo alumno ===
    nuevos.forEach(al => {
      const baseName = sanitizeSheetName(al.primerApellido + "_" + al.nombre);
      Logger.log("Añadiendo nuevo alumno: " + baseName);

      // a) Crear desglose (usa createDesgloseForAlumno)
      let desgloseInfo = null;
      try {
        desgloseInfo = createDesgloseForAlumno(ss, al, criterios, instrumentos);
      } catch (e) {
        Logger.log("Error al crear desglose para " + baseName + ": " + e);
      }

      // b) Crear media (usa createMediaForAlumno)
      let mediaInfo = null;
      try {
        mediaInfo = createMediaForAlumno(ss, al, criterios, instrumentos);
      } catch (e) {
        Logger.log("Error al crear media para " + baseName + ": " + e);
      }

      // c) Si el desglose existe, asegurar que enlace a la media
      if (desgloseInfo && mediaInfo) {
        const shDesglose = ss.getSheetByName(desgloseInfo.sheetName);
        if (shDesglose) {
          const range = shDesglose.getDataRange();
          const values = range.getValues();
          for (let r = 0; r < values.length; r++) {
            for (let c = 0; c < values[0].length; c++) {
              if (values[r][c] === "Media no encontrada") {
                shDesglose.getRange(r + 1, c + 1)
                  .setFormula(`=HYPERLINK("#gid=${ss.getSheetByName(mediaInfo.sheetName).getSheetId()}","Ver media")`);
              }
            }
          }
        }
      }

      // d) Obtener referencias de las medias (usa getMediaRefsForAlumno)
      // d) Usar directamente las referencias devueltas por createMediaForAlumno
      const mediaRefs = mediaInfo ? {
        trim1: mediaInfo.trim1,
        trim2: mediaInfo.trim2,
        trim3: mediaInfo.trim3
      } : getMediaRefsForAlumno(ss, al); // fallback por si acaso


      // e) Crear fila en 'general'
      const url = ss.getUrl();
      const shDesglose = ss.getSheetByName(baseName + "_desglose");
      const linkFormula = shDesglose ? `=HYPERLINK("${url}#gid=${shDesglose.getSheetId()}","Abrir desglose")` : "";

      const formulaTrim1 = mediaRefs ? `=IFERROR('${baseName}_media'!${mediaRefs.trim1},"")` : "";
      const formulaTrim2 = mediaRefs ? `=IFERROR('${baseName}_media'!${mediaRefs.trim2},"")` : "";
      const formulaTrim3 = mediaRefs ? `=IFERROR('${baseName}_media'!${mediaRefs.trim3},"")` : "";

      const rowValues = [
        al.nombre,
        al.primerApellido,
        al.segundoApellido,
        formulaTrim1,
        formulaTrim2,
        formulaTrim3,
        linkFormula
      ];

      // f) Insertar la fila en orden alfabético
      const apellidosExistentes = sheetGeneral.getRange(2, 2, sheetGeneral.getLastRow() - 1, 1).getValues().map(r => (r[0] || "").toString().trim());
      let insertRow = sheetGeneral.getLastRow() + 1;
      for (let i = 0; i < apellidosExistentes.length; i++) {
        if (apellidosExistentes[i].localeCompare(al.primerApellido) > 0) {
          insertRow = i + 2;
          break;
        }
      }

      if (insertRow <= sheetGeneral.getLastRow()) {
        sheetGeneral.insertRowBefore(insertRow);
        sheetGeneral.getRange(insertRow, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sheetGeneral.appendRow(rowValues);
      }

      const targetRow = (insertRow <= sheetGeneral.getLastRow()) ? insertRow : sheetGeneral.getLastRow();
      sheetGeneral.getRange(targetRow, 4, 1, 3).setNumberFormat("0.00");
    });

    // === 7) Ajustes visuales básicos ===
    const headers = ["Nombre", "Primer Apellido", "Segundo Apellido",
      "Calificación 1er Trim", "Calificación 2º Trim", "Calificación 3er Trim", "Hoja desglose"];
    for (let i = 1; i <= headers.length; i++) sheetGeneral.setColumnWidth(i, headers[i - 1].length * 10);
    sheetGeneral.getRange(1, 1, 1, headers.length).setFontWeight("bold").setFontSize(11);

    SpreadsheetApp.flush();
    Logger.log(`Añadidos ${nuevos.length} alumno(s) nuevos.`);
    Logger.log("Nuevos: " + nuevos.map(a => a.primerApellido + ", " + a.nombre).join("; "));

  } catch (e) {
    Logger.log("Error en actualizarListado: " + e);
    Logger.log(e.stack || e);
  }
}
