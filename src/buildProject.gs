/***************************************************
 * buildProject.gs
 * Crea el archivo de calificaciones completo para una clase
 ***************************************************/

// Config: cambia al nombre de la carpeta de clase correspondiente
const CLASS_FOLDER_NAME = "claseEjemplo";

function buildProjectClase() {
  try {
    const classFolder = findClassFolderByName(CLASS_FOLDER_NAME);

    const fileListado = findFileInFolder(classFolder, "listado");
    const fileCriterios = findFileInFolder(classFolder, "criteriosDeEvaluacion");
    const fileInstrumentos = findFileInFolder(classFolder, "instrumentos");
    const instrumentos = parseInstrumentosFromSheet(fileInstrumentos);


    const parentFolder = getProjectFolder();
    const spreadsheetName = "calificaciones_" + CLASS_FOLDER_NAME;

    // Crear el nuevo spreadsheet
    const newSs = SpreadsheetApp.create(spreadsheetName);
    const file = DriveApp.getFileById(newSs.getId());
    parentFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    // Preparar hoja general
    const generalSheet = newSs.getSheets()[0];
    generalSheet.setName("general");
    for (let i = 1; i < newSs.getSheets().length; i++) {
      newSs.deleteSheet(newSs.getSheets()[i]);
    }

    // Leer listado y criterios
    const ssListado = SpreadsheetApp.openById(fileListado.getId());
    const alumnos = parseListadoRows(ssListado.getSheets()[0].getDataRange().getValues())
      .sort((a, b) => a.primerApellido.localeCompare(b.primerApellido));

    const ssCriterios = SpreadsheetApp.openById(fileCriterios.getId());
    const sheetCriterios = ssCriterios.getSheets()[0];
    const criterios = parseCriteriosRows(sheetCriterios.getDataRange().getValues(), sheetCriterios);


    // Crear hojas individuales
    alumnos.forEach(al => {
      const baseName = sanitizeSheetName(al.primerApellido + "_" + al.nombre);
      newSs.insertSheet(makeUniqueSheetName(newSs, baseName + "_desglose"));
      newSs.insertSheet(makeUniqueSheetName(newSs, baseName + "_media"));
    });

    // Construcción de hojas
    const mapCeldas = buildMedias(newSs, alumnos, criterios, instrumentos);
    buildDesgloses(newSs, alumnos, criterios, instrumentos);

    populateGeneral(newSs, alumnos, mapCeldas);

    SpreadsheetApp.flush();
    Logger.log("Proyecto creado: " + newSs.getUrl());
    return newSs.getUrl();
  } catch (e) {
    Logger.log("Error en buildProjectClase: " + e);
    throw e;
  }
}
