/***************************************************
 * utils.gs
 * Funciones utilitarias globales
 ***************************************************/

function getProjectFolder() {
  try {
    const scriptFile = DriveApp.getFileById(ScriptApp.getScriptId());
    const parents = scriptFile.getParents();
    if (parents.hasNext()) return parents.next();
  } catch (e) {
    try {
      const file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
      const parents = file.getParents();
      if (parents.hasNext()) return parents.next();
    } catch (err) {}
  }
  return DriveApp.getRootFolder();
}

/** Busca la carpeta de clase por nombre dentro de la carpeta del proyecto */
function findClassFolderByName(classFolderName) {
  const baseFolder = getProjectFolder();
  const folders = baseFolder.getFoldersByName(classFolderName);
  if (!folders.hasNext()) throw new Error("No se encontró la carpeta de clase: " + classFolderName);
  return folders.next();
}

/** Busca un archivo dentro de una carpeta, cuyo nombre contenga un prefijo */
function findFileInFolder(folder, prefix) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().toLowerCase().indexOf(prefix.toLowerCase()) !== -1) {
      return f;
    }
  }
  throw new Error("No se encontró archivo que contenga: " + prefix + " en carpeta " + folder.getName());
}

/** Parsea las filas del sheet 'listado' */
function parseListadoRows(rows) {
  const alumnos = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    alumnos.push({
      nombre: r[0],
      primerApellido: r[1],
      segundoApellido: r[2] || ""
    });
  }
  return alumnos;
}

/**
 * Parsea las filas del sheet 'criteriosDeEvaluacion'
 * Lee el texto tal como se muestra en la celda (getDisplayValue)
 * y asigna colores por competencia.
 */
function parseCriteriosRows(rows, sheet) {
  const criterios = [];
  const colorMap = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    const competencia = r[1];

    // Leemos el valor tal y como se muestra en la hoja (texto exacto)
    let indexText = "";
    try {
      if (sheet) {
        indexText = (sheet.getRange(i + 1, 1).getDisplayValue() || "").toString().trim();
      } else {
        indexText = ("" + r[0]).trim();
      }
    } catch (e) {
      indexText = ("" + r[0]).trim();
    }

    // color de la celda 'Index' (columna 1) — si falta o es blanco, asignamos uno distinto
    let color = "#ffffff";
    try {
      if (sheet) color = sheet.getRange(i + 1, 1).getBackground();
    } catch (e) { color = "#ffffff"; }

    if (!colorMap[competencia]) {
      if (!color || color === "#ffffff" || color === "#000000") {
        color = getUnusedColor(Object.values(colorMap));
      }
      colorMap[competencia] = color;
    }

    criterios.push({
      index: indexText,
      competencia: competencia,
      criterio: r[2],
      color: colorMap[competencia]
    });
  }

  return criterios;
}

/** Devuelve un color no usado aún */
function getUnusedColor(usedColors) {
  const palette = ["#fce5cd", "#d9ead3", "#c9daf8", "#f4cccc", "#fff2cc", "#d0e0e3", "#ead1dc"];
  for (let c of palette) {
    if (!usedColors.includes(c)) return c;
  }
  return "#e2efda";
}

/** Sanitiza nombres de hoja para evitar caracteres inválidos */
function sanitizeSheetName(name) {
  return name.replace(/[\\\/\?\*\[\]]/g, "_").substring(0, 90);
}

/** Asegura nombres únicos de hoja */
function makeUniqueSheetName(ss, baseName) {
  let name = baseName;
  let counter = 1;
  while (ss.getSheetByName(name)) {
    name = baseName + "_" + counter++;
  }
  return name;
}

/** Convierte índice de columna a letra A1 */
function columnLetter(col) {
  let temp, letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

/**
 * Lee los instrumentos de evaluación desde el sheet 'instrumentos' del folder de clase.
 * Devuelve un objeto con tres arrays: { trim1: [...], trim2: [...], trim3: [...] }.
 */
function parseInstrumentosFromSheet(fileInstrumentos) {
  const ss = SpreadsheetApp.openById(fileInstrumentos.getId());
  const sh = ss.getSheetByName("instrumentos") || ss.getSheets()[0];
  const data = sh.getDataRange().getValues();

  const headers = data[0].map(h => h.toString().trim().toLowerCase());
  const colTrim1 = headers.indexOf("primer trimestre");
  const colTrim2 = headers.indexOf("segundo trimestre");
  const colTrim3 = headers.indexOf("tercer trimestre");

  const instrumentos = { trim1: [], trim2: [], trim3: [] };
  for (let i = 1; i < data.length; i++) {
    if (colTrim1 >= 0 && data[i][colTrim1]) instrumentos.trim1.push(data[i][colTrim1]);
    if (colTrim2 >= 0 && data[i][colTrim2]) instrumentos.trim2.push(data[i][colTrim2]);
    if (colTrim3 >= 0 && data[i][colTrim3]) instrumentos.trim3.push(data[i][colTrim3]);
  }

  // Añadir una fila adicional “por si acaso”
  if (instrumentos.trim1.length) instrumentos.trim1.push("");
  if (instrumentos.trim2.length) instrumentos.trim2.push("");
  if (instrumentos.trim3.length) instrumentos.trim3.push("");

  return instrumentos;
}
