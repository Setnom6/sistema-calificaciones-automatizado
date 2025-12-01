/**
 * main.gs
 * Punto de entrada: genera/actualiza calificacionesN y mediasN para un trimestre.
 * Llama a calificacionesConstructor.buildCalificaciones y madieasConstructor.buildMedias.
 */

function generateTrimester(n) {
  const ss = SpreadsheetApp.getActive();
  const sheetList = ss.getSheetByName("listado");
  const sheetCriteria = ss.getSheetByName("criterios");
  const sheetInstruments = ss.getSheetByName("instrumentos");

  if (!sheetList || !sheetCriteria || !sheetInstruments) {
    SpreadsheetApp.getUi().alert("Faltan hojas necesarias: 'listado', 'criterios' o 'instrumentos'.");
    return;
  }

  // ---------- locate columns "TrimestreN" and "CriteriosN" in 'instrumentos' sheet ----------
  const hdrInst = sheetInstruments.getRange(1,1,1, sheetInstruments.getLastColumn()).getValues()[0];
  const colTrimestreIdx = hdrInst.indexOf("Trimestre" + n);
  const colCriteriosIdx = hdrInst.indexOf("Criterios" + n);

  if (colTrimestreIdx === -1 || colCriteriosIdx === -1) {
    SpreadsheetApp.getUi().alert("No se han encontrado las cabeceras 'Trimestre" + n + "' o 'Criterios" + n + "' en la hoja 'instrumentos'.");
    return;
  }

  const trimestreCol = colTrimestreIdx + 1;
  const criteriosCol = colCriteriosIdx + 1;

  // ---------- build student list ordered by surname (apellidos) ----------
  // Se lee hasta 3 columnas por si el listado tiene Nombre, Apellido1, Apellido2.
  const listadoLastRow = Math.max( sheetList.getLastRow(), 2 );
  const datosListado = sheetList.getRange(2,1, Math.max(0, listadoLastRow-1), 3).getValues();
  const alumnosRaw = datosListado
    .filter(r => r[0] && r[1])
    .map(r => {
      const nombres = (r[0] || "").toString().trim();
      const apellido1 = (r[1] || "").toString().trim();
      const apellido2 = (r[2] || "").toString().trim();
      // displayName: solo nombre + primer apellido (para mostrar y para emparejar con datos antiguos)
      const displayName = (nombres + " " + apellido1).trim();
      // surnameKey: usado solo para ordenar (apellido1 + apellido2 si existe)
      const surnameKey = (apellido1 + (apellido2 ? " " + apellido2 : "")).trim();
      // sourceRow: guardamos la fila original (columnas A..C leídas) para comparar igualdad completa
      const sourceRow = [ (r[0] || "").toString().trim(), (r[1] || "").toString().trim(), (r[2] || "").toString().trim() ];
      return { displayName, surnameKey, nombres, sourceRow };
    });

  // Ordenar por apellidos (clave), usando locale 'es' y desempatar por nombre para estabilidad
  alumnosRaw.sort((a, b) => {
    const cmp = a.surnameKey.localeCompare(b.surnameKey, 'es', { sensitivity: 'base', numeric: true });
    if (cmp !== 0) return cmp;
    return a.nombres.localeCompare(b.nombres, 'es', { sensitivity: 'base', numeric: true });
  });

  // Eliminar duplicados por displayName SOLO si toda la fila original (A..C) es idéntica.
  // Si hay dos entradas con mismo displayName pero datos distintos, conservamos ambas y lo dejamos al usuario.
  const seenMap = {}; // displayName -> [sourceRowString, ...]
  const uniqueAlumnosRaw = [];
  alumnosRaw.forEach(a => {
    const nameKey = a.displayName || "";
    const rowSig = (a.sourceRow || []).join("|");
    if (!seenMap[nameKey]) {
      seenMap[nameKey] = [rowSig];
      uniqueAlumnosRaw.push(a);
    } else {
      // ya existe al menos una fila con ese displayName: comprobar si alguna coincide exactamente
      const matches = seenMap[nameKey].some(sig => sig === rowSig);
      if (!matches) {
        // distinto contenido: conservarlo (posible homónimo)
        seenMap[nameKey].push(rowSig);
        uniqueAlumnosRaw.push(a);
        Logger.log(`Homónimo detectado para '${nameKey}' con distinto contenido; se conservan ambas filas.`);
      } else {
        // fila idéntica ya existente: ignoramos esta entrada (duplicado exacto)
        Logger.log(`Duplicado exacto detectado para '${nameKey}' — una de las filas idénticas será ignorada.`);
      }
    }
  });

  const alumnos = uniqueAlumnosRaw.map(a => [a.displayName]); // para setValues (solo nombre + primer apellido)

  // ---------- get instruments and their criteria (SORTED LEXICOGRAPHICALLY) ----------
  const instLastRow = Math.max( sheetInstruments.getLastRow(), 2 );
  const instNames = instLastRow - 1 > 0 ? sheetInstruments.getRange(2, trimestreCol, Math.max(0, instLastRow-1)).getValues().map(r=>r[0]) : [];
  const instCriteria = instLastRow - 1 > 0 ? sheetInstruments.getRange(2, criteriosCol, Math.max(0, instLastRow-1)).getValues().map(r=>r[0]) : [];

  let instrumentos = [];
  for (let i=0; i<instNames.length; i++){
    const name = instNames[i];
    if (name && name.toString().trim() !== "") {
      const criteriosCell = (instCriteria[i] || "").toString();
      
      // Clean and split criteria
      let criteriosList = criteriosCell === "" ? [] : criteriosCell.split(",").map(s=>s.trim()).filter(s=>s!=="");
      
      // SORT CRITERIA LEXICOGRAPHICALLY with numeric awareness
      if (criteriosList.length > 0) {
        criteriosList.sort((a, b) => {
          return a.localeCompare(b, 'es', { 
            numeric: true, 
            sensitivity: 'base'
          });
        });
        
        // DEBUG: Log the sorting result
        Logger.log(`Instrument: ${name}`);
        Logger.log(`Original: ${criteriosCell}`);
        Logger.log(`Sorted: ${criteriosList.join(', ')}`);
      }
      
      if (criteriosList.length > 0) {
        instrumentos.push({ nombre: name.toString().trim(), criterios: criteriosList });
      }

    }
  }

  // ---------- build mapping clave->color from 'criterios' sheet (same logic as original) ----------
  const criteriosHdr = sheetCriteria.getRange(1,1,1,sheetCriteria.getLastColumn()).getValues()[0].map(h => h ? h.toString().trim().toLowerCase() : "");
  const colClaveIdx = criteriosHdr.indexOf("clave"); // 0-based
  let claveToColor = {};
  if (colClaveIdx !== -1) {
    const numCrit = Math.max(0, sheetCriteria.getLastRow() - 1);
    if (numCrit > 0) {
      const arrClaves = sheetCriteria.getRange(2, colClaveIdx+1, numCrit, 1).getValues().map(r=> r[0] ? r[0].toString().trim() : "");
      const arrColors = sheetCriteria.getRange(2, colClaveIdx+1, numCrit, 1).getBackgrounds().map(r=> r[0]);
      for (let i=0;i<arrClaves.length;i++){
        const clave = arrClaves[i];
        if (clave && clave !== "") claveToColor[clave] = arrColors[i];
      }
    }
  } else {
    // fallback: try column D (4)
    try {
      const numCrit = Math.max(0, sheetCriteria.getLastRow() - 1);
      const arrClaves = sheetCriteria.getRange(2,4, numCrit, 1).getValues().map(r=> r[0] ? r[0].toString().trim() : "");
      const arrColors = sheetCriteria.getRange(2,4, numCrit, 1).getBackgrounds().map(r=> r[0]);
      for (let i=0;i<arrClaves.length;i++){
        const clave = arrClaves[i];
        if (clave && clave !== "") claveToColor[clave] = arrColors[i];
      }
    } catch(e){}
  }

  // ---------- Call calificaciones constructor (build or update sheet calificacionesN) ----------
  const calificacionesResult = buildCalificaciones(n, alumnos, instrumentos, claveToColor);

  if (!calificacionesResult || !calificacionesResult.sheetCalif) {
    SpreadsheetApp.getUi().alert("Error construyendo/actualizando calificaciones" + n);
    return;
  }

  // ---------- Call medias constructor (build mediasN). It will compute columns for each clave  ----------
  sheetMedias = buildMedias(n, calificacionesResult.sheetCalif, alumnos, sheetCriteria, claveToColor);

  // ------------------ Call to get links --------------
  writeLinks(n, calificacionesResult.sheetCalif, sheetMedias);


  SpreadsheetApp.getUi().alert("Calificaciones y medias para Trimestre " + n + " generadas/actualizadas correctamente.");
}

/* helpers and exposed functions */

function trimester1(){ generateTrimester(1); }
function trimester2(){ generateTrimester(2); }
function trimester3(){ generateTrimester(3); }

function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    let temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }
  return letter;
}

function arraysEqual(a,b) {
  if (!a || !b) return false;
  if (a.length !== b.length) return false;
  for (let i=0;i<a.length;i++){
    const va = a[i] ? a[i].toString().trim() : "";
    const vb = b[i] ? b[i].toString().trim() : "";
    if (va !== vb) return false;
  }
  return true;
}

// ===== ENLACES EN HOJA INSTRUMENTOS =====
function writeLinks(n, sheetCalif, sheetMedias) {

  const ss = SpreadsheetApp.getActive();
  const sheetInstr = ss.getSheetByName("instrumentos");
  if (!sheetInstr) return;

  const califGid = sheetCalif.getSheetId();
  const mediasGid = sheetMedias.getSheetId();

  // mapa de posiciones
  const posiciones = {
    1: { calif: "K3",  medias: "K5"  },
    2: { calif: "K10", medias: "K12" },
    3: { calif: "K17", medias: "K19" }
  };

  if (!posiciones[n]) return;

  // hipervínculos internos usando el gid
  sheetInstr.getRange(posiciones[n].calif)
    .setFormula(`=HYPERLINK("#gid=${califGid}"; "calificaciones${n}")`);

  sheetInstr.getRange(posiciones[n].medias)
    .setFormula(`=HYPERLINK("#gid=${mediasGid}"; "medias${n}")`);
}