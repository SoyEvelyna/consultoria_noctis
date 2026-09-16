/**
 * Noctis Tracker — backend de Google Apps Script.
 *
 * Proyecto INDEPENDIENTE de Apps Script (script.google.com), creado desde la
 * cuenta Gmail personal (el Workspace de soyevelyna.com bloquea Web Apps) y
 * publicado como Web App. La web del tracker le habla por HTTP.
 *
 * DISEÑO:
 * - "01 I Plan de trabajo", "02 I Proceso de trabajo" y "03 I Definición
 *   métricas" son la base: la web las lee en cada pedido. Tareas y reuniones
 *   se ESCRIBEN ahí mismo (02 e 01); 03 es solo lectura.
 * - Las columnas de 02 se ubican por el texto del encabezado (PRIORIDAD,
 *   ÁREA, TEMA, TAREA, RESPONSABLE, INICIO, TIEMPO, CIERRE, ESTADO,
 *   OBSERVACIONES), así que se pueden mover sin romper la web.
 * - Las celdas con desplegable solo reciben valores de su lista: si un valor
 *   no entra, no se pierde — queda anotado en Observaciones.
 * - Los links del entregable (hasta 3) se guardan en OBSERVACIONES de 02: debajo
 *   del texto, un link por línea y cada uno clickeable.
 *   Las notas van a "WebApp - Notas" ("WebApp - Overrides" solo guarda tareas
 *   ocultas y links viejos), pestañas que el script crea si no existen.
 * - Resumen diario por email (18 h, hora de Argentina): compara las tareas y
 *   reuniones con la foto guardada en "WebApp - Resumen" y avisa lo que cambió.
 * - Autenticación: token compartido (ACCESS_TOKEN), el mismo que usa la web.
 */

var ACCESS_TOKEN = "nX7qK2vR9tLm4PzW8sYc3HdF6jBa1GeU";

/* Id de la Hoja de cálculo de Google "Noctis I Tablero: PLAN DE TRABAJO". */
var SHEET_ID = "1ELzLOSN7QbpKgmJ64WXPWhf_Yc1qR1h0UVcutlurhTI";
/* Se abre una sola vez por pedido: abrirla en cada lectura hacía lenta la carga. */
var SS_ = null;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(SHEET_ID)); }

var SHEET_ETAPA1 = "01 I Plan de trabajo";
var SHEET_PROCESO = "02 I Proceso de trabajo";
var SHEET_METRICAS = "03 I Definición métricas";
var SHEET_OVERRIDES = "WebApp - Overrides";
var SHEET_NOTES = "WebApp - Notas";

var TITULO_REUNION = "Encuentro I Estado del proceso de trabajo";

var SHEET_SCHEMAS = {};
SHEET_SCHEMAS[SHEET_OVERRIDES] = ["task_id", "estado", "link", "inicio", "cierre", "hidden", "updated_at"];
SHEET_SCHEMAS[SHEET_NOTES] = ["id", "text", "author", "createdAt"];

/* Valores posibles en la hoja para cada estado de la web (se usa el primero
   que exista en el desplegable de la celda). */
var ESTADO_CANDIDATOS = {
  "Por hacer": ["Pendiente", "Atrasada"],
  "En proceso": ["Proceso", "En proceso"],
  "En revisión": ["Revisar", "En revisión"],
  "Testear": ["Testear"],
  "Completado": ["Finalizada", "Finalizado"]
};
var ESTADOS_SHEET = ["Pendiente", "Proceso", "Revisar", "Testear", "Finalizada"];

/** EJECUTAR A MANO UNA VEZ (opcional): crea las pestañas WebApp si faltan. */
function crearPestanasWebApp() {
  ensureSheets_();
  Logger.log("Listo: " + Object.keys(SHEET_SCHEMAS).join(", "));
}

/* =====================================================================
   ENTRY POINTS
   ===================================================================== */

function doGet(e) {
  try {
    checkToken_(e.parameter.token);
    if (e.parameter.action !== "read") {
      return jsonOut_({ ok: false, error: "acción GET no soportada: " + e.parameter.action });
    }
    ensureSheets_();
    return jsonOut_({
      ok: true,
      seed: readSeed_(),
      overrides: readOverrides_(),
      notes: readSimpleRows_(SHEET_NOTES, SHEET_SCHEMAS[SHEET_NOTES]),
      customTasks: [],
      meetings: []
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var body = JSON.parse(e.postData.contents || "{}");
    checkToken_(body.token);
    ensureSheets_();
    return jsonOut_({ ok: true, result: handleAction_(body.action, body.payload || {}) });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function handleAction_(action, p) {
  switch (action) {
    case "setOverride": return setOverride_(p.taskId, p.patch || {});
    case "addNote": return addRow_(SHEET_NOTES, SHEET_SCHEMAS[SHEET_NOTES],
      Object.assign({ id: "n" + Date.now(), createdAt: nowIso_() }, p));
    case "deleteNote": return deleteRow_(SHEET_NOTES, p.id);
    case "updateTask": return updateTask_(p.id, p.fields || {});
    case "addTask": return addTask_(p.fields || {});
    case "deleteTask": return deleteTask_(p.id);
    case "addMeetingSheet": return addMeetingSheet_(p);
    case "deleteMeetingSheet": return deleteMeetingSheet_(p);
    case "migrarLinks": return migrarLinks_();
    case "addOpcion": return addOpcion_(p);
    case "migrarEstados": return migrarEstados_();
    case "guardarFoto": guardarFoto_(p.foto); return { guardada: true };
    case "probarResumen": return probarResumen_();
    case "activarResumen": return activarResumen_();
    case "desactivarResumen": return desactivarResumen_();
    default: throw new Error("acción POST no soportada: " + action);
  }
}

/* =====================================================================
   UTILS
   ===================================================================== */

function checkToken_(token) {
  if (!token || token !== ACCESS_TOKEN) throw new Error("token inválido");
}
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function nowIso_() { return new Date().toISOString(); }

/* Texto de celda limpio: sin espacios de más ni caracteres invisibles. */
function cell_(row, idx) {
  if (!row) return null;
  var v = row[idx];
  if (v === null || v === undefined) return null;
  if (typeof v === "string") {
    v = v.replace(/[​-‍﻿]/g, "").trim();
    return v === "" ? null : v;
  }
  return v;
}

/* Para comparar textos: mayúsculas, sin tildes ni espacios extra. */
function norm_(s) {
  return String(s === null || s === undefined ? "" : s)
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[​-‍﻿]/g, "").replace(/\s+/g, " ").trim().toUpperCase();
}

function toIsoDate_(value) {
  if (value === "" || value === null || value === undefined) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(value).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var y = m[3].length === 2 ? "20" + m[3] : m[3];
    return y + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  }
  return s;
}

/* "yyyy-mm-dd" -> fecha a mediodía, para que no se corra de día por zona horaria. */
function toSheetDate_(iso) {
  if (!iso) return "";
  var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return iso;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

/* Mismo id que calcula la web: hash del contenido + contador de repetidas. */
function hashId_(str) {
  var h = 5381;
  for (var i = 0; i < str.length; i++) h = (((h << 5) + h) ^ str.charCodeAt(i)) >>> 0;
  return "t" + h.toString(36);
}

function findLabelRow_(values, label, fromRow) {
  var target = norm_(label);
  for (var r = fromRow || 0; r < values.length; r++) {
    for (var c = 0; c < Math.min(3, values[r].length); c++) {
      if (norm_(values[r][c]) === target) return r;
    }
  }
  return -1;
}

/* Opciones del desplegable de una celda (null si no tiene lista). */
function listOptions_(cell) {
  var dv = cell.getDataValidation();
  if (!dv || dv.getAllowInvalid()) return null;
  var type = dv.getCriteriaType();
  var crit = dv.getCriteriaValues();
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return crit[0];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return crit[0].getValues().map(function (r) { return r[0]; }).filter(function (v) { return v !== ""; });
  }
  return null;
}

/* Escribe el primer candidato que la celda acepte. Devuelve false si ninguno entra. */
function setSafe_(cell, candidates) {
  var options = listOptions_(cell);
  var value = candidates[0];
  if (options) {
    var hit = null;
    candidates.forEach(function (c) {
      if (hit !== null) return;
      options.forEach(function (o) { if (hit === null && norm_(o) === norm_(c)) hit = o; });
    });
    if (hit === null) return false;
    value = hit;
  }
  try { cell.setValue(value); return true; } catch (err) { return false; }
}

/* =====================================================================
   PESTAÑAS WEBAPP (overrides y notas)
   ===================================================================== */

function ensureSheets_() {
  var ss = ss_();
  Object.keys(SHEET_SCHEMAS).forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      var sheet = ss.insertSheet(name);
      sheet.appendRow(SHEET_SCHEMAS[name]);
      sheet.setFrozenRows(1);
    }
  });
}

function readSimpleRows_(sheetName, cols, dateCols) {
  var values = ss_().getSheetByName(sheetName).getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (!cell_(values[r], 0)) continue;
    var obj = {};
    cols.forEach(function (c, i) {
      var v = cell_(values[r], i);
      if (dateCols && dateCols.indexOf(c) !== -1) v = toIsoDate_(v);
      obj[c] = v;
    });
    out.push(obj);
  }
  return out;
}

function readOverrides_() {
  var out = {};
  readSimpleRows_(SHEET_OVERRIDES, SHEET_SCHEMAS[SHEET_OVERRIDES], ["inicio", "cierre"]).forEach(function (r) {
    out[r.task_id] = { estado: r.estado, link: r.link, inicio: r.inicio, cierre: r.cierre,
      hidden: r.hidden === true || r.hidden === "true", updatedAt: r.updated_at };
  });
  return out;
}

function findRowIndexById_(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(id)) return r + 1;
  }
  return -1;
}

function addRow_(sheetName, cols, obj) {
  ss_().getSheetByName(sheetName).appendRow(cols.map(function (c) { return obj[c] === undefined ? "" : obj[c]; }));
  return obj;
}

function deleteRow_(sheetName, id) {
  var sheet = ss_().getSheetByName(sheetName);
  var rowIdx = findRowIndexById_(sheet, id);
  if (rowIdx !== -1) sheet.deleteRow(rowIdx);
  return { id: id, deleted: rowIdx !== -1 };
}

function setOverride_(taskId, patch) {
  var sheet = ss_().getSheetByName(SHEET_OVERRIDES);
  var cols = SHEET_SCHEMAS[SHEET_OVERRIDES];
  var rowIdx = findRowIndexById_(sheet, taskId);
  if (rowIdx === -1) {
    var row = { task_id: taskId, estado: "", link: "", inicio: "", cierre: "", hidden: false, updated_at: nowIso_() };
    Object.keys(patch).forEach(function (k) { row[k] = patch[k] === null ? "" : patch[k]; });
    addRow_(SHEET_OVERRIDES, cols, row);
  } else {
    Object.keys(patch).forEach(function (key) {
      var colIdx = cols.indexOf(key);
      if (colIdx !== -1) sheet.getRange(rowIdx, colIdx + 1).setValue(patch[key] === null ? "" : patch[key]);
    });
    sheet.getRange(rowIdx, cols.indexOf("updated_at") + 1).setValue(nowIso_());
  }
  return { taskId: taskId, patch: patch };
}

/* El id cambia si cambia el texto: el link guardado pasa al id nuevo. */
function migrateOverride_(oldId, newId, fields) {
  var sheet = ss_().getSheetByName(SHEET_OVERRIDES);
  var cols = SHEET_SCHEMAS[SHEET_OVERRIDES];
  var rowIdx = findRowIndexById_(sheet, oldId);
  if (rowIdx !== -1) {
    // El link ahora vive en OBSERVACIONES de 02: si se editó, se limpia del override.
    var link = fields.link !== undefined ? "" : sheet.getRange(rowIdx, cols.indexOf("link") + 1).getValue();
    sheet.getRange(rowIdx, 1, 1, cols.length).setValues([[newId, "", link, "", "", false, nowIso_()]]);
  }
}

/* =====================================================================
   LECTURA
   ===================================================================== */

/* La hoja 02 y la 01 se leen una sola vez y se reutilizan en todo el pedido. */
function readSeed_() {
  var L = procesoLayout_();
  var T = etapa1Table_();
  var proceso = readProceso_(L);
  proceso.etapa1 = readEtapa1_(T);
  proceso.metricas = readMetricas_();
  proceso.opciones = readOpciones_(L, T);
  return proceso;
}

/* Valores de un desplegable de la Hoja (sin importar si acepta otros). */
function dropdownValues_(cell) {
  var dv = cell.getDataValidation();
  if (!dv) return [];
  var type = dv.getCriteriaType();
  var crit = dv.getCriteriaValues();
  var vals = [];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) vals = crit[0];
  else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    vals = crit[0].getValues().map(function (r) { return r[0]; });
  }
  return vals.map(function (v) { return String(v).trim(); }).filter(Boolean);
}

/* Opciones de los desplegables, para que la web ofrezca exactamente las mismas:
   si se suma un responsable o un área en la Hoja, aparece en la web al recargar. */
function readOpciones_(L, T) {
  var out = { responsable: [], area: [], responsableReuniones: [] };
  try {
    L = L || procesoLayout_();
    var row = L.tasks.length ? L.tasks[0].row : L.header + 2;
    ["responsable", "area"].forEach(function (k) {
      if (L.cols[k] !== undefined) out[k] = dropdownValues_(L.sheet.getRange(row, L.cols[k] + 1));
    });
  } catch (err) {}
  try {
    T = T || etapa1Table_();
    out.responsableReuniones = dropdownValues_(T.sheet.getRange(T.header + 2, 4));
  } catch (err2) {}
  return out;
}

function etapa1Table_() {
  var sheet = ss_().getSheetByName(SHEET_ETAPA1);
  if (!sheet) throw new Error("No encuentro la hoja '" + SHEET_ETAPA1 + "'");
  var values = sheet.getDataRange().getValues();
  var header = findLabelRow_(values, "Fecha");
  if (header === -1) throw new Error("No encuentro la tabla de reuniones (encabezado 'Fecha') en '" + SHEET_ETAPA1 + "'");
  var last = header;
  while (last + 1 < values.length && cell_(values[last + 1], 0)) last++;
  return { sheet: sheet, values: values, header: header, last: last };
}

function readEtapa1_(T) {
  T = T || etapa1Table_();
  var out = [];
  for (var r = T.header + 1; r <= T.last; r++) {
    var row = T.values[r];
    out.push({ fecha: toIsoDate_(row[0]), hs: cell_(row, 1), tarea: cell_(row, 2), responsable: cell_(row, 3),
      estado: cell_(row, 4), resultado: cell_(row, 5), obs: cell_(row, 6) });
  }
  return out;
}

var COLS_02 = {
  prioridad: "PRIORIDAD", area: "AREA", tema: "TEMA", tarea: "TAREA", responsable: "RESPONSABLE",
  inicio: "INICIO", tiempo: "TIEMPO", cierre: "CIERRE", estado: "ESTADO", obs: "OBSERVACIONES"
};

/* Ubica encabezado, columnas y tareas de 02 (mismo recorrido que usa la web para los ids). */
function procesoLayout_() {
  var sheet = ss_().getSheetByName(SHEET_PROCESO);
  if (!sheet) throw new Error("No encuentro la hoja '" + SHEET_PROCESO + "'");
  var values = sheet.getDataRange().getValues();
  var iniciativasRow = findLabelRow_(values, "INICIATIVAS");
  var header = -1;
  for (var r = Math.max(0, iniciativasRow); r < values.length && header === -1; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (norm_(values[r][c]).indexOf("TAREA") === 0) { header = r; break; }
    }
  }
  if (header === -1) throw new Error("No encuentro el encabezado de tareas (TAREA) en '" + SHEET_PROCESO + "'");
  var cols = {};
  values[header].forEach(function (v, i) {
    var h = norm_(v);
    Object.keys(COLS_02).forEach(function (k) {
      if (cols[k] !== undefined) return;
      if (k === "tarea" ? h.indexOf("TAREA") === 0 : h === COLS_02[k]) cols[k] = i;
    });
  });
  ["area", "tarea", "estado"].forEach(function (k) {
    if (cols[k] === undefined) throw new Error("Falta la columna " + COLS_02[k] + " en '" + SHEET_PROCESO + "'");
  });
  var fin = findLabelRow_(values, "FINALIZADOS", header + 1);
  var end = fin === -1 ? values.length : fin;
  var tasks = [], seen = {}, last = header;
  function get(row, k) { return cols[k] === undefined ? null : cell_(row, cols[k]); }
  for (var tr = header + 1; tr < end; tr++) {
    var row = values[tr];
    var area = get(row, "area"), tarea = get(row, "tarea");
    if (!area && !tarea) continue;
    var key = ["iniciativa", area || "", get(row, "tema") || "", tarea || ""].join("|");
    seen[key] = (seen[key] || 0) + 1;
    tasks.push({ id: hashId_(key + "#" + seen[key]), row: tr + 1 });
    last = tr;
  }
  return { sheet: sheet, values: values, cols: cols, header: header, tasks: tasks, lastTaskRow: last + 1, get: get };
}

function readProceso_(L) {
  L = L || procesoLayout_();
  var values = L.values;
  var objetivoRow = findLabelRow_(values, "OBJETIVO 1");
  var prioridadesRow = findLabelRow_(values, "PRIORIDADES");
  var iniciativasRow = findLabelRow_(values, "INICIATIVAS");

  var objetivo = "";
  if (objetivoRow !== -1) {
    for (var r = objetivoRow + 1; r < (prioridadesRow === -1 ? values.length : prioridadesRow); r++) {
      var v = cell_(values[r], 0);
      if (v && String(v).charAt(0) !== "¿") { objetivo = String(v); break; }
    }
  }

  var prioridades = [];
  if (prioridadesRow !== -1) {
    for (var p = prioridadesRow + 1; p < (iniciativasRow === -1 ? L.header : iniciativasRow); p++) {
      var n = cell_(values[p], 0), txt = cell_(values[p], 1);
      if (n === null || txt === null || isNaN(Number(n))) continue;
      var lines = String(txt).split(/\n+/).map(function (s) { return s.trim(); }).filter(Boolean);
      prioridades.push({ n: Number(n), tt: lines[0] || "", desc: lines.slice(1).join(" ") });
    }
  }

  var rich = obsRich_(L);
  var iniciativas = L.tasks.map(function (t) {
    var row = values[t.row - 1];
    var obsText = L.get(row, "obs");
    var links = rich ? obsLinks_(rich[t.row - 1][0], obsText) : [];
    var link = links.length ? links.join("\n") : null;
    return {
      link: link,
      prioridad: L.get(row, "prioridad"), area: L.get(row, "area"), tema: L.get(row, "tema"),
      tarea: L.get(row, "tarea"), responsable: L.get(row, "responsable"),
      inicio: toIsoDate_(L.get(row, "inicio")), tiempo: toIsoDate_(L.get(row, "tiempo")),
      cierre: toIsoDate_(L.get(row, "cierre")), estado: L.get(row, "estado"),
      obs: obsSinLinks_(obsText, links)
    };
  });

  return { objetivo: objetivo, prioridades: prioridades, iniciativas: iniciativas, finalizados: [], backlog: [] };
}

var MAX_LINKS = 3;

/* Links de OBSERVACIONES (hasta 3): textos enlazados en la celda y URLs escritas. */
function obsLinks_(richValue, text) {
  var out = [];
  function add(u) { if (u && out.indexOf(u) === -1 && out.length < MAX_LINKS) out.push(u); }
  if (richValue) {
    add(richValue.getLinkUrl());
    richValue.getRuns().forEach(function (run) { add(run.getLinkUrl()); });
  }
  (String(text || "").match(/https?:\/\/\S+/g) || []).forEach(add);
  return out;
}

/* Texto de OBSERVACIONES sin las líneas que son solo un link. */
function obsSinLinks_(text, links) {
  if (!text) return null;
  var lines = String(text).split(/\n/).filter(function (line) {
    var l = line.trim();
    return l && links.indexOf(l) === -1 && !/^https?:\/\/\S+$/.test(l);
  });
  var t = lines.join("\n").trim();
  return t || null;
}

function obsRich_(L) {
  if (L.cols.obs === undefined) return null;
  return L.sheet.getRange(1, L.cols.obs + 1, L.values.length, 1).getRichTextValues();
}

function readMetricas_() {
  var sheet = ss_().getSheetByName(SHEET_METRICAS);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var header = -1;
  for (var r = 0; r < values.length && header === -1; r++) {
    for (var c = 0; c < values[r].length; c++) if (norm_(values[r][c]) === "METRICA") { header = r; break; }
  }
  if (header === -1) return [];
  var idx = {};
  values[header].forEach(function (v, i) {
    var h = norm_(v);
    if (h === "PRIORIDAD") idx.prioridad = i;
    else if (h === "METRICA") idx.metrica = i;
    else if (h.indexOf("QUE NOS DICE") === 0) idx.dice = i;
    else if (h.indexOf("COMO SE MIDE") === 0) idx.mide = i;
    else if (h === "FRECUENCIA") idx.frecuencia = i;
    else if (h.indexOf("FUENTE") === 0) idx.fuente = i;
  });
  var out = [], prio = null;
  for (var mr = header + 1; mr < values.length; mr++) {
    var row = values[mr];
    if (idx.prioridad !== undefined && cell_(row, idx.prioridad) !== null) prio = cell_(row, idx.prioridad);
    var metrica = cell_(row, idx.metrica);
    if (!metrica) continue;
    out.push({ prioridad: prio, metrica: metrica, dice: cell_(row, idx.dice), mide: cell_(row, idx.mide),
      frecuencia: cell_(row, idx.frecuencia), fuente: cell_(row, idx.fuente) });
  }
  return out;
}

/* =====================================================================
   ESCRITURA DE TAREAS EN "02"
   ===================================================================== */

function findTask_(L, id) {
  for (var i = 0; i < L.tasks.length; i++) if (L.tasks[i].id === id) return L.tasks[i];
  return null;
}

function idAtRow_(row) {
  var L = procesoLayout_();
  for (var i = 0; i < L.tasks.length; i++) if (L.tasks[i].row === row) return L.tasks[i].id;
  return null;
}

/* Escribe los campos en su columna. Lo que un desplegable rechaza va a Observaciones. */
function writeTaskCells_(L, row, fields) {
  var perdidos = [];
  Object.keys(COLS_02).forEach(function (k) {
    if (k === "obs" || fields[k] === undefined || L.cols[k] === undefined) return;
    var cell = L.sheet.getRange(row, L.cols[k] + 1);
    var v = fields[k];
    if (v === null || v === "") { try { cell.setValue(""); } catch (err) {} return; }
    var candidates = k === "estado" ? (ESTADO_CANDIDATOS[v] || [v])
      : (k === "inicio" || k === "cierre") ? [toSheetDate_(v)] : [v];
    if (!setSafe_(cell, candidates)) perdidos.push(COLS_02[k].charAt(0) + COLS_02[k].slice(1).toLowerCase() + ": " + v);
  });
  // OBSERVACIONES guarda el texto y, si hay, el link del entregable (texto enlazado).
  if (L.cols.obs !== undefined && (fields.obs !== undefined || fields.link !== undefined || perdidos.length)) {
    // Queda: texto de observaciones y, debajo, un link por línea (cada uno clickeable).
    var obsCell = L.sheet.getRange(row, L.cols.obs + 1);
    var currentText = String(obsCell.getValue() || "");
    var currentLinks = obsLinks_(obsCell.getRichTextValue(), currentText);
    var base = fields.obs !== undefined ? (fields.obs || "") : (obsSinLinks_(currentText, currentLinks) || "");
    base = [base].concat(perdidos).filter(Boolean).join(" | ");
    var links = fields.link !== undefined
      ? String(fields.link || "").split(/\s*\n\s*|\s+\|\s+/).map(function (s) { return s.trim(); }).filter(Boolean).slice(0, MAX_LINKS)
      : currentLinks;
    var text = [base].concat(links).filter(Boolean).join("\n");
    var builder = SpreadsheetApp.newRichTextValue().setText(text);
    var pos = base ? base.length + 1 : 0;
    links.forEach(function (url) {
      builder = builder.setLinkUrl(pos, pos + url.length, url);
      pos += url.length + 1;
    });
    obsCell.setRichTextValue(builder.build());
  }
  return perdidos;
}

function updateTask_(id, fields) {
  var L = procesoLayout_();
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en '" + SHEET_PROCESO + "'. Puede haber cambiado en el Sheet: recargá la página.");
  var perdidos = writeTaskCells_(L, t.row, fields);
  SpreadsheetApp.flush();
  var newId = idAtRow_(t.row);
  migrateOverride_(id, newId, fields);
  return { id: newId, enObservaciones: perdidos };
}

function addTask_(fields) {
  var L = procesoLayout_();
  var after = L.lastTaskRow;
  var width = L.sheet.getMaxColumns();
  L.sheet.insertRowAfter(after);
  var row = after + 1;
  var template = L.sheet.getRange(after, 1, 1, width);
  var target = L.sheet.getRange(row, 1, 1, width);
  template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  template.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  target.clearContent();
  var L2 = procesoLayout_();
  var perdidos = writeTaskCells_(L2, row, Object.assign({ estado: "Por hacer" }, fields));
  SpreadsheetApp.flush();
  var newId = idAtRow_(row);
  return { id: newId, enObservaciones: perdidos };
}

function deleteTask_(id) {
  var L = procesoLayout_();
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en '" + SHEET_PROCESO + "'. Recargá la página.");
  L.sheet.deleteRow(t.row);
  deleteRow_(SHEET_OVERRIDES, id);
  return { deleted: id };
}

/* =====================================================================
   REUNIONES EN "01"
   ===================================================================== */

/* "2 hs" -> 2 ; "1.30" se deja como texto. */
function horas_(duracion) {
  var m = String(duracion || "").match(/\d+(?:[.,]\d+)?/);
  if (!m) return "";
  return /[.,]/.test(m[0]) ? m[0] : Number(m[0]);
}

/* Si debajo de la última reunión hay un encuentro previsto (sin fecha), se
   completa ese; si no, se agrega una fila copiando formato y desplegables. */
function addMeetingSheet_(p) {
  var T = etapa1Table_();
  var sheet = T.sheet;
  var nextIdx = T.last + 1;
  var previsto = nextIdx < T.values.length && !cell_(T.values[nextIdx], 0) && cell_(T.values[nextIdx], 2);
  var row;
  if (previsto) {
    row = nextIdx + 1;
  } else {
    sheet.insertRowAfter(T.last + 1);
    row = T.last + 2;
    sheet.getRange(T.last + 1, 1, 1, 7).copyTo(sheet.getRange(row, 1, 1, 7));
  }
  var tarea = (p.tarea && p.tarea !== TITULO_REUNION) ? p.tarea
    : (previsto ? cell_(T.values[nextIdx], 2) : (p.tarea || TITULO_REUNION));

  var perdidos = [];
  function put(col, label, candidates, keepIfEmpty) {
    var cell = sheet.getRange(row, col);
    var v = candidates[0];
    if (v === "" || v === null || v === undefined) {
      if (!keepIfEmpty) cell.setValue("");
      return;
    }
    if (!setSafe_(cell, candidates)) perdidos.push(label + ": " + (v instanceof Date ? p.fecha : v));
  }
  put(1, "Fecha", [toSheetDate_(p.fecha)], false);
  put(2, "Hs", [horas_(p.duracion)], true);
  put(3, "Tarea", [tarea], false);
  put(4, "Responsable", [p.responsable], true);
  put(5, "Estado", ["Finalizado", "Finalizada"], true);
  put(6, "Resultado", [p.resumen], false);
  sheet.getRange(row, 7).setValue(perdidos.join(" | "));
  SpreadsheetApp.flush();
  return { row: row, guardado: sheet.getRange(row, 1, 1, 7).getDisplayValues()[0], enObservaciones: perdidos };
}

function deleteMeetingSheet_(p) {
  var T = etapa1Table_();
  for (var r = T.header + 1; r <= T.last; r++) {
    var row = T.values[r];
    if (toIsoDate_(row[0]) === p.fecha &&
        String(cell_(row, 2) || "") === String(p.tarea || "") &&
        String(cell_(row, 5) || "") === String(p.resultado || "")) {
      T.sheet.deleteRow(r + 1);
      return { deleted: true };
    }
  }
  throw new Error("No encuentro esa reunión en '" + SHEET_ETAPA1 + "'. Recargá la página.");
}

/* Pasa a OBSERVACIONES de 02 los links que quedaron en "WebApp - Overrides". */
function migrarLinks_() {
  var overrides = readOverrides_();
  var L = procesoLayout_();
  var log = [];
  Object.keys(overrides).forEach(function (id) {
    var link = overrides[id].link;
    if (!link) return;
    var t = findTask_(L, id);
    if (!t) { log.push(id + ": no está en 02"); return; }
    writeTaskCells_(L, t.row, { link: link });
    migrateOverride_(id, id, { link: null });
    log.push(id + " -> link en OBSERVACIONES");
  });
  return log;
}

/* Suma un nombre al desplegable de RESPONSABLE de 02 (toda la columna debajo del
   encabezado), para poder asignarlo desde la web. Si ya estaba, no cambia nada. */
function addOpcion_(p) {
  var valor = String(p.valor || "").trim();
  if (!valor) throw new Error("Falta el nombre");
  if (p.campo !== "responsable") throw new Error("Campo no soportado: " + p.campo);
  var L = procesoLayout_();
  if (L.cols.responsable === undefined) throw new Error("No encuentro la columna RESPONSABLE en '" + SHEET_PROCESO + "'");
  var col = L.cols.responsable + 1;
  var firstRow = L.header + 2;
  var sample = L.sheet.getRange(L.tasks.length ? L.tasks[0].row : firstRow, col);
  var actuales = dropdownValues_(sample);
  if (actuales.map(norm_).indexOf(norm_(valor)) !== -1) return { valor: valor, yaEstaba: true, opciones: actuales };
  var nuevas = actuales.concat([valor]);
  var dv = sample.getDataValidation();
  var builder = dv ? dv.copy() : SpreadsheetApp.newDataValidation().setAllowInvalid(false);
  var rule = builder.requireValueInList(nuevas, true).build();
  L.sheet.getRange(firstRow, col, L.sheet.getMaxRows() - firstRow + 1, 1).setDataValidation(rule);
  return { valor: valor, opciones: nuevas };
}

/* Estado viejo de la hoja -> estado nuevo (Pendiente, Proceso, Revisar, Testear, Finalizada). */
function estadoNuevo_(v) {
  var s = String(v || "").trim().toLowerCase();
  if (s.indexOf("final") === 0) return "Finalizada";
  if (s.indexOf("en proceso") === 0 || s.indexOf("proceso") === 0 || s.indexOf("actualiz") === 0) return "Proceso";
  if (s.indexOf("revis") === 0 || s.indexOf("propuesta") === 0) return "Revisar";
  if (s.indexOf("test") === 0) return "Testear";
  return "Pendiente";
}

/* Una vez: deja el desplegable de ESTADO de 02 con las 5 opciones nuevas (toda la
   columna debajo del encabezado) y pasa cada tarea a su estado nuevo. */
function migrarEstados_() {
  var L = procesoLayout_();
  if (L.cols.estado === undefined) throw new Error("No encuentro la columna ESTADO en '" + SHEET_PROCESO + "'");
  var col = L.cols.estado + 1;
  var firstRow = L.header + 2;
  var sample = L.sheet.getRange(L.tasks.length ? L.tasks[0].row : firstRow, col);
  var dv = sample.getDataValidation();
  var builder = dv ? dv.copy() : SpreadsheetApp.newDataValidation().setAllowInvalid(false);
  L.sheet.getRange(firstRow, col, L.sheet.getMaxRows() - firstRow + 1, 1)
    .setDataValidation(builder.requireValueInList(ESTADOS_SHEET, true).build());
  var cambios = {};
  L.tasks.forEach(function (t) {
    var cell = L.sheet.getRange(t.row, col);
    var antes = String(cell.getValue() || "");
    var despues = estadoNuevo_(antes);
    if (antes !== despues) cell.setValue(despues);
    var k = (antes || "(vacío)") + " -> " + despues;
    cambios[k] = (cambios[k] || 0) + 1;
  });
  return cambios;
}

/* =====================================================================
   RESUMEN DIARIO POR EMAIL (18 h)
   Compara las tareas y reuniones de hoy con la "foto" guardada en el último
   resumen y manda un email con lo que cambió (desde la web o a mano en la
   Hoja). Si no cambió nada, no manda nada.
   ===================================================================== */

var RESUMEN_DESTINATARIOS = [
  "hola@soyevelyna.com",
  "marketingdigital.push@gmail.com",
  "leandro.pachame@gmail.com",
  "ramirezaldana97@gmail.com",
  "manataliapereira.NP@gmail.com"
];
var RESUMEN_PRUEBA = "hola@soyevelyna.com";
var RESUMEN_ZONA = "America/Argentina/Buenos_Aires";
var RESUMEN_HORA = 18;
var TABLERO_URL = "https://soyevelyna.github.io/consultoria_noctis/";
var SHEET_RESUMEN = "WebApp - Resumen";

function fotoActual_() {
  var seed = readSeed_();
  return {
    fecha: nowIso_(),
    tareas: seed.iniciativas.map(function (t) {
      return { prioridad: t.prioridad, area: t.area, tema: t.tema, tarea: t.tarea, responsable: t.responsable,
        inicio: t.inicio, cierre: t.cierre, estado: t.estado, obs: t.obs };
    }),
    reuniones: seed.etapa1.map(function (m) {
      return { fecha: m.fecha, tarea: m.tarea, responsable: m.responsable, resultado: m.resultado };
    })
  };
}

function leerFoto_() {
  var sh = ss_().getSheetByName(SHEET_RESUMEN);
  if (!sh) return null;
  var v = sh.getRange(1, 1).getValue();
  try { return v ? JSON.parse(v) : null; } catch (err) { return null; }
}

function guardarFoto_(foto) {
  var texto = JSON.stringify(foto);
  if (texto.length > 49000) throw new Error("La foto del resumen es demasiado grande para una celda");
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_RESUMEN) || ss.insertSheet(SHEET_RESUMEN);
  sh.getRange(1, 1).setValue(texto);
  sh.getRange(2, 1).setValue("Foto que usa el resumen diario por email. No editar. Última: " + (foto.fecha || nowIso_()));
}

function claveTarea_(t) { return String(t.tarea || "").trim().toLowerCase(); }

function difResumen_(antes, ahora) {
  var r = { nuevas: [], finalizadas: [], estados: [], otros: [], eliminadas: [], reuniones: [] };
  var mapaAntes = {};
  (antes.tareas || []).forEach(function (t) { mapaAntes[claveTarea_(t)] = t; });
  var vistas = {};
  var campos = [["responsable", "Responsable"], ["inicio", "Inicio"], ["cierre", "Cierre"], ["prioridad", "Prioridad"],
    ["area", "Área"], ["tema", "Tema"], ["obs", "Observaciones"]];
  (ahora.tareas || []).forEach(function (t) {
    var k = claveTarea_(t);
    vistas[k] = true;
    var p = mapaAntes[k];
    if (!p) { r.nuevas.push(t); return; }
    var eAntes = String(p.estado || ""), eAhora = String(t.estado || "");
    if (eAntes !== eAhora) {
      if (/^final/i.test(eAhora)) r.finalizadas.push(t);
      else r.estados.push({ t: t, de: eAntes || "(sin estado)", a: eAhora || "(sin estado)" });
    }
    var cambios = [];
    campos.forEach(function (c) {
      var a = p[c[0]] == null ? "" : String(p[c[0]]);
      var b = t[c[0]] == null ? "" : String(t[c[0]]);
      if (a !== b) cambios.push(c[1] + ": " + (a || "—") + " → " + (b || "—"));
    });
    if (cambios.length) r.otros.push({ t: t, cambios: cambios });
  });
  (antes.tareas || []).forEach(function (p) { if (!vistas[claveTarea_(p)]) r.eliminadas.push(p); });
  var reunAntes = {};
  (antes.reuniones || []).forEach(function (m) { reunAntes[m.fecha + "|" + m.tarea] = true; });
  (ahora.reuniones || []).forEach(function (m) { if (!reunAntes[m.fecha + "|" + m.tarea]) r.reuniones.push(m); });
  r.total = r.nuevas.length + r.finalizadas.length + r.estados.length + r.otros.length + r.eliminadas.length + r.reuniones.length;
  return r;
}

function htmlResumen_(r, hoy, prueba) {
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function corto(s) { s = String(s == null ? "" : s); return s.length > 110 ? s.slice(0, 107) + "…" : s; }
  function linea(t, extra) {
    return "<li style='margin-bottom:6px'><b>" + esc(t.tarea) + "</b>" + (t.responsable ? " · " + esc(t.responsable) : "") +
      (extra ? "<br><span style='color:#555'>" + extra + "</span>" : "") + "</li>";
  }
  function seccion(titulo, color, items) {
    if (!items.length) return "";
    return "<h3 style='margin:20px 0 6px;font-family:Arial,sans-serif;font-size:16px;color:" + color + "'>" + titulo + " (" + items.length + ")</h3>" +
      "<ul style='margin:0;padding-left:18px;font-family:Arial,sans-serif;font-size:14px;line-height:1.45'>" + items.join("") + "</ul>";
  }
  var html = "<div style='max-width:640px'>";
  if (prueba) {
    html += "<p style='background:#FFF3CD;padding:8px 10px;font-family:Arial,sans-serif;font-size:13px'>Email de prueba: así se va a ver el resumen diario de las 18 h. Por ahora solo te llega a vos.</p>";
  }
  html += "<h2 style='font-family:Arial,sans-serif;margin:0 0 4px'>Noctis · Resumen del " + esc(hoy) + "</h2>";
  html += "<p style='font-family:Arial,sans-serif;color:#555;margin:0'>Cambios en el tablero desde el último resumen.</p>";
  html += seccion("Tareas nuevas", "#212121", r.nuevas.map(function (t) { return linea(t, esc(t.estado || "")); }));
  html += seccion("Finalizadas", "#23914A", r.finalizadas.map(function (t) { return linea(t); }));
  html += seccion("Cambios de estado", "#7447B8", r.estados.map(function (x) { return linea(x.t, esc(x.de) + " → " + esc(x.a)); }));
  html += seccion("Otros cambios", "#1F7FB5", r.otros.map(function (x) {
    return linea(x.t, x.cambios.map(function (c) { return esc(corto(c)); }).join("<br>"));
  }));
  html += seccion("Eliminadas", "#B9282D", r.eliminadas.map(function (t) { return linea(t); }));
  html += seccion("Reuniones nuevas", "#212121", r.reuniones.map(function (m) {
    return "<li style='margin-bottom:6px'><b>" + esc(m.tarea) + "</b> · " + esc(m.fecha) +
      (m.resultado ? "<br><span style='color:#555'>" + esc(corto(m.resultado)) + "</span>" : "") + "</li>";
  }));
  if (!r.total) html += "<p style='font-family:Arial,sans-serif;margin-top:16px'>No hubo cambios.</p>";
  html += "<p style='font-family:Arial,sans-serif;margin-top:24px'><a href='" + TABLERO_URL + "'>Abrir el tablero de Noctis</a></p></div>";
  return html;
}

/* Para ejecutar a mano desde el editor (el menú no muestra funciones con "_").
   El resumen sale desde la cuenta dueña del proyecto que lo ejecuta: se usa
   una copia de este archivo en un proyecto de hola@soyevelyna.com (sin Web App). */
function quienEnvia() { Logger.log(JSON.stringify({ efectivo: Session.getEffectiveUser().getEmail(), activo: Session.getActiveUser().getEmail() })); }
function probarResumenAhora() { Logger.log(JSON.stringify(probarResumen_())); }
function activarResumenDiario() { Logger.log(JSON.stringify(activarResumen_())); }
function desactivarResumenDiario() { Logger.log(JSON.stringify(desactivarResumen_())); }

/* Lo ejecuta el activador diario (18 h). También guarda la foto nueva. */
function enviarResumenDiario() {
  var ahora = fotoActual_();
  var antes = leerFoto_();
  if (!antes) { guardarFoto_(ahora); return { enviado: false, motivo: "primera foto guardada" }; }
  var r = difResumen_(antes, ahora);
  if (r.total) {
    var hoy = Utilities.formatDate(new Date(), RESUMEN_ZONA, "dd/MM/yyyy");
    MailApp.sendEmail({
      to: RESUMEN_DESTINATARIOS.join(","),
      subject: "Noctis · Resumen del día " + hoy + " (" + r.total + (r.total === 1 ? " cambio)" : " cambios)"),
      htmlBody: htmlResumen_(r, hoy, false),
      name: "Tablero Noctis"
    });
  }
  guardarFoto_(ahora);
  return { enviado: !!r.total, cambios: r.total };
}

/* Prueba: manda el resumen solo a RESUMEN_PRUEBA, sin tocar la foto guardada. */
function probarResumen_() {
  var ahora = fotoActual_();
  var antes = leerFoto_() || { tareas: [], reuniones: [] };
  var r = difResumen_(antes, ahora);
  var hoy = Utilities.formatDate(new Date(), RESUMEN_ZONA, "dd/MM/yyyy");
  MailApp.sendEmail({
    to: RESUMEN_PRUEBA,
    subject: "[Prueba] Noctis · Resumen del día " + hoy + " (" + r.total + (r.total === 1 ? " cambio)" : " cambios)"),
    htmlBody: htmlResumen_(r, hoy, true),
    name: "Tablero Noctis"
  });
  return { enviadoA: RESUMEN_PRUEBA, cambios: r.total, cuotaRestante: MailApp.getRemainingDailyQuota() };
}

/* Crea (o recrea) el activador diario y arranca desde la foto de ahora. */
function activarResumen_() {
  desactivarResumen_();
  ScriptApp.newTrigger("enviarResumenDiario").timeBased().everyDays(1).atHour(RESUMEN_HORA).inTimezone(RESUMEN_ZONA).create();
  guardarFoto_(fotoActual_());
  return { activado: true, hora: RESUMEN_HORA, zona: RESUMEN_ZONA, destinatarios: RESUMEN_DESTINATARIOS };
}

function desactivarResumen_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === "enviarResumenDiario") { ScriptApp.deleteTrigger(tr); n++; }
  });
  return { desactivados: n };
}
