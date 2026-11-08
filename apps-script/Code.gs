/**
 * Puente entre el Laboratorio Virtual «Movimiento en dos dimensiones» (UTP) y una hoja de cálculo de Google.
 * © 2026 Luis Hernando Blandón Dávila. Todos los derechos reservados.
 * El propio script CREA la hoja de cálculo la primera vez que se ejecuta.
 *
 * Instalación (una sola vez):
 *  1. Abra https://script.new con su cuenta de Google → borre el contenido, pegue este código y guarde.
 *  2. (Opcional) Escriba en FOLDER_ID el ID de una carpeta (p. ej. de la unidad compartida del curso).
 *  3. Elija la función «configurar» en la barra superior y pulse ▶ Ejecutar. Autorice los permisos.
 *     Se crea la hoja «Calificaciones · Laboratorio Movimiento 2D» y su enlace aparece en el registro de ejecución.
 *  4. Implementar → Nueva implementación → tipo «Aplicación web».
 *     Ejecutar como: Yo · Quién tiene acceso: Cualquier persona → Implementar.
 *  5. Copie la URL que termina en /exec y péguela en la app (botón Docente).
 * Si cambia el código, use Implementar → Gestionar implementaciones → Editar → Versión nueva.
 * (También funciona si el script se abre desde una hoja con Extensiones → Apps Script: entonces usa esa hoja.)
 */
var SPREADSHEET_NAME = 'Calificaciones · Laboratorio Movimiento 2D';
var SHEET_NAME = 'Calificaciones';
var FOLDER_ID = '';   // Opcional: carpeta donde se crea la hoja y se guarda un comprobante JSON por envío

var HEADERS = ['Fecha y hora', 'ID envío', 'Intento', 'Estudiantes', 'Códigos', 'Grupo', 'Docente', 'Curso',
  'Nota (1–5)', 'Puntaje %', 'Datos %', 'Cálculos %', 'Conceptos %', 'Análisis %',
  'v₀ prom (m/s)', 'g ajuste (m/s²)', 'g guía (m/s²)', 'Error g %', 'R²', 'Ángulos completos', 'y₀ (m)',
  'Origen de los datos', 'Respuestas abiertas', 'Detalle (JSON)'];

/** Ejecútela UNA VEZ desde el editor (▶ Ejecutar): crea la hoja de cálculo y muestra su enlace. */
function configurar() {
  var ss = getSpreadsheet_();
  getSheet_();
  Logger.log('Hoja de calificaciones lista: ' + ss.getUrl());
  return ss.getUrl();
}

function doGet(e) {
  var sh = getSheet_();
  return json_({ ok: true, app: 'Laboratorio Movimiento 2D', sheet: sh.getParent().getName(), rows: Math.max(0, sh.getLastRow() - 1) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = JSON.parse(e.postData.contents);
    if (!p || !p.id) return json_({ ok: false, error: 'Envío sin ID' });
    var sh = getSheet_();
    var last = sh.getLastRow();
    if (last > 1) {
      var ids = sh.getRange(2, 2, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(p.id)) return json_({ ok: true, duplicate: true, row: i + 2 });
      }
    }
    var row = [new Date(), p.id, p.intento, p.estudiantes, p.codigos, p.grupo, p.docente, p.curso,
      Number(p.nota), p.porcentaje, p.s_datos, p.s_calculos, p.s_conceptos, p.s_analisis,
      p.v0, p.g_ajuste, p.g_guia, p.error_g, p.r2, p.angulos, p.y0, p.origen,
      String(p.respuestas_abiertas || '').slice(0, 45000), JSON.stringify(p.detalle || {}).slice(0, 45000)];
    sh.appendRow(row);
    var r = sh.getLastRow();
    sh.getRange(r, 9).setNumberFormat('0.0');
    if (FOLDER_ID) {
      try {
        DriveApp.getFolderById(FOLDER_ID).createFile('Taller_' + (p.grupo || 'grupo') + '_' + p.id + '.json', JSON.stringify(p, null, 2), MimeType.PLAIN_TEXT);
      } catch (err) { /* la fila ya quedó registrada */ }
    }
    return json_({ ok: true, row: r });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* Usa la hoja vinculada si existe; si no, la creada por este script (su ID se guarda en las propiedades); si no, la crea. */
function getSpreadsheet_() {
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (active) return active;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* la hoja se borró: se crea otra */ }
  }
  var ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  ss.getSheets()[0].setName(SHEET_NAME);
  if (FOLDER_ID) {
    try { DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(FOLDER_ID)); } catch (e) { /* queda en Mi unidad */ }
  }
  return ss;
}

function getSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1d5d8c').setFontColor('#ffffff');
    sh.setColumnWidth(4, 260);
    sh.setColumnWidth(23, 360);
    var notas = sh.getRange('I2:I');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(4).setBackground('#dcefe3').setRanges([notas]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(3, 3.99).setBackground('#f7ebd3').setRanges([notas]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(3).setBackground('#f6dedc').setRanges([notas]).build()
    ]);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
