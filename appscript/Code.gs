/**
 * Salmones Austral — Backend de "Resultados por Área" (Filete).
 *
 * Publicar como Web App (Implementar > Nueva implementación > Aplicación web):
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario
 * La URL /exec que entrega el deploy es la que se configura en:
 *   - OnSign TV (fuente de datos "JSON desde URL"), y
 *   - la constante APPS_SCRIPT_URL en index.html.
 */

var WRITE_KEY = "sa-resultados-2026"; // clave de escritura. Cámbiala aquí y en index.html si quieres rotarla.
var SHEET_NAME = "Filete";
var LOG_SHEET_NAME = "Log";

// Proyecto standalone (no atado a una Sheet): se crea una Spreadsheet propia la
// primera vez que corre, y su ID queda cacheado en las Script Properties.
function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("SPREADSHEET_ID");
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      // el ID guardado ya no es válido (ej. se borró la hoja) — se recrea abajo.
    }
  }
  var ss = SpreadsheetApp.create("Salmones Austral - Resultados por Área");
  props.setProperty("SPREADSHEET_ID", ss.getId());
  return ss;
}

function getSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Linea", "Piezas"]);
    sheet.appendRow(["L1", 0]);
    sheet.appendRow(["L2", 0]);
    sheet.appendRow(["Fecha", ""]);
  }
  return sheet;
}

function getLogSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(["Fecha y hora", "Linea1", "Linea2", "Fecha turno"]);
  }
  return sheet;
}

function readData_() {
  var sheet = getSheet_();
  var linea1 = sheet.getRange(2, 2).getValue();
  var linea2 = sheet.getRange(3, 2).getValue();
  var fecha = sheet.getRange(4, 2).getValue();
  return {
    linea1: Number(linea1) || 0,
    linea2: Number(linea2) || 0,
    fecha: String(fecha || "")
  };
}

function writeData_(linea1, linea2) {
  var sheet = getSheet_();
  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var fecha = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
  sheet.getRange(2, 2).setValue(linea1);
  sheet.getRange(3, 2).setValue(linea2);
  sheet.getRange(4, 2).setValue(fecha);

  var timestamp = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
  getLogSheet_().appendRow([timestamp, linea1, linea2, fecha]);

  return { linea1: linea1, linea2: linea2, fecha: fecha };
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || "read";

  if (action === "meta") {
    return jsonOutput_({ spreadsheetUrl: getSpreadsheet_().getUrl() });
  }

  if (action === "update") {
    if (params.key !== WRITE_KEY) {
      return jsonOutput_({ error: "unauthorized" });
    }
    var linea1 = Number(params.linea1);
    var linea2 = Number(params.linea2);
    if (isNaN(linea1) || isNaN(linea2)) {
      return jsonOutput_({ error: "invalid_numbers" });
    }
    var data = writeData_(linea1, linea2);
    data.ok = true;
    return jsonOutput_(data);
  }

  return jsonOutput_(readData_());
}
