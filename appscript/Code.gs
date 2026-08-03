/**
 * Salmones Austral — Backend de "Resultados por Área" (Filete).
 *
 * Arquitectura: esta página escribe Línea1/Línea2 directamente en el Google
 * Sheet COMPARTIDO que lee el script de sincronización a Slides (proyecto de
 * Pablo). Ese sync a Slides lo dispara el navegador (index.html) justo
 * después de que este endpoint confirme el guardado — no se hace desde aquí
 * con UrlFetchApp porque esta cuenta de Workspace no tiene autorizado el
 * scope script.external_request (bloqueo, aparentemente, a nivel de admin).
 * OnSign ya no consulta ningún JSON nuestro — reproduce la Slide directamente.
 *
 * Publicar como Web App (Implementar > Nueva implementación > Aplicación web):
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario
 * La URL /exec que entrega el deploy es la que usa la constante
 * APPS_SCRIPT_URL en index.html (no cambia entre redeploys si se usa
 * `clasp deploy --deploymentId <mismo-id>`).
 */

var WRITE_KEY = "sa-resultados-2026"; // clave de escritura. Cámbiala aquí y en index.html si quieres rotarla.

// Sheet compartido que lee el script de sync-a-Slides (no es propiedad nuestra).
var SHARED_SHEET_ID = "1sF8TLMUcYGfOupQfTiE5n63TPqyuXoc7Z9ljiu3z2LA";
var SHARED_SHEET_TAB = "Hoja 1";

var LOG_SHEET_NAME = "Log";

function getSharedSheet_() {
  return SpreadsheetApp.openById(SHARED_SHEET_ID).getSheetByName(SHARED_SHEET_TAB);
}

// Spreadsheet PROPIA (privada), solo para el log de auditoría — separada del
// Sheet compartido de arriba. Se crea sola la primera vez que corre.
function getAuditSpreadsheet_() {
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

function getLogSheet_() {
  var ss = getAuditSpreadsheet_();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(["Fecha y hora", "Linea1", "Linea2", "Fecha turno"]);
  }
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("D:D").setNumberFormat("@");
  return sheet;
}

function readData_() {
  var sheet = getSharedSheet_();
  var linea1 = sheet.getRange("B2").getValue();
  var linea2 = sheet.getRange("B3").getValue();
  var fechaCell = sheet.getRange("B5").getValue();
  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var fecha;
  if (Object.prototype.toString.call(fechaCell) === "[object Date]") {
    // Fórmula =HOY() (u otro valor tipo Fecha real): se formatea directo.
    fecha = Utilities.formatDate(fechaCell, tz, "dd/MM/yyyy");
  } else if (typeof fechaCell === "number") {
    // Serial de fecha de Sheets (días desde 1899-12-30, epoch de Sheets/Excel)
    // — pasa esto cuando la celda tiene formato Texto pero el valor interno
    // sigue siendo una fecha real. Se reconstruye el Date antes de formatear.
    var asDate = new Date(Math.round((fechaCell - 25569) * 86400 * 1000));
    fecha = Utilities.formatDate(asDate, tz, "dd/MM/yyyy");
  } else {
    fecha = String(fechaCell || "");
  }
  return {
    linea1: Number(linea1) || 0,
    linea2: Number(linea2) || 0,
    fecha: fecha
  };
}

function writeData_(linea1, linea2) {
  var sheet = getSharedSheet_();
  sheet.getRange("B2").setValue(linea1);
  sheet.getRange("B3").setValue(linea2);
  // B5 (Fecha) no se toca: queda como fórmula =HOY() en el Sheet compartido.

  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var timestamp = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
  var fecha = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
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
    return jsonOutput_({ spreadsheetUrl: getAuditSpreadsheet_().getUrl() });
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
