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

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  return { linea1: linea1, linea2: linea2, fecha: fecha };
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || "read";

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
