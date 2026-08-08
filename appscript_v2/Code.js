/**
 * Salmones Austral — Backend UNIFICADO de "Resultados por Área" (Filete).
 *
 * Un solo proyecto, un solo dueño (Diego): escribe Línea1/Línea2 en el Sheet
 * propio, sincroniza la Slide propia EN EL MISMO PROCESO (sin llamada HTTP a
 * otro script), y registra cada guardado en la pestaña "Log" del mismo Sheet.
 *
 * Optimización de velocidad: cada SpreadsheetApp.openById / SlidesApp.openById
 * es un viaje de red aparte. Se abre el Sheet UNA sola vez por request y esa
 * referencia se reutiliza en todo el flujo (antes se abría 3 veces).
 *
 * Publicar como Web App (Implementar > Nueva implementación > Aplicación web):
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario
 */

var WRITE_KEY = "sa-resultados-2026"; // clave de escritura. Debe coincidir con index.html.

var SHEET_ID = "1hkzQJJnTtBi_3k9LhlieLu8msHnoorS3Ikf5RsI8vgo";
var SHEET_TAB = "Hoja 1";
var LOG_TAB = "Log";
var SLIDE_ID = "1SUnpb0vz5XmA5QDpOX2D5dU3UdPKa9CoE9KqP1ez8wo";
var AREA_NAME = "Filete";

function formatFecha_(fechaCell, tz) {
  if (Object.prototype.toString.call(fechaCell) === "[object Date]") {
    return Utilities.formatDate(fechaCell, tz, "dd/MM/yyyy");
  }
  if (typeof fechaCell === "number") {
    // Serial de fecha de Sheets (días desde 1899-12-30) — pasa si la celda
    // quedó con formato Texto pero el valor interno sigue siendo una fecha.
    var asDate = new Date(Math.round((fechaCell - 25569) * 86400 * 1000));
    return Utilities.formatDate(asDate, tz, "dd/MM/yyyy");
  }
  return String(fechaCell || "");
}

function readData_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
  var vals = sheet.getRange("B2:B5").getValues(); // 1 sola llamada en vez de 3
  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var result = {
    linea1: Number(vals[0][0]) || 0,
    linea2: Number(vals[1][0]) || 0,
    fecha: formatFecha_(vals[3][0], tz)
  };
  Logger.log("readData_: " + JSON.stringify(result));
  return result;
}

// ---- Sync Sheet -> Slide (fusionado del script de Pablo, mismo proceso) ----
// Recibe el `sheet` ya abierto para no reabrir el Spreadsheet de nuevo.
function syncSheetToSlide_(sheet) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    Logger.log("No se obtuvo el lock, se salta el sync: " + e);
    return;
  }
  try {
    var data = sheet.getDataRange().getValues();
    var slide = SlidesApp.openById(SLIDE_ID).getSlides()[0];
    var shapeByTag = {};
    slide.getShapes().forEach(function (shape) {
      var tag = (shape.getTitle() || "").trim();
      if (!tag) return;
      if (!shapeByTag[tag]) shapeByTag[tag] = [];
      shapeByTag[tag].push(shape);
    });

    var now = new Date();
    for (var r = 0; r < data.length; r++) {
      for (var c = 0; c < data[r].length; c++) {
        var rawValue = data[r][c];
        if (rawValue === "" || rawValue === null || rawValue === undefined) continue;
        var tag = "R" + (r + 1) + "C" + (c + 1);
        var shapes = shapeByTag[tag];
        if (!shapes) continue;
        var formatted = formatValue_(rawValue, now);
        shapes.forEach(function (shape) {
          if (shape.getText().asString() !== formatted) {
            shape.getText().setText(formatted);
            Logger.log("syncSheetToSlide_: " + tag + " -> " + formatted);
          }
        });
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function formatValue_(rawValue, now) {
  if (rawValue instanceof Date) {
    return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return typeof rawValue === "number" ? rawValue.toLocaleString("es-CL") : String(rawValue);
}

function writeData_(linea1, linea2) {
  Logger.log("writeData_: guardando linea1=" + linea1 + " linea2=" + linea2);
  var ss = SpreadsheetApp.openById(SHEET_ID); // 1 sola apertura, reutilizada abajo
  var sheet = ss.getSheetByName(SHEET_TAB);

  sheet.getRange("B2:B3").setValues([[linea1], [linea2]]); // 1 sola escritura en vez de 2
  // B5 (Fecha) no se toca: queda como fórmula =HOY() propia del Sheet.

  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var dia = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
  var hora = Utilities.formatDate(new Date(), tz, "HH:mm:ss");

  var logSheet = ss.getSheetByName(LOG_TAB);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_TAB);
    logSheet.appendRow(["Día", "Hora", "Área", "Filete L1 (Piezas)", "Filete L2 (Piezas)"]);
  }
  logSheet.appendRow([dia, hora, AREA_NAME, linea1, linea2]);

  syncSheetToSlide_(sheet);

  var fecha = formatFecha_(sheet.getRange("B5").getValue(), tz);
  var result = { linea1: linea1, linea2: linea2, fecha: fecha };
  Logger.log("writeData_: resultado final -> " + JSON.stringify(result));
  return result;
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
