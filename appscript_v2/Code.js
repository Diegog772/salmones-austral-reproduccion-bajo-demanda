/**
 * Salmones Austral — Backend UNIFICADO de "Resultados por Área" (Filete).
 *
 * Un solo proyecto, un solo dueño (Diego): escribe Línea1/Línea2 en el Sheet
 * propio, sincroniza la Slide propia EN EL MISMO PROCESO (sin llamada HTTP a
 * otro script), y registra cada guardado en la pestaña "Log" del mismo Sheet.
 * Reemplaza el split anterior (PSP_Filete + script de sync de Pablo) una vez
 * que esto quede verificado — NO se conecta todavía a index.html.
 *
 * Publicar como Web App (Implementar > Nueva implementación > Aplicación web):
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario
 */

var WRITE_KEY = "sa-resultados-2026"; // clave de escritura. Debe coincidir con index.html cuando se conecte.

var SHEET_ID = "1hkzQJJnTtBi_3k9LhlieLu8msHnoorS3Ikf5RsI8vgo";
var SHEET_TAB = "Hoja 1";
var LOG_TAB = "Log";
var SLIDE_ID = "1SUnpb0vz5XmA5QDpOX2D5dU3UdPKa9CoE9KqP1ez8wo";
var AREA_NAME = "Filete";

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
}

function getLogSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(LOG_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_TAB);
    sheet.appendRow(["Día", "Hora", "Área", "Filete L1 (Piezas)", "Filete L2 (Piezas)"]);
  }
  return sheet;
}

function readData_() {
  var sheet = getSheet_();
  var linea1 = sheet.getRange("B2").getValue();
  var linea2 = sheet.getRange("B3").getValue();
  var fechaCell = sheet.getRange("B5").getValue();
  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var fecha;
  if (Object.prototype.toString.call(fechaCell) === "[object Date]") {
    fecha = Utilities.formatDate(fechaCell, tz, "dd/MM/yyyy");
  } else if (typeof fechaCell === "number") {
    // Serial de fecha de Sheets (días desde 1899-12-30) — pasa si la celda
    // quedó con formato Texto pero el valor interno sigue siendo una fecha.
    var asDate = new Date(Math.round((fechaCell - 25569) * 86400 * 1000));
    fecha = Utilities.formatDate(asDate, tz, "dd/MM/yyyy");
  } else {
    fecha = String(fechaCell || "");
  }
  Logger.log("readData_: linea1=" + linea1 + " linea2=" + linea2 + " fecha=" + fecha);
  return {
    linea1: Number(linea1) || 0,
    linea2: Number(linea2) || 0,
    fecha: fecha
  };
}

// ---- Sync Sheet -> Slide (fusionado del script de Pablo, mismo proceso) ----
function syncSheetToSlide_() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    Logger.log("No se obtuvo el lock, se salta el sync: " + e);
    return;
  }
  try {
    var sheet = getSheet_();
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
  var sheet = getSheet_();
  sheet.getRange("B2").setValue(linea1);
  sheet.getRange("B3").setValue(linea2);
  // B5 (Fecha) no se toca: queda como fórmula =HOY() propia del Sheet.

  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var dia = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
  var hora = Utilities.formatDate(new Date(), tz, "HH:mm:ss");
  getLogSheet_().appendRow([dia, hora, AREA_NAME, linea1, linea2]);
  Logger.log("writeData_: fila agregada al Log -> " + dia + " " + hora + " " + AREA_NAME + " " + linea1 + " " + linea2);

  syncSheetToSlide_();

  var result = { linea1: linea1, linea2: linea2, fecha: readData_().fecha };
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
