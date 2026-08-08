/**
 * Salmones Austral — Backend UNIFICADO de "Resultados por Área" (Filete).
 *
 * Un solo proyecto, un solo dueño (Diego): escribe los datos de Línea1/Línea2
 * en el Sheet propio, sincroniza la Slide propia EN EL MISMO PROCESO (sin
 * llamada HTTP a otro script), y registra cada guardado en la pestaña "Log"
 * del mismo Sheet.
 *
 * Esquema por línea (fila 2 = L1, fila 3 = L2), columnas B..G:
 *   B=Piezas, C=Rango HR, D=Trim, E=Calibre, F=Cliente, G=Acumulado
 * Fila 5 = Fecha (columna B, fórmula =HOY() propia del Sheet, no se toca).
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

// Orden de columnas B..G para cada línea.
var FIELDS = ["piezas", "rangoHr", "trim", "calibre", "cliente", "acumulado"];
var FIELD_HEADERS = ["Piezas", "Rango HR", "Trim", "Calibre", "Cliente", "Acumulado"];

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

function rowToLinea_(row) {
  var out = {};
  for (var i = 0; i < FIELDS.length; i++) {
    var val = row[i];
    out[FIELDS[i]] = FIELDS[i] === "piezas" || FIELDS[i] === "acumulado" ? (Number(val) || 0) : String(val || "");
  }
  return out;
}

function readData_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
  var rows = sheet.getRange("B2:G3").getValues(); // fila2=L1, fila3=L2
  var fechaCell = sheet.getRange("B5").getValue();
  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var result = {
    l1: rowToLinea_(rows[0]),
    l2: rowToLinea_(rows[1]),
    fecha: formatFecha_(fechaCell, tz)
  };
  Logger.log("readData_: " + JSON.stringify(result));
  return result;
}

// ---- Sync Sheet -> Slide (fusionado del script de Pablo, mismo proceso) ----
// Genérico: cualquier celda con contenido se copia al shape de la Slide
// tageado R{fila}C{columna} — agregar columnas nuevas no requiere tocar esto,
// solo hace falta que exista el shape tageado correspondiente en la Slide.
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
        if (!shapes) continue; // todavía no existe ese shape en la Slide -- se ignora sin error
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

function lineaToRow_(params, prefix) {
  return FIELDS.map(function (field) {
    var raw = params[prefix + field];
    if (field === "piezas" || field === "acumulado") return Number(raw) || 0;
    return raw || "";
  });
}

function writeData_(params) {
  var ss = SpreadsheetApp.openById(SHEET_ID); // 1 sola apertura, reutilizada abajo
  var sheet = ss.getSheetByName(SHEET_TAB);

  // Fuerza texto plano en las columnas de texto (Rango HR, Trim, Calibre,
  // Cliente) para que Sheets no las autoconvierta a fecha/número cuando el
  // valor "parece" una fecha (ej. Rango HR "10-11") — mismo bug que ya
  // arreglamos para la celda Fecha.
  sheet.getRange("C2:F3").setNumberFormat("@");

  var row1 = lineaToRow_(params, "l1_");
  var row2 = lineaToRow_(params, "l2_");
  Logger.log("writeData_: L1=" + JSON.stringify(row1) + " L2=" + JSON.stringify(row2));
  sheet.getRange("B2:G3").setValues([row1, row2]);
  // B5 (Fecha) no se toca: queda como fórmula =HOY() propia del Sheet.

  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var dia = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
  var hora = Utilities.formatDate(new Date(), tz, "HH:mm:ss");

  var logSheet = ss.getSheetByName(LOG_TAB);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_TAB);
    logSheet.appendRow(["Día", "Hora", "Área"].concat(
      FIELD_HEADERS.map(function (h) { return "L1 " + h; }),
      FIELD_HEADERS.map(function (h) { return "L2 " + h; })
    ));
  }
  logSheet.appendRow([dia, hora, AREA_NAME].concat(row1, row2));

  syncSheetToSlide_(sheet);

  var fecha = formatFecha_(sheet.getRange("B5").getValue(), tz);
  var result = { l1: rowToLinea_(row1), l2: rowToLinea_(row2), fecha: fecha };
  Logger.log("writeData_: resultado final -> " + JSON.stringify(result));
  return result;
}

function setupHeaders_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
  sheet.getRange("B1:G1").setValues([FIELD_HEADERS]);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || "read";

  if (action === "setup") {
    if (params.key !== WRITE_KEY) return jsonOutput_({ error: "unauthorized" });
    setupHeaders_();
    return jsonOutput_({ ok: true });
  }

  if (action === "update") {
    if (params.key !== WRITE_KEY) {
      return jsonOutput_({ error: "unauthorized" });
    }
    var data = writeData_(params);
    data.ok = true;
    return jsonOutput_(data);
  }

  return jsonOutput_(readData_());
}
