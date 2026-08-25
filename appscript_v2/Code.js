/**
 * Salmones Austral — Backend UNIFICADO de "Resultados por Área" (multi-área).
 *
 * Un solo proyecto: escribe los datos de cada área en su propia pestaña del
 * Sheet, sincroniza las Slides correspondientes EN EL MISMO PROCESO (sin
 * llamada HTTP a otro script), y registra cada guardado en la pestaña de Log
 * propia de esa área — todo dentro del mismo Sheet (SHEET_ID).
 *
 * Agregar un área nueva = agregar una entrada a AREAS. No hace falta tocar
 * el resto del código: readData_/writeData_/setupHeaders_ son genéricos.
 *
 * Publicar como Web App (Implementar > Nueva implementación > Aplicación web):
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario
 */

var WRITE_KEY = "sa-resultados-2026"; // clave de escritura. Debe coincidir con index.html.

var SHEET_ID = "1hkzQJJnTtBi_3k9LhlieLu8msHnoorS3Ikf5RsI8vgo";

var AREAS = {
  filete: {
    label: "Filete",
    sheetTab: "Filete",
    logTab: "Log",
    hasLineas: true, // 2 filas de datos (fila2=L1, fila3=L2)
    fields: ["piezas", "rangoHr", "trim", "calibre", "cliente", "acumulado", "supervisor"],
    fieldHeaders: ["Piezas", "Rango HR", "Trim", "Calibre", "Cliente", "Acumulado", "Supervisor"],
    numericFields: ["piezas", "acumulado"],
    // Transición: se sincroniza a las dos Slides (la vieja y la nueva que la
    // va a reemplazar). Cuando Pablo confirme, sacar el ID viejo de esta lista.
    slideIds: [
      "1SUnpb0vz5XmA5QDpOX2D5dU3UdPKa9CoE9KqP1ez8wo", // Turno Filete (actual, en uso por OnSign)
      "1RtT-RlhLnr26Y8HPAGPkY5QmjRQDw_yaLbe0xVcD06I"  // Turno Filete2 (reemplazo futuro)
    ]
  },
  porcionado: {
    label: "Porcionado",
    sheetTab: "Porcionado",
    logTab: "Log Porcionado",
    hasLineas: false, // 1 sola fila de datos, sin L1/L2
    fields: ["kilos", "kilosHr"],
    fieldHeaders: ["Kilos", "Kilos Hr"],
    numericFields: ["kilos", "kilosHr"],
    slideIds: [
      "1vTQC-6971ROuSl_qY-EXz_d2ytccsUf0qs_0qVIyeDs" // Turno Porcionado
    ]
  }
};

// Convierte un número de columna 1-based a su letra (2 -> "B", 8 -> "H").
function colLetter_(n) {
  var s = "";
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

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

function rowToObj_(row, area) {
  var out = {};
  area.fields.forEach(function (field, i) {
    var val = row[i];
    var isNumeric = area.numericFields.indexOf(field) !== -1;
    out[field] = isNumeric ? (Number(val) || 0) : String(val || "");
  });
  return out;
}

function objToRow_(params, area, prefix) {
  return area.fields.map(function (field) {
    var raw = params[(prefix || "") + field];
    var isNumeric = area.numericFields.indexOf(field) !== -1;
    return isNumeric ? (Number(raw) || 0) : (raw || "");
  });
}

function fechaRow_(area) {
  return area.hasLineas ? 5 : 4; // 1 fila menos si no hay L2
}

function getAreaSheet_(ss, area) {
  return ss.getSheetByName(area.sheetTab) || ss.insertSheet(area.sheetTab);
}

function readData_(area) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getAreaSheet_(ss, area);
  var lastRow = area.hasLineas ? 3 : 2;
  var lastCol = colLetter_(area.fields.length + 1);
  var rows = sheet.getRange("B2:" + lastCol + lastRow).getValues();
  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var fecha = formatFecha_(sheet.getRange("B" + fechaRow_(area)).getValue(), tz);

  var result = area.hasLineas
    ? { l1: rowToObj_(rows[0], area), l2: rowToObj_(rows[1], area), fecha: fecha }
    : Object.assign(rowToObj_(rows[0], area), { fecha: fecha });
  Logger.log("readData_[" + area.label + "]: " + JSON.stringify(result));
  return result;
}

// ---- Sync Sheet -> Slide (fusionado del script de Pablo, mismo proceso) ----
// Genérico: cada celda (con o sin contenido) se copia al shape de la Slide
// tageado R{fila}C{columna} — si la celda está vacía, el shape también queda
// vacío. Agregar columnas nuevas no requiere tocar esto, solo hace falta que
// exista el shape tageado correspondiente en la Slide.
function syncSheetToSlide_(sheet, slideIds) {
  if (!slideIds || !slideIds.length) return;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    Logger.log("No se obtuvo el lock, se salta el sync: " + e);
    return;
  }
  try {
    var data = sheet.getDataRange().getValues();
    slideIds.forEach(function (slideId) {
      syncSheetToOneSlide_(data, slideId);
    });
  } finally {
    lock.releaseLock();
  }
}

function syncSheetToOneSlide_(data, slideId) {
  var slide = SlidesApp.openById(slideId).getSlides()[0];
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
      var tag = "R" + (r + 1) + "C" + (c + 1);
      var shapes = shapeByTag[tag];
      if (!shapes) continue; // todavía no existe ese shape en la Slide -- se ignora sin error
      // Celda vacía -> el shape también se vacía (no se queda con el dato viejo).
      var formatted = (rawValue === "" || rawValue === null || rawValue === undefined) ? "" : formatValue_(rawValue, now);
      shapes.forEach(function (shape) {
        if (shape.getText().asString() !== formatted) {
          shape.getText().setText(formatted);
          Logger.log("syncSheetToSlide_: [" + slideId + "] " + tag + " -> '" + formatted + "'");
        }
      });
    }
  }
}

function formatValue_(rawValue, now) {
  if (rawValue instanceof Date) {
    return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return typeof rawValue === "number" ? rawValue.toLocaleString("es-CL") : String(rawValue);
}

function writeData_(params, area) {
  var ss = SpreadsheetApp.openById(SHEET_ID); // 1 sola apertura, reutilizada abajo
  var sheet = getAreaSheet_(ss, area);
  var lastRow = area.hasLineas ? 3 : 2;

  // Fuerza texto plano en las columnas no numéricas para que Sheets no las
  // autoconvierta a fecha/número cuando el valor "parece" una fecha (ej.
  // Rango HR "10-11") — mismo bug que ya arreglamos para la celda Fecha.
  area.fields.forEach(function (field, i) {
    if (area.numericFields.indexOf(field) !== -1) return;
    var col = colLetter_(i + 2);
    sheet.getRange(col + "2:" + col + lastRow).setNumberFormat("@");
  });

  var rows = [objToRow_(params, area, area.hasLineas ? "l1_" : "")];
  if (area.hasLineas) rows.push(objToRow_(params, area, "l2_"));

  var lastCol = colLetter_(area.fields.length + 1);
  Logger.log("writeData_[" + area.label + "]: " + JSON.stringify(rows));
  sheet.getRange("B2:" + lastCol + lastRow).setValues(rows);
  // La celda de Fecha no se toca: queda como fórmula =TODAY() propia del Sheet.

  var tz = Session.getScriptTimeZone() || "America/Santiago";
  var dia = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
  var hora = Utilities.formatDate(new Date(), tz, "HH:mm:ss");

  var logSheet = ss.getSheetByName(area.logTab);
  if (!logSheet) {
    logSheet = ss.insertSheet(area.logTab);
    logSheet.appendRow(logHeaderRow_(area));
  }
  var logRow = [dia, hora, area.label].concat(rows[0]);
  if (area.hasLineas) logRow = logRow.concat(rows[1]);
  logSheet.appendRow(logRow);

  syncSheetToSlide_(sheet, area.slideIds);

  var fecha = formatFecha_(sheet.getRange("B" + fechaRow_(area)).getValue(), tz);
  var result = area.hasLineas
    ? { l1: rowToObj_(rows[0], area), l2: rowToObj_(rows[1], area), fecha: fecha }
    : Object.assign(rowToObj_(rows[0], area), { fecha: fecha });
  Logger.log("writeData_[" + area.label + "]: resultado final -> " + JSON.stringify(result));
  return result;
}

function logHeaderRow_(area) {
  if (!area.hasLineas) {
    return ["Día", "Hora", "Área"].concat(area.fieldHeaders);
  }
  return ["Día", "Hora", "Área"].concat(
    area.fieldHeaders.map(function (h) { return "L1 " + h; }),
    area.fieldHeaders.map(function (h) { return "L2 " + h; })
  );
}

function setupHeaders_(area) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getAreaSheet_(ss, area);
  var lastCol = colLetter_(area.fieldHeaders.length + 1);
  sheet.getRange("B1:" + lastCol + "1").setValues([area.fieldHeaders]);

  // Asegura la celda de Fecha (label + fórmula =HOY()) — no pisa nada si ya
  // existía con el mismo valor.
  var fr = fechaRow_(area);
  sheet.getRange("A" + fr).setValue("Fecha");
  sheet.getRange("B" + fr).setFormula("=TODAY()"); // Apps Script exige el nombre en inglés aunque el Sheet esté en español

  var logSheet = ss.getSheetByName(area.logTab);
  if (!logSheet) {
    logSheet = ss.insertSheet(area.logTab);
  }
  var header = logHeaderRow_(area);
  logSheet.getRange(1, 1, 1, header.length).setValues([header]);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || "read";

  var area = AREAS[params.area || "filete"];
  if (!area) return jsonOutput_({ error: "unknown_area" });

  if (action === "setup") {
    if (params.key !== WRITE_KEY) return jsonOutput_({ error: "unauthorized" });
    setupHeaders_(area);
    return jsonOutput_({ ok: true });
  }

  if (action === "update") {
    if (params.key !== WRITE_KEY) {
      return jsonOutput_({ error: "unauthorized" });
    }
    var data = writeData_(params, area);
    data.ok = true;
    return jsonOutput_(data);
  }

  return jsonOutput_(readData_(area));
}
