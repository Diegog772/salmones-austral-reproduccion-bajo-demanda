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

// Clave separada para "Equipo Control Producción" (Resultados por Día/Área).
// Independiente de WRITE_KEY — acceso distinto del de los operarios de planta.
var CONTROL_KEY = "sa-control-2026"; // debe coincidir con CONTROL_WRITE_KEY en index.html.

var SHEET_ID = "1hkzQJJnTtBi_3k9LhlieLu8msHnoorS3Ikf5RsI8vgo";

// "Resumen por Día/Área": resumen semanal por supervisor, independiente del
// sistema AREAS de arriba (ese es "estado actual del turno"; esto es
// "historial semanal", una fila por supervisor por semana, sin pestaña de
// Log aparte — la pestaña entera ES el historial).
var SUMMARY_AREAS = {
  filete: {
    label: "Filete",
    sheetTab: "Resumen Filete",
    slideIds: ["1D2InuuGHTql1v7ZFhswvfIGCW6ioQrlJZz0uRNw5wtw"]
  },
  lavado: { label: "Lavado", sheetTab: "Resumen Lavado", slideIds: [] }
};
var SUMMARY_MAX_SLIDE_ROWS = 6; // cantidad de filas de supervisor que soporta la Slide
var SUMMARY_HEADERS = ["Semana", "Supervisor", "Día 1", "Día 2", "Día 3", "Día 4", "Día 5", "Día 6", "Promedio", "Total"];

var AREAS = {
  filete: {
    label: "Filete",
    sheetTab: "Filete",
    logTab: "Log",
    hasLineas: true, // 2 filas de datos (fila2=L1, fila3=L2)
    fields: ["piezas", "rangoHr", "trim", "calibre", "cliente", "acumulado", "supervisor"],
    fieldHeaders: ["Piezas", "Rango HR", "Trim", "Calibre", "Cliente", "Acumulado", "Supervisor"],
    numericFields: ["piezas", "acumulado"],
    // Transición: se sincroniza a las tres Slides (la vieja, la nueva que la
    // va a reemplazar, y Empaque que solo muestra un subset de los datos de
    // Filete). Cuando Pablo confirme el corte de Filete2, sacar el ID viejo.
    slideIds: [
      "1SUnpb0vz5XmA5QDpOX2D5dU3UdPKa9CoE9KqP1ez8wo", // Turno Filete (actual, en uso por OnSign)
      "1RtT-RlhLnr26Y8HPAGPkY5QmjRQDw_yaLbe0xVcD06I", // Turno Filete2 (reemplazo futuro)
      "16km6hEAitpp2h8-OIvmahmeuWng1e8NNy242TSlfVCs"  // Turno Empaque (espeja Piezas Hr/Rango HR/Trim/Calibre/Cliente)
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
  },
  lavado: {
    label: "Lavado",
    sheetTab: "Lavado",
    logTab: "Log Lavado",
    hasLineas: false, // 1 sola fila de datos, sin L1/L2
    fields: ["piezas", "acumulado", "supervisor"],
    fieldHeaders: ["Piezas Hr", "Acumulado", "Supervisor"],
    numericFields: ["piezas", "acumulado"],
    slideIds: [
      "1K9mt7U_EF2WacSg-vJA_G1lkzw1MJ08K7CQTR8nlnCU" // Turno Lavado
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

function formatHora_(horaCell, tz) {
  if (Object.prototype.toString.call(horaCell) === "[object Date]") {
    return Utilities.formatDate(horaCell, tz, "HH:mm:ss");
  }
  return String(horaCell || "");
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
  var fr = fechaRow_(area);
  var fecha = formatFecha_(sheet.getRange("B" + fr).getValue(), tz);
  var hora = formatHora_(sheet.getRange("D" + fr).getValue(), tz);

  var result = area.hasLineas
    ? { l1: rowToObj_(rows[0], area), l2: rowToObj_(rows[1], area), fecha: fecha, hora: hora }
    : Object.assign(rowToObj_(rows[0], area), { fecha: fecha, hora: hora });
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

  // Hora del guardado (celda D de la fila de Fecha) — se pisa cada vez, a
  // diferencia de la Fecha (fórmula =TODAY() propia del Sheet, no se toca).
  // Forzar texto plano para que Sheets no la autoconvierta a un valor Hora.
  var horaCell = sheet.getRange("D" + fechaRow_(area));
  horaCell.setNumberFormat("@");
  horaCell.setValue(hora);

  syncSheetToSlide_(sheet, area.slideIds);

  var fecha = formatFecha_(sheet.getRange("B" + fechaRow_(area)).getValue(), tz);
  var result = area.hasLineas
    ? { l1: rowToObj_(rows[0], area), l2: rowToObj_(rows[1], area), fecha: fecha, hora: hora }
    : Object.assign(rowToObj_(rows[0], area), { fecha: fecha, hora: hora });
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
  sheet.getRange("C" + fr).setValue("Hora"); // valor plano, se pisa en cada guardado (ver writeData_)

  var logSheet = ss.getSheetByName(area.logTab);
  if (!logSheet) {
    logSheet = ss.insertSheet(area.logTab);
  }
  var header = logHeaderRow_(area);
  logSheet.getRange(1, 1, 1, header.length).setValues([header]);
}

// Construye un layout básico (título + N valores en columnas + fecha) en una
// Slide en blanco, con los shapes ya tageados. Solo se usa una vez para
// arrancar la Slide — el diseño se puede embellecer después en Slides sin
// tocar el Alt Text de estos shapes.
// columns: [{ label: "Kilos Hr", tag: "R2C3" }, ...]
function buildSimpleAreaSlide_(slideId, title, columns, fechaTag) {
  var presentation = SlidesApp.openById(slideId);
  var slide = presentation.getSlides()[0];
  var pageW = presentation.getPageWidth();
  var pageH = presentation.getPageHeight();

  function addBox(text, x, y, w, h, fontSize, bold, tag) {
    var box = slide.insertTextBox(text, x, y, w, h);
    var textRange = box.getText();
    textRange.getTextStyle().setFontSize(fontSize).setBold(!!bold);
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    if (tag) box.setTitle(tag);
    return box;
  }

  addBox(title, pageW * 0.05, pageH * 0.06, pageW * 0.9, pageH * 0.14, 32, true, null);

  var margin = pageW * 0.05;
  var gap = pageW * 0.04;
  var colW = (pageW * 0.9 - gap * (columns.length - 1)) / columns.length;

  columns.forEach(function (col, i) {
    var x = margin + i * (colW + gap);
    addBox(col.label, x, pageH * 0.28, colW, pageH * 0.08, 16, true, null);
    addBox("-", x, pageH * 0.37, colW, pageH * 0.22, 36, true, col.tag);
  });

  addBox("Fecha", pageW * 0.05, pageH * 0.70, pageW * 0.3, pageH * 0.08, 14, true, null);
  addBox("--/--/----", pageW * 0.05, pageH * 0.78, pageW * 0.3, pageH * 0.10, 16, false, fechaTag);
}

// Igual que buildSimpleAreaSlide_ pero con 2 filas (L1/L2) por columna —
// para Slides que reflejan datos de un área "hasLineas" (ej. Empaque
// mostrando los mismos datos que Filete). columns: [{ label, col }] donde
// "col" es el número de columna del Sheet (B=2, C=3, ...).
function buildLineasTableSlide_(slideId, title, columns, fechaTag) {
  var presentation = SlidesApp.openById(slideId);
  var slide = presentation.getSlides()[0];
  var pageW = presentation.getPageWidth();
  var pageH = presentation.getPageHeight();

  function addBox(text, x, y, w, h, fontSize, bold, tag) {
    var box = slide.insertTextBox(text, x, y, w, h);
    var textRange = box.getText();
    textRange.getTextStyle().setFontSize(fontSize).setBold(!!bold);
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    if (tag) box.setTitle(tag);
    return box;
  }

  addBox(title, pageW * 0.03, pageH * 0.05, pageW * 0.94, pageH * 0.10, 26, true, null);

  var startX = pageW * 0.03;
  var labelColW = pageW * 0.08;
  var tableW = pageW * 0.94 - labelColW;
  var colW = tableW / columns.length;

  var headerY = pageH * 0.20;
  var headerH = pageH * 0.07;
  var row1Y = pageH * 0.30;
  var row2Y = pageH * 0.48;
  var rowH = pageH * 0.15;

  columns.forEach(function (f, i) {
    var x = startX + labelColW + i * colW;
    addBox(f.label, x, headerY, colW, headerH, 13, true, null);
    addBox("-", x, row1Y, colW, rowH, 22, true, "R2C" + f.col);
    addBox("-", x, row2Y, colW, rowH, 22, true, "R3C" + f.col);
  });

  addBox("L1", startX, row1Y, labelColW, rowH, 18, true, null);
  addBox("L2", startX, row2Y, labelColW, rowH, 18, true, null);

  if (fechaTag) {
    addBox("Fecha", startX, pageH * 0.70, pageW * 0.25, pageH * 0.07, 13, true, null);
    addBox("--/--/----", startX, pageH * 0.77, pageW * 0.25, pageH * 0.09, 14, false, fechaTag);
  }
}

// Agrega un shape "Hora" (label + valor tageado) al lado del shape de Fecha
// ya existente (tageado fechaTag) en una Slide. Usa la posición del shape de
// Fecha como referencia, así no hace falta conocer el layout de cada Slide.
function addHoraNextToFecha_(slideId, fechaTag, horaTag) {
  var presentation = SlidesApp.openById(slideId);
  var slide = presentation.getSlides()[0];
  var fechaValueShape = null;
  slide.getShapes().forEach(function (s) {
    if ((s.getTitle() || "").trim() === fechaTag) fechaValueShape = s;
  });
  if (!fechaValueShape) throw new Error("No se encontró shape con tag " + fechaTag);

  var gap = presentation.getPageWidth() * 0.02;
  var vw = fechaValueShape.getWidth();
  var vh = fechaValueShape.getHeight();
  var vx = Math.max(0, fechaValueShape.getLeft() - gap - vw);
  var vy = fechaValueShape.getTop();
  var labelH = vh * 0.5;
  var labelY = Math.max(0, vy - labelH - presentation.getPageHeight() * 0.01);

  function addBox(text, x, y, w, h, fontSize, bold, tag) {
    var box = slide.insertTextBox(text, x, y, w, h);
    var tr = box.getText();
    tr.getTextStyle().setFontSize(fontSize).setBold(!!bold);
    tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    if (tag) box.setTitle(tag);
    return box;
  }

  addBox("Hora", vx, labelY, vw, labelH, 13, true, null);
  addBox("--:--:--", vx, vy, vw, vh, 14, false, horaTag);
}

// ---- Resumen por Día/Área (Equipo Control Producción) ----

function getSummarySheet_(ss, summaryArea) {
  return ss.getSheetByName(summaryArea.sheetTab) || ss.insertSheet(summaryArea.sheetTab);
}

function ensureSummaryHeaders_(sheet) {
  var firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell === SUMMARY_HEADERS[0]) return; // ya están puestos
  sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setFontWeight("bold");
  // Semana/Supervisor en texto plano — mismo bug conocido de Sheets autoconvirtiendo texto a fecha/número.
  sheet.getRange("A2:B").setNumberFormat("@");
}

function computeSummaryStats_(dias) {
  var total = 0;
  var nonZeroCount = 0;
  dias.forEach(function (v) {
    var n = Number(v) || 0;
    total += n;
    if (n !== 0) nonZeroCount++;
  });
  var promedio = nonZeroCount ? Math.round(total / nonZeroCount) : 0;
  return { total: total, promedio: promedio };
}

function readSummary_(summaryArea, week) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getSummarySheet_(ss, summaryArea);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [] };
  var data = sheet.getRange(2, 1, lastRow - 1, SUMMARY_HEADERS.length).getValues();
  var rows = [];
  data.forEach(function (row) {
    if (String(row[0]) !== week) return;
    rows.push({
      supervisor: String(row[1] || ""),
      dias: row.slice(2, 8).map(function (v) { return Number(v) || 0; }),
      promedio: Number(row[8]) || 0,
      total: Number(row[9]) || 0
    });
  });
  return { rows: rows };
}

// Borra las filas existentes de esa (área, semana) y reinserta las nuevas —
// Promedio/Total se recalculan acá, no se confía en lo que mande el navegador.
function writeSummary_(summaryArea, week, incomingRows) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getSummarySheet_(ss, summaryArea);
  ensureSummaryHeaders_(sheet);

  var lastRow = sheet.getLastRow();
  var existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, SUMMARY_HEADERS.length).getValues() : [];
  var keep = existing.filter(function (row) { return String(row[0]) !== week; });

  var newRows = incomingRows.map(function (r) {
    var dias = (r.dias || []).map(function (v) { return Number(v) || 0; });
    while (dias.length < 6) dias.push(0);
    var stats = computeSummaryStats_(dias);
    return [week, String(r.supervisor || "")].concat(dias, [stats.promedio, stats.total]);
  });

  var finalRows = keep.concat(newRows);
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, SUMMARY_HEADERS.length).clearContent();
  }
  if (finalRows.length) {
    sheet.getRange(2, 1, finalRows.length, SUMMARY_HEADERS.length).setValues(finalRows);
  }

  var resultRows = newRows.map(function (row) {
    return { supervisor: row[1], dias: row.slice(2, 8), promedio: row[8], total: row[9] };
  });

  syncSummaryToSlide_(summaryArea, week, resultRows);

  return { rows: resultRows };
}

// ---- Sync Resumen -> Slide ----
// A diferencia de syncSheetToSlide_ (que refleja el Sheet completo), acá se
// arma una "grilla virtual" con las filas que se acaban de guardar para esa
// semana — la pestaña Resumen tiene el historial de TODAS las semanas
// apiladas, así que no tiene sentido sincronizar por posición real de fila.
// R1C1 = título (semana). R1C2..R1C7 = fecha real de cada columna de día
// (encabezados). Fila de supervisor N -> R{N+2}C{1..9} (1=Supervisor,
// 2-7=Día 1-6, 8=Promedio, 9=Total).
function syncSummaryToSlide_(summaryArea, week, rows) {
  if (!summaryArea.slideIds || !summaryArea.slideIds.length) return;

  var grid = {};
  grid["R1C1"] = formatWeekLabel_(week);
  var dayLabels = summaryWeekDayLabels_(week);
  for (var d0 = 0; d0 < 6; d0++) {
    grid["R1C" + (d0 + 2)] = dayLabels[d0];
  }
  for (var i = 0; i < SUMMARY_MAX_SLIDE_ROWS; i++) {
    var r = i + 2;
    var row = rows[i];
    grid["R" + r + "C1"] = row ? row.supervisor : "";
    for (var d = 0; d < 6; d++) {
      grid["R" + r + "C" + (d + 2)] = row ? row.dias[d] : "";
    }
    grid["R" + r + "C8"] = row ? row.promedio : "";
    grid["R" + r + "C9"] = row ? row.total : "";
  }

  summaryArea.slideIds.forEach(function (slideId) {
    syncGridToSlide_(grid, slideId);
  });
}

function formatWeekLabel_(week) {
  var parts = String(week).split("-W");
  return parts.length === 2 ? "SEMANA " + Number(parts[1]) : String(week);
}

var MESES_ES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];

// Mismo cálculo ISO-8601 semana->lunes que usa index.html (isoWeekToMonday),
// para que los encabezados de día de la Slide muestren la fecha real.
function isoWeekToMonday_(weekValue) {
  var parts = String(weekValue).split("-W");
  var year = Number(parts[0]);
  var week = Number(parts[1]);
  var jan4 = new Date(Date.UTC(year, 0, 4));
  var jan4Day = jan4.getUTCDay() || 7;
  var monday = new Date(jan4);
  monday.setUTCDate(jan4.getUTCDate() - jan4Day + 1 + (week - 1) * 7);
  return monday;
}

function summaryWeekDayLabels_(week) {
  if (!/^\d{4}-W\d{1,2}$/.test(String(week))) return ["", "", "", "", "", ""];
  var monday = isoWeekToMonday_(week);
  var labels = [];
  for (var i = 0; i < 6; i++) {
    var d = new Date(monday);
    d.setUTCDate(monday.getUTCDate() + i);
    labels.push(("0" + d.getUTCDate()).slice(-2) + "-" + MESES_ES_CORTOS[d.getUTCMonth()]);
  }
  return labels;
}

function syncGridToSlide_(grid, slideId) {
  var slide = SlidesApp.openById(slideId).getSlides()[0];
  var shapeByTag = {};
  slide.getShapes().forEach(function (shape) {
    var tag = (shape.getTitle() || "").trim();
    if (!tag) return;
    if (!shapeByTag[tag]) shapeByTag[tag] = [];
    shapeByTag[tag].push(shape);
  });

  Object.keys(grid).forEach(function (tag) {
    var shapes = shapeByTag[tag];
    if (!shapes) return;
    var raw = grid[tag];
    var formatted = raw === undefined || raw === null ? "" : (typeof raw === "number" ? raw.toLocaleString("es-CL") : String(raw));
    shapes.forEach(function (shape) {
      if (shape.getText().asString() !== formatted) {
        shape.getText().setText(formatted);
      }
    });
  });
}

// Construye el layout de la Slide de Resumen semanal (título + hasta
// SUMMARY_MAX_SLIDE_ROWS filas de supervisor x 9 columnas), ya tageado.
// Solo se usa una vez para arrancar una Slide en blanco.
function buildSummarySlide_(slideId) {
  var presentation = SlidesApp.openById(slideId);
  var slide = presentation.getSlides()[0];
  var pageW = presentation.getPageWidth();
  var pageH = presentation.getPageHeight();

  function addBox(text, x, y, w, h, fontSize, bold, tag) {
    var box = slide.insertTextBox(text, x, y, w, h);
    var tr = box.getText();
    tr.getTextStyle().setFontSize(fontSize).setBold(!!bold);
    tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    if (tag) box.setTitle(tag);
    return box;
  }

  addBox("RESUMEN SEMANA", pageW * 0.05, pageH * 0.04, pageW * 0.9, pageH * 0.10, 24, true, "R1C1");

  var headers = ["Supervisor", "Día 1", "Día 2", "Día 3", "Día 4", "Día 5", "Día 6", "Promedio", "Total"];
  var startX = pageW * 0.03;
  var tableW = pageW * 0.94;
  var supColW = tableW * 0.18;
  var colW = (tableW - supColW) / 8;
  var colWidths = [supColW, colW, colW, colW, colW, colW, colW, colW, colW];

  var headerY = pageH * 0.16;
  var headerH = pageH * 0.06;
  var rowH = (pageH * 0.92 - headerY - headerH) / SUMMARY_MAX_SLIDE_ROWS;

  var x = startX;
  headers.forEach(function (h, i) {
    // i=1..6 son las 6 columnas de día -> se sincronizan con la fecha real.
    var tag = i >= 1 && i <= 6 ? "R1C" + (i + 1) : null;
    addBox(h, x, headerY, colWidths[i], headerH, 12, true, tag);
    x += colWidths[i];
  });

  for (var r = 0; r < SUMMARY_MAX_SLIDE_ROWS; r++) {
    var y = headerY + headerH + r * rowH;
    x = startX;
    addBox("-", x, y, supColW, rowH, 13, true, "R" + (r + 2) + "C1");
    x += supColW;
    for (var c = 0; c < 6; c++) {
      addBox("-", x, y, colW, rowH, 13, false, "R" + (r + 2) + "C" + (c + 2));
      x += colW;
    }
    addBox("-", x, y, colW, rowH, 13, true, "R" + (r + 2) + "C8");
    x += colW;
    addBox("-", x, y, colW, rowH, 13, true, "R" + (r + 2) + "C9");
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || "read";

  if (action === "summaryRead") {
    var summaryAreaRead = SUMMARY_AREAS[params.area];
    if (!summaryAreaRead) return jsonOutput_({ error: "unknown_area" });
    return jsonOutput_(readSummary_(summaryAreaRead, params.week || ""));
  }

  if (action === "summaryUpdate") {
    if (params.key !== CONTROL_KEY) return jsonOutput_({ error: "unauthorized" });
    var summaryAreaWrite = SUMMARY_AREAS[params.area];
    if (!summaryAreaWrite) return jsonOutput_({ error: "unknown_area" });
    var incomingRows;
    try {
      incomingRows = JSON.parse(params.rows || "[]");
    } catch (e2) {
      return jsonOutput_({ error: "invalid_rows" });
    }
    var summaryResult = writeSummary_(summaryAreaWrite, params.week || "", incomingRows);
    summaryResult.ok = true;
    return jsonOutput_(summaryResult);
  }

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
