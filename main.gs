/**
 * main.gs
 * Cambios mínimos: mejorar carga en Sheets usando la fila de encabezados para mapear datos
 * y añadir logs para depuración.
 */

const SPREADSHEET_ID = "1V2lfTB51FioZUYKdAvSxe9odzz1kZrgY5pF0F0jB_HE";
const SHEET_NAME = "Applications";
const LISTASFIJAS_SHEET = "ListasFijas";

function doGet(e) {
  if (e && e.parameter && e.parameter.mod) {
    return HtmlService.createHtmlOutputFromFile('ui/' + e.parameter.mod);
  }
  const tpl = HtmlService.createTemplateFromFile('ui/index');
  return tpl.evaluate().setTitle("Pasar Modular").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getModuloHTML(nombre) {
  return HtmlService.createHtmlOutputFromFile('ui/' + nombre).getContent();
}

function getListasFijas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(LISTASFIJAS_SHEET);
  if (!sheet) throw new Error("No se encontró la hoja 'ListasFijas'");
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const map = {};
  for (let c = 0; c < headers.length; c++) {
    const header = (headers[c] || "").toString().trim();
    if (!header) continue;
    const col = [];
    for (let r = 1; r < data.length; r++) {
      const v = (data[r][c] || "").toString().trim();
      if (v) col.push(v);
    }
    map[header] = col;
  }
  return map;
}

/**
 * guardar datos en sheet respetando el orden de columnas:
 * Se mapea cada encabezado -> key normalizado (normalizeKey) y se extrae datos[key]
 * Esto evita que Object.values(datos) desordene columnas.
 */
function cargarDemanda(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("No se encontró la hoja '" + SHEET_NAME + "'");
  Logger.log("🟢 cargarDemanda recibido: " + JSON.stringify(datos));

  // obtener encabezados de la hoja (fila 1)
  const headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headersRange.getValues()[0].map(h => (h || "").toString().trim());

  // función para normalizar header a la clave esperada por el cliente
  function keyFromHeader(h) {
    return (h || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9_]/g, "").toLowerCase();
  }

  const fila = headers.map(h => {
    const k = keyFromHeader(h);
    // si datos tiene la clave normalizada (ej 'nombrecontacto'), retornarla
    if (k in datos) return datos[k];
    // también intentar con header original (en caso el cliente envía keys con mayúsculas/espacios)
    if (h in datos) return datos[h];
    // si el cliente envía claves adicionales, ignorarlas si no mapean
    return "";
  });

  // append o insertar row en la segunda fila para mantener orden (como antes insertRowBefore(2))
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, fila.length).setValues([fila]);

  Logger.log("✅ Fila insertada: " + JSON.stringify(fila));
  return { ok: true, message: "Carga guardada", filaInserted: fila };
}

function logDebug(msg) {
  console.log("DEBUG: " + msg);
  Logger.log(msg);
}
