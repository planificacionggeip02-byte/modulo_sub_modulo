/**
 * @file main.gs
 * Punto de entrada principal del proyecto Pasar Modular.
 * Esta versión está preparada para ejecutarse como una aplicación web
 * (idéntica al comportamiento original del proyecto ggeip.gpon.modular).
 * 
 * Cuando se despliegue como WebApp, Google Apps Script llamará
 * automáticamente a la función doGet(), que devuelve la interfaz HTML.
 */

// ============================================================
// 🗂️ CONFIGURACIÓN BASE
// ============================================================
const SPREADSHEET_ID = "1V2lfTB51FioZUYKdAvSxe9odzz1kZrgY5pF0F0jB_HE";
const SHEET_NAME = "Applications";
const LISTASFIJAS_SHEET = "ListasFijas";

// ============================================================
// 🌐 FUNCIÓN PRINCIPAL DE ACCESO WEB
// ============================================================
function doGet(e) {
  console.log("🟢 doGet ejecutado con parámetros:", JSON.stringify(e));

  // Si viene un parámetro "mod", cargamos directamente ese módulo HTML
  if (e && e.parameter.mod) {
    console.log("📂 Cargando módulo directo:", e.parameter.mod);
    return HtmlService.createHtmlOutputFromFile(`ui/${e.parameter.mod}`);
  }

  // Si no hay parámetro, cargamos el index principal como plantilla
  console.log("📂 Cargando index principal");

  const tpl = HtmlService.createTemplateFromFile('ui/index'); // usar plantilla
  const html = tpl.evaluate()                                 // evaluar para procesar includes
    .setTitle("📋 Pasar Modular — Formulario Principal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return html;
}
// ============================================================
// ⚙️ FUNCIONES DE UTILIDAD
// ============================================================
function include(filename) {
  console.log("📥 include llamado con:", filename);
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getModuloHTML(nombre) {
  console.log("📥 getModuloHTML llamado con:", nombre);
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/' + nombre).getContent();
    console.log("✅ Módulo encontrado:", 'ui/' + nombre);
    return html;
  } catch (err) {
    console.error("❌ Error cargando módulo:", nombre, err);
    throw err;
  }
}

// ============================================================
// 📋 FUNCIONES DE LISTAS (para selects fijos/dependientes)
// ============================================================
function getListasFijas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(LISTASFIJAS_SHEET);
  if (!sheet) throw new Error("❌ No se encontró la hoja 'ListasFijas'");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const map = {};
  for (let c = 0; c < headers.length; c++) {
    const header = (headers[c] || "").toString().trim();
    if (!header) continue;

    const values = [];
    for (let r = 1; r < data.length; r++) {
      const v = (data[r][c] || "").toString().trim();
      if (v) values.push(v);
    }
    map[header] = values;
  }
  return map;
}

// ============================================================
// 📋 FUNCIÓN DE REGISTRO (guardar datos en Sheets)
// RENOMBRADA A cargarDemanda para uniformidad con UI
// ============================================================
function cargarDemanda(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  // Construir fila con los valores recibidos
  const fila = Object.values(datos);
  fila.push(new Date()); // fecha de registro

  // Inserción en fila 2, desplazando hacia abajo
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, fila.length).setValues([fila]);

  return "✅ Carga de demanda guardada correctamente.";
}

function logDebug(msg) {
  console.log("🟢 [DEBUG] " + msg);
  Logger.log("🟢 [DEBUG] " + msg);
}
