/**
 * @file core/code.gs
 * 📁 CONFIGURACIÓN BASE DEL PROYECTO PASAR MODULAR
 * Este archivo concentra las variables globales que identifican
 * el archivo principal de datos y las hojas que utiliza el sistema.
 */

// ============================================================
// ⚙️ OBJETOS Y FUNCIONES DE CONFIGURACIÓN
// ============================================================

/**
 * Devuelve una referencia al archivo principal del sistema.
 * @returns {SpreadsheetApp.Spreadsheet} Archivo de cálculo activo según SPREADSHEET_ID.
 */
function getMainSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Devuelve una hoja específica dentro del archivo principal.
 * @param {string} name - Nombre de la hoja que se desea obtener.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Hoja solicitada.
 */
function getSheetByName(name) {
  const ss = getMainSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error("❌ No se encontró la hoja: " + name);
  return sh;
}

// ============================================================
// 📥 CARGA DE MÓDULOS HTML
// ============================================================

/**
 * Devuelve el contenido HTML del submódulo Registrar Usuario.
 * Nombre de archivo exacto en el editor: ui/registrar_usuario (sin .html)
 */
function getModuloHTML(nombre) {
  const FIXED_NAME = "ui/registrar_usuario";  // usamos SIEMPRE este nombre
  const requested = nombre;

  try {
    Logger.log("🚀 getModuloHTML invocado");
    Logger.log("🧩 nombre solicitado (cliente): " + requested);
    Logger.log("🧩 nombre usado (servidor): " + FIXED_NAME);

    var output = HtmlService.createHtmlOutputFromFile(FIXED_NAME).getContent();
    var len = (output && output.length) ? output.length : 0;
    Logger.log("📦 contenido devuelto (bytes): " + len);

    if (!output || output.trim() === "") {
      Logger.log("⚠️ contenido vacío: devolviendo fallback diagnóstico");
      output = '<div style="padding:12px;color:#b00020;">[DIAGNÓSTICO] El archivo ui/registrar_usuario se cargó pero devolvió vacío.</div>';
    }

    return output;
  } catch (err) {
    Logger.log("❌ Error en getModuloHTML: " + err.message);
    throw new Error("❌ No se pudo cargar ui/registrar_usuario → " + err.message);
  }
}

/**
 * Prueba directa desde el editor para ver logs en “Ejecuciones”.
 */
function test_getModuloHTML() {
  Logger.log("▶️ Iniciando prueba: ui/registrar_usuario");
  try {
    var html = HtmlService.createHtmlOutputFromFile("ui/registrar_usuario").getContent();
    Logger.log("✅ OK: longitud de HTML = " + (html ? html.length : 0));
    if (!html || html.trim() === "") {
      Logger.log("⚠️ El archivo existe pero el contenido está vacío.");
    }
  } catch (e) {
    Logger.log("❌ Falló la carga: " + e.message);
    throw e;
  }
}
