/**
 * main.gs
 * Cambios mínimos: mejorar carga en Sheets usando la fila de encabezados para mapear datos
 * y añadir logs para depuración.
 */

const SPREADSHEET_ID = "xxxxxxxxxxxxxxxxxxxxxx";
const SHEET_NAME = "Applications";
const LISTASFIJAS_SHEET = "ListasFijas";

function doGet(e) {
  if (e && e.parameter && e.parameter.mod) {
    return HtmlService.createHtmlOutputFromFile('ui/' + e.parameter.mod);
  }
  const tpl = HtmlService.createTemplateFromFile('ui/index');
  return tpl.evaluate()
    .setTitle("Pasar Modular")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getModuloHTML(nombre) {
  return HtmlService.createHtmlOutputFromFile('ui/' + nombre).getContent();
}

// ============================================================
// 📋 OBTENER LISTAS FIJAS (Para inicializar el formulario)
// ============================================================
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

// ============================================================
// 🧹 UTILIDAD: NORMALIZAR TEXTOS (Ignora mayúsculas, tildes y espacios)
// ============================================================
function limpiarTextoGS(texto) {
  if (!texto) return "";
  return texto.toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Quita acentos/tildes
    .trim()
    .toLowerCase();
}

// ============================================================
// 📍 OBTENER ESTADOS POR REGIÓN (VERSIÓN CON DIAGNÓSTICO)
// ============================================================
function getEstadosPorRegion(region) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Lista_Estado"); 
  
  // 1. ¿Existe la pestaña?
  if (!sheet) return ["❌ ERROR: No se encontró la pestaña 'Lista_Estado'"];
  
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return ["❌ ERROR: La pestaña está totalmente vacía"];

  const headers = data[0].map(limpiarTextoGS);
  const colRegion = headers.indexOf(limpiarTextoGS("Region"));
  const colEstado = headers.indexOf(limpiarTextoGS("Estado"));

  // 2. ¿Encontró las columnas en la Fila 1?
  if (colRegion < 0 || colEstado < 0) {
    return ["❌ ERROR COLUMNAS: La fila 1 tiene estos datos -> " + data[0].join(" | ")];
  }

  const regionBuscada = limpiarTextoGS(region);

  const resultados = data.slice(1)
    .filter(r => limpiarTextoGS(r[colRegion]) === regionBuscada)
    .map(r => r[colEstado])
    .filter(v => v);

  // 3. ¿Encontró la región escrita?
  if (resultados.length === 0) {
    return ["❌ ERROR DATOS: No se encontró la región '" + region + "' en la columna Region"];
  }

  return [...new Set(resultados)];
}

  // ============================================================
  // 👨‍💼 OBTENER GERENTES POR REGIÓN
  // Busca en la pestaña "Lista_Gerente"
  // ============================================================
  function getGerentesPorRegion(region) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Lista_Gerente"); // <--- Pestaña correcta
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(limpiarTextoGS);
    const colRegion = headers.indexOf(limpiarTextoGS("Region"));
    const colGerente = headers.indexOf(limpiarTextoGS("Gerente"));

    if (colRegion < 0 || colGerente < 0) return [];

    const regionBuscada = limpiarTextoGS(region);

    return data.slice(1)
      .filter(r => limpiarTextoGS(r[colRegion]) === regionBuscada)
      .map(r => r[colGerente])
      .filter(v => v);
  }

  // ============================================================
  // 🏙️ OBTENER MUNICIPIOS POR ESTADO
  // Busca en la pestaña "Lista_Municipio"
  // ============================================================
  function getMunicipiosPorEstado(estado) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Lista_Municipio"); // <--- Pestaña correcta
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(limpiarTextoGS);
    const colEstado = headers.indexOf(limpiarTextoGS("Estado"));
    const colMunicipio = headers.indexOf(limpiarTextoGS("Municipio"));

    if (colEstado < 0 || colMunicipio < 0) return [];

    const estadoBuscado = limpiarTextoGS(estado);

    return data.slice(1)
      .filter(r => limpiarTextoGS(r[colEstado]) === estadoBuscado)
      .map(r => r[colMunicipio])
      .filter(v => v);
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

  // ============================================================
  // ✅ LOG 1: Datos que llegan desde el frontend
  // ============================================================
  Logger.log("🟢 cargarDemanda recibido (ANTES de automáticos): " + JSON.stringify(datos));

  // obtener encabezados de la hoja (fila 1)
  const headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headersRange.getValues()[0].map(h => (h || "").toString().trim());

  // función para normalizar header a la clave esperada por el cliente
  function keyFromHeader(h) {
    return (h || "")
      .toString()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9_]/g, "")
      .toLowerCase();
  }

  // ============================================================
  // 🆔 GENERAR ID UNICO gpon-[número]-[DDMMAAAA]-[HHMMSS]
  // ============================================================
  const idCol = headers.findIndex(h => keyFromHeader(h) === 'id') + 1;
  let nextNum = 1;

  if (idCol > 0 && sheet.getLastRow() > 1) {
    const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues().flat();
    const numeros = ids
      .map(v => {
        if (v && typeof v === "string" && v.startsWith("gpon-")) {
          const n = parseInt(v.split("-")[1]);
          return isNaN(n) ? 0 : n;
        }
        return 0;
      })
      .filter(n => n > 0);

    if (numeros.length > 0) {
      nextNum = Math.max(...numeros) + 1;
    }
  }

  const fecha = Utilities.formatDate(new Date(), "America/Caracas", "ddMMyyyy");
  const hora  = Utilities.formatDate(new Date(), "America/Caracas", "HHmmss");
  const nuevoID = `gpon-${nextNum}-${fecha}-${hora}`;
  datos.id = nuevoID;

  // ============================================================
  // ✅ CAMPOS AUTOMÁTICOS — AHORA EN EL LUGAR CORRECTO
  // ============================================================
  datos["tipo_data"] = "Nuevo";
  datos["usuario_registro"] = getUserEmail();
  datos["fecha_registro"] = Utilities.formatDate(
    new Date(),
    "America/Caracas",
    "dd/MM/yyyy HH:mm:ss"
  );

  // ============================================================
  // ✅ LOG 2: Datos después de agregar automáticos
  // ============================================================
  Logger.log("📌 Datos DESPUÉS de automáticos: " + JSON.stringify(datos));

  // ============================================================
  // ✅ LOG 3: Headers y su normalización
  // ============================================================
  const headersDebug = headers.map(h => ({
    headerOriginal: h,
    headerNormalizado: keyFromHeader(h)
  }));
  Logger.log("📋 Headers detectados en Sheets: " + JSON.stringify(headersDebug));

    // ============================================================
    // ✅ VALIDACIONES DE FECHAS
    // ============================================================

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    const fechaAbordaje = datos.fechaabordaje ? new Date(datos.fechaabordaje) : null;
    const fechaAceptacion = datos.fechaaceptacion ? new Date(datos.fechaaceptacion) : null;

    // ✅ Regla 1: Fecha Abordaje no puede ser mayor que hoy
    if (fechaAbordaje && fechaAbordaje > hoy) {
      return { ok: false, campo: "fechaabordaje", mensaje: "La fecha de abordaje no puede ser mayor al día de hoy." };
    }

    // ✅ Regla 2: Si oferta aceptada = "Sí", fecha aceptación es obligatoria
    if (datos.ofertaaceptada === "Sí" && !fechaAceptacion) {
      return { ok: false, campo: "fechaaceptacion", mensaje: "Debe ingresar la fecha de aceptación." };
    }

    // ✅ Regla 3: Si oferta aceptada = "No", fecha aceptación debe estar vacía
    if (datos.ofertaaceptada === "No" && fechaAceptacion) {
      return { ok: false, campo: "fechaaceptacion", mensaje: "No debe ingresar fecha de aceptación si la oferta no fue aceptada." };
    }

    // ✅ Regla 4: Fecha Aceptación no puede ser mayor que hoy
    if (fechaAceptacion && fechaAceptacion > hoy) {
      return { ok: false, campo: "fechaaceptacion", mensaje: "La fecha de aceptación no puede ser mayor al día de hoy." };
    }

    // ✅ Regla 5: Fecha Abordaje no puede ser mayor que Fecha Aceptación
    if (fechaAbordaje && fechaAceptacion && fechaAbordaje > fechaAceptacion) {
      return { ok: false, campo: "fechaabordaje", mensaje: "La fecha de abordaje no puede ser mayor que la fecha de aceptación." };
    }

    // ✅ Regla 6: Fecha Aceptación no puede ser menor que Fecha Abordaje
    if (fechaAbordaje && fechaAceptacion && fechaAceptacion < fechaAbordaje) {
      return { ok: false, campo: "fechaaceptacion", mensaje: "La fecha de aceptación no puede ser menor que la fecha de abordaje." };
    }

    // ✅ Mes y Año Abordaje automáticos
    if (fechaAbordaje) {
      const meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
      datos.mesabordaje = meses[fechaAbordaje.getMonth()];
      datos.anoabordaje = fechaAbordaje.getFullYear().toString();
    }

  // ============================================================
  // ✅ Construir fila respetando el orden de headers
  // ============================================================
  const fila = headers.map(h => {
    const k = keyFromHeader(h);
    if (k in datos) return datos[k];
    if (h in datos) return datos[h];
    return "";
  });

  // ============================================================
  // ✅ LOG 4: Fila final que se insertará
  // ============================================================
  Logger.log("✅ Fila FINAL a insertar: " + JSON.stringify(fila));

  // insertar fila en la segunda posición
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, fila.length).setValues([fila]);

  return { ok: true, message: "Carga guardada", filaInserted: fila };
}

// ============================================================
// 📌 NUEVA FUNCIÓN: obtener correo Gmail del usuario autenticado
// ============================================================
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function logDebug(msg) {
  console.log("DEBUG: " + msg);
  Logger.log(msg);
}

// ============================================================
// 🔍 BUSCAR CLIENTE POR RIF (Para auto-rellenado Checkbox)
// ============================================================
function buscarClientePorRIF(rifBuscado) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return null;

  const headers = data[0];
  // Normalizamos igual que tu cargarDemanda
  const headersNormalizados = headers.map(h => (h || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9_]/g, "").toLowerCase());
  
  const colRifIndex = headersNormalizados.indexOf("rif");
  if (colRifIndex === -1) return null;

  const buscadoLimpio = rifBuscado.toString().replace(/[^a-zA-Z0-9]/g, "").toUpperCase();

  for (let i = data.length - 1; i >= 1; i--) {
    const valorCelda = data[i][colRifIndex].toString().replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
    
    if (valorCelda === buscadoLimpio) {
      console.log("✅ Match encontrado en fila " + (i + 1));
      
      const cliente = {};
      headersNormalizados.forEach((key, index) => {
        // 🔑 EL TRUCO: Convertimos todo a String para que no de NULL al viajar al HTML
        let valor = data[i][index];
        
        if (valor instanceof Date) {
          cliente[key] = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else {
          cliente[key] = (valor === null || valor === undefined) ? "" : valor.toString();
        }
      });
      
      return cliente; // Ahora sí viajará con datos
    }
  }
  return null;
}

// =====================================================================================
// 📊 INICIO MODULO DASHBOARD DE VENTAS
// =====================================================================================

/**
 * Función para obtener la data cruda del Dashboard de Ventas.
 * Extrae solo los 8 campos solicitados para mantener la velocidad.
 */
function getVentasDashData() {
  try {
    // 1. CONEXIÓN A LA HOJA
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // Asegúrate de que SPREADSHEET_ID esté en tu config
    const sheet = ss.getSheetByName(SHEET_NAME); // O el nombre exacto de tu hoja de Demandas
    
    if (!sheet) return { ok: false, error: "Hoja de Demandas no encontrada" };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: true, records: [] };

    // 2. NORMALIZACIÓN DE ENCABEZADOS (Para evitar errores por mayúsculas o espacios)
    const headers = data[0].map(h => (h || "").toString().trim().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9_]/g, ""));

    // 3. MAPEO EXACTO DE LOS 8 CAMPOS (Basado en tu instrucción)
    // Busca el índice (número de columna) de cada campo.
    const idx = {
      numero: headers.indexOf("numero"),
      region: headers.indexOf("region"),
      estado: headers.indexOf("estado"),
      huella: headers.indexOf("huella"),
      sector: headers.indexOf("sector"),
      situacion: headers.indexOf("situacionabordaje"),
      mes: headers.indexOf("mesabordaje"),
      ano: headers.indexOf("anoabordaje") // "año" se normaliza a "ano"
    };

    // 4. EXTRACCIÓN Y LIMPIEZA DE FILAS
    const records = data.slice(1).map(row => ({
      numero: row[idx.numero] || "N/A",
      region: row[idx.region] || "N/A",
      estado: row[idx.estado] || "N/A",
      huella: row[idx.huella] || "N/A",
      sector: row[idx.sector] || "N/A",
      situacion: row[idx.situacion] || "N/A",
      mes: row[idx.mes] || "N/A",
      ano: row[idx.ano] || "N/A"
    }));

    // 5. RETORNO AL FRONTEND
    return { ok: true, records: records };

  } catch (error) {
    return { ok: false, error: error.message };
  }
}
