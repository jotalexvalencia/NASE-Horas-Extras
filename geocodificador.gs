// =================================================================
// 📁 geocodificador.gs – Módulo de Geocodificación (NASE 2026 - Horas Extras)
// =================================================================
/**
 * @summary Módulo inteligente para geocodificación y control de asistencia por ubicación.
 * @description Administra la conversión de coordenadas GPS (Lat/Lng) a direcciones y ciudades.
 *              Calcula distancias (Haversine) entre el empleado y su centro asignado.
 *              **ACTUALIZACIÓN:** Índices ajustados al esquema de "Horas Extras" (Code.gs).
 * 
 * @author NASE Team
 * @version 2.1 (Ajuste de Columnas Horas Extras)
 */

// =================================================================
// 1. CONFIGURACIÓN DEL SISTEMA
// =================================================================

// Objeto global GEO que encapsula configuración y funciones (Namespace Pattern)
if (typeof GEO === 'undefined') var GEO = {};

GEO.CONFIG = {
  SHEET_RESPUESTAS: "Respuestas", // Nombre de la hoja principal
  SHEET_CENTROS: "Centros",      // Nombre de la hoja de referencia de centros
  
  // ✅ MAPEO DE COLUMNAS ACTUALIZADO (Basado en RESP_HEADERS de Code.gs)
  // Indices 0-based basados en:
  // 0:Ced, 1:Centro, 2:Ciudad, 3:Lat, 4:Lng, 5:Acepto, 6:Ciudad_Geo, 7:Dir_Geo, 8:Accuracy, 9:Dentro, 10:Distancia...
  
  R_COL_LAT: 3,            // Columna D: Latitud GPS (Índice 3)
  R_COL_LNG: 4,            // Columna E: Longitud GPS (Índice 4)
  R_COL_CIUDAD_GEO: 6,    // Columna G: Ciudad Geocodificada (Índice 6)
  R_COL_DIR_GEO: 7,       // Columna H: Dirección Geocodificada (Índice 7)
  R_COL_ACCURACY: 8,     // Columna I: Precisión GPS (Índice 8)
  R_COL_DENTRO_CENTRO: 9, // Columna J: ¿Está dentro del centro? (Índice 9)
  R_COL_DISTANCIA: 10,     // Columna K: Distancia al centro en metros (Índice 10)
 
  // Configuración de OpenCage (API Pagada de respaldo)
  OPENCAGE_DAILY_LIMIT: 2500, // Límite de peticiones gratuitas diarias (aproximado)
  OPENCAGE_API_PROP: 'OPENCAGE_API_KEY', // Nombre de la propiedad donde se guarda la API Key
  OPENCAGE_QUOTA_PROP: 'OPENCAGE_QUOTA',    // Propiedad para contar uso diario
  
  // Retrasos para APIs gratuitas (evitar bloqueos por Rate Limiting)
  REQUEST_DELAY_MS: 300
};

/**
 * @summary Guarda la API Key de OpenCage de forma segura.
 * @description Esta función debe ejecutarse UNA SOLA VEZ manualmente desde el editor.
 */
function guardarApiKeyOpenCage() {
  // ⚠️ ATENCIÓN: Reemplaza 'TU_CLAVE_API_DE_OPENCAGE' con la clave real
  const apiKey = '0f4d42a072704ffc8ad51d03a21fcea0';
  PropertiesService.getScriptProperties().setProperty(GEO.CONFIG.OPENCAGE_API_PROP, apiKey);
  Logger.log('✅ API Key de OpenCage guardada de forma segura.');
  SpreadsheetApp.getUi().alert('✅ API Key de OpenCage guardada correctamente. Ya puedes borrar esta función si lo deseas.');
}

// =================================================================
// 2. FUNCIONES AUXILIARES (Normalización y Cálculo)
// =================================================================

/**
 * @summary Normaliza coordenadas.
 * @param {String|Number} v - Valor de latitud/longitud.
 * @returns {Number|Null} Valor flotante o null si es inválido.
 */
GEO.normalizarCoord = function(v) {
  if (!v) return null;
  let s = String(v).replace(/,/g, '.').trim();
  return isNaN(s) ? null : parseFloat(s);
};

/**
 * @summary Determina la ciudad basándose en reglas de negocio y coordenadas.
 */
GEO.normalizarCiudadLocal = function(ciudad, lat, lng) {
  if (!ciudad) return "No encontrado";
  const l = ciudad.toString().toLowerCase();
  
  if (l.includes("risaralda") || l.includes("eje") || l.includes("amco") || l.includes("perimetro urbano pereira") || l.includes("perímetro urbano pereira")) {
    return lat >= 4.825 ? "Dosquebradas" : "Pereira";
  }
  if (l.includes("valle") || l.includes("pac") || l.includes("perímetro urbano santiago de cali") || l.includes("cali")) return "Cali";
  if (l.includes("cartago")) return "Cartago";
  if (l.includes("bogot") || l.includes("cundinamarca")) return "Bogotá";
  if (l.includes("medellín") || l.includes("perímetro urbano medellín")) return "Medellín";
  
  return ciudad.replace(/^(Per[íi]metro urbano|AMCO|Area Metropolitana Centro Occidente|.*,\s*)+/gi, '').split(',')[0].trim() || "No encontrado";
};

/**
 * @summary Limpia la dirección eliminando código postal y país.
 */
GEO.limpiarDireccion = function(dir) {
  if (!dir) return "Sin dirección específica";
  return dir.split(',').map(p => p.trim()).filter(p => p && !/^\d{5,}$/i.test(p) && !/colombia|rap eje|risaralda|valle|cundinamarca|antioquia|66000|16000|05000|76000/i.test(p.toLowerCase())).join(', ').replace(/\s+/g, ' ').trim();
};

/**
 * @summary Calcula la distancia entre dos puntos GPS (Fórmula Haversine).
 */
GEO.calcularDistancia = function(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a =
    Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon/2) * Math.sin(dLon/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
};

/**
 * @summary Busca dinámicamente el índice de una columna por nombre.
 */
GEO.findHeaderIndex = function(headers, names) {
  for (let name of names) {
    const index = headers.findIndex(h => h && h.toString().toLowerCase().trim() === name.toLowerCase().trim());
    if (index !== -1) return index + 1; // Devuelve el número de columna (1-based)
  }
  return -1;
};

// =================================================================
// 3. CONTROL DE CUOTAS DE OPENCAGE (Optimización de Costos)
// =================================================================

GEO.checkAndUpdateOpenCageQuota = function() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date().toISOString().slice(0, 10);
  
  let quotaData = JSON.parse(scriptProperties.getProperty(GEO.CONFIG.OPENCAGE_QUOTA_PROP) || '{"date": "", "count": 0}');

  if (quotaData.date !== today) {
    quotaData.date = today;
    quotaData.count = 0;
  }

  if (quotaData.count >= GEO.CONFIG.OPENCAGE_DAILY_LIMIT) {
    Logger.log(`❌ LÍMITE DIARIO DE OPENCAGE ALCANZADO (${GEO.CONFIG.OPENCAGE_DAILY_LIMIT}).`);
    return false; 
  }

  quotaData.count++;
  scriptProperties.setProperty(GEO.CONFIG.OPENCAGE_QUOTA_PROP, JSON.stringify(quotaData));
  Logger.log(`✅ Llamada a OpenCage permitida. (${quotaData.count}/${GEO.CONFIG.OPENCAGE_DAILY_LIMIT}) hoy.`);
  return true; 
};

// =================================================================
// 4. FUENTES DE GEOCODIFICACIÓN (Wrappers de API)
// =================================================================

GEO.getFromMapsCo = (lat, lng) => {
  const url = `https://geocode.maps.co/reverse?lat=${lat}&lon=${lng}&accept-language=es`;
  const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return JSON.parse(r.getContentText());
};

GEO.getFromNominatim = (lat, lng) => {
  const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${lat}&lon=${lng}&addressdetails=1&accept-language=es`;
  const r = UrlFetchApp.fetch(url, { headers: { 'User-Agent': 'NASE-Geocoder/1.0 (contact: analistaoperaciones@nasecolombia.com.co)' } });
  return JSON.parse(r.getContentText());
};

GEO.getFromOpenCage = (lat, lng) => {
  if (!GEO.checkAndUpdateOpenCageQuota()) {
    return { error: 'Límite diario de OpenCage alcanzado.' };
  }
  
  const apiKey = PropertiesService.getScriptProperties().getProperty(GEO.CONFIG.OPENCAGE_API_PROP);
  if (!apiKey) return { error: 'API Key de OpenCage no configurada.' };
  
  const url = `https://api.opencagedata.com/geocode/v1/json?q=${lat}+${lng}&key=${apiKey}&language=es&limit=1`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(response.getContentText());

  if (response.getResponseCode() === 200 && json.results && json.results.length > 0) {
    const result = json.results[0];
    const components = result.components;
    const ciudad = GEO.normalizarCiudadLocal(components.city || components.town || components.county || '', lat, lng);
    const direccion = GEO.limpiarDireccion(result.formatted);
    return { ciudad: ciudad, direccion: direccion, fuente: 'OpenCage' };
  }
  return null;
};

// =================================================================
// 5. LÓGICA PRINCIPAL DE GEOCODIFICACIÓN
// =================================================================

GEO.reverseGeocodeInternal = function(lat, lng) {
  let result = null;
  let source = '';

  try {
    const j = GEO.getFromMapsCo(lat, lng);
    if (j && j.display_name) {
      const a = j.address || {};
      const ciudad = GEO.normalizarCiudadLocal(a.city || a.town || a.county || '', lat, lng);
      const direccion = GEO.limpiarDireccion(j.display_name);
      result = { ciudad: ciudad, direccion: direccion };
      source = 'Maps.co';
    }
  } catch (e) { Logger.log(e); }

  if (!result) {
    try {
      const j = GEO.getFromNominatim(lat, lng);
      if (j && j.display_name) {
        const a = j.address || {};
        const ciudad = GEO.normalizarCiudadLocal(a.city || a.town || a.county || '', lat, lng);
        const direccion = GEO.limpiarDireccion(j.display_name);
        result = { ciudad: ciudad, direccion: direccion };
        source = 'Nominatim';
        Utilities.sleep(1000);
      }
    } catch (e) { Logger.log(e); }
  }
 
  if (!result) {
    const openCageResult = GEO.getFromOpenCage(lat, lng);
    if (openCageResult && !openCageResult.error) {
      result = { ciudad: openCageResult.ciudad, direccion: openCageResult.direccion };
      source = 'OpenCage';
    } else if (openCageResult && openCageResult.error) {
       return { ciudad: 'Error', direccion: openCageResult.error, fuente: 'OpenCage' };
    }
  }

  if (!result) {
    result = { ciudad: 'No encontrado', direccion: 'Sin dirección específica' };
    source = 'Ninguna';
  }
 
  result.fuente = source;
  result.accuracy = 0;
 
  if (!source.includes('Nominatim')) {
    Utilities.sleep(GEO.CONFIG.REQUEST_DELAY_MS);
  }
 
  return result;
};

// =================================================================
// 6. FUNCIÓN PÚBLICA PRINCIPAL (Trigger de Usuario)
// =================================================================

/**
 * @summary Función principal que se ejecuta desde el Menú.
 * @description Corrige los índices de escritura para coincidir con la hoja Horas Extras.
 */
function geocodificarFilaActiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRes = ss.getActiveSheet();
  
  if (hojaRes.getName() !== GEO.CONFIG.SHEET_RESPUESTAS) {
    SpreadsheetApp.getUi().alert("Por favor, ejecuta esta función desde la hoja 'Respuestas'.");
    return;
  }

  // ----------------------------------------------------------------
  // 1. OBTENER DATOS DE LA FILA ACTIVA
  // ----------------------------------------------------------------
  const fila = hojaRes.getActiveCell().getRow();
  
  if (fila === 1) { 
    SpreadsheetApp.getUi().alert("Selecciona una fila con datos, no el encabezado."); 
    return; 
  }

  const headers = hojaRes.getRange(1, 1, 1, hojaRes.getLastColumn()).getValues()[0];
  const getCol = (name) => GEO.findHeaderIndex(headers, [name]);

  // Usamos los índices corregidos para leer
  const lat = GEO.normalizarCoord(hojaRes.getRange(fila, GEO.CONFIG.R_COL_LAT + 1).getValue());
  const lng = GEO.normalizarCoord(hojaRes.getRange(fila, GEO.CONFIG.R_COL_LNG + 1).getValue());
  const nombreCentro = (hojaRes.getRange(fila, getCol("centro")).getValue() || "").toString().trim();
  const ciudadCentro = (hojaRes.getRange(fila, getCol("ciudad")).getValue() || "").toString().trim();

  if (lat == null || lng == null || !nombreCentro) {
    SpreadsheetApp.getUi().alert("Faltan datos clave en la fila: Latitud, Longitud o Centro.");
    return;
  }

  // ----------------------------------------------------------------
  // 2. BUSCAR EL CENTRO EN LA HOJA 'Centros'
  // ----------------------------------------------------------------
  const hojaCentros = ss.getSheetByName(GEO.CONFIG.SHEET_CENTROS);
  if (!hojaCentros) { SpreadsheetApp.getUi().alert("La hoja 'Centros' no existe."); return; } 
  
  const dataCentros = hojaCentros.getDataRange().getValues();
  const headersCentros = dataCentros[0];
  const getColC = (name) => GEO.findHeaderIndex(headersCentros, [name]);

  let centroRef = null;
  
  for (let i = 1; i < dataCentros.length; i++) {
    const rowC = dataCentros[i];
    const nombreC = (rowC[getColC("centro") - 1] || "").toString().trim();
    const ciudadC = (rowC[getColC("ciudad") - 1] || "").toString().trim();
   
    if (nombreC.toUpperCase() === nombreCentro.toUpperCase() && ciudadC.toUpperCase() === ciudadCentro.toUpperCase()) {
      centroRef = {
        lat: GEO.normalizarCoord(rowC[getColC("lat") - 1]), // Asumiendo header "Lat" en centros
        lng: GEO.normalizarCoord(rowC[getColC("lng") - 1]), // Asumiendo header "Lng" en centros
        radio: Number(rowC[getColC("radio") - 1]) || 30,
        nombre: nombreC,
        ciudad: ciudadC
      };
      break;
    }
  }

  if (!centroRef) {
    SpreadsheetApp.getUi().alert(`No se encontró el centro "${nombreCentro}" en la ciudad "${ciudadCentro}" en la hoja 'Centros'.`);
    return;
  }

  const distancia = GEO.calcularDistancia(lat, lng, centroRef.lat, centroRef.lng);
  const estaDentro = distancia <= centroRef.radio;

  // ----------------------------------------------------------------
  // 3. ESCRIBIR RESULTADOS EN LA HOJA (Usando indices corregidos)
  // ----------------------------------------------------------------
  // Recordar: GEO.CONFIG tiene indices 0-based, getRange usa 1-based. Sumamos 1.
  const rangoCiudad = hojaRes.getRange(fila, GEO.CONFIG.R_COL_CIUDAD_GEO + 1);
  const rangoDir = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DIR_GEO + 1);
  const rangoDentro = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DENTRO_CENTRO + 1);
  const rangoDistancia = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DISTANCIA + 1);

  if (estaDentro) {
    rangoCiudad.setValue(centroRef.ciudad);
    rangoDir.setValue(centroRef.nombre);
    rangoDentro.setValue("Sí");
    rangoDistancia.setValue(Number(distancia.toFixed(2)));
    
    SpreadsheetApp.getUi().alert(`✅ La ubicación está DENTRO del centro "${centroRef.nombre}".\nNo se consumieron cuotas de geocodificación.`);
  } else {
    rangoDentro.setValue("No");
    rangoDistancia.setValue(Number(distancia.toFixed(2)));
    
    SpreadsheetApp.getUi().alert('📍 La ubicación está FUERA del centro. Geocodificando con APIs externas... Por favor, espere.');
    
    const resultado = GEO.reverseGeocodeInternal(lat, lng);
   
    rangoCiudad.setValue(resultado.ciudad);
    rangoDir.setValue(resultado.direccion);
    hojaRes.getRange(fila, GEO.CONFIG.R_COL_ACCURACY + 1).setValue(resultado.fuente);
   
    SpreadsheetApp.getUi().alert(`✔ Geocodificación completada.\n\nCiudad: ${resultado.ciudad}\nDirección: ${resultado.direccion}\nFuente: ${resultado.fuente}`);
  }
}
