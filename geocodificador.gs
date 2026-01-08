// =================================================================
// 📁 geocodificador.gs – Módulo de Geocodificación (NASE 2026 - Horas Extras)
// ----------------------------------------------------------------------
/**
 * @summary Módulo inteligente para geocodificación y control de asistencia.
 * @description Administra la conversión de coordenadas GPS (Lat/Lng) a direcciones.
 *              Calcula distancias (Haversine) contra el libro de Centros EXTERNO.
 *              Ajuste de índices a la estructura original solicitada por el usuario.
 *
 * @features
 *   - 🌍 Geocodificación inversa (APIs OpenCage, Maps.co, Nominatim).
 *   - 📏 Cálculo de distancia al centro asignado.
 *   - 🚦 Semáforo "Dentro/Fuera" del centro.
 *   - 📂 Lectura de "Centros" desde el Libro Base Operativa.
 *
 * @author NASE Team
 * @version 2.2 (Lectura Externa + Índices Originales)
 */

// =================================================================
// 1. CONFIGURACIÓN DEL SISTEMA
// =================================================================

// Objeto global GEO que encapsula configuración y funciones
if (typeof GEO === 'undefined') var GEO = {};

// ID DEL LIBRO Nase Control de Entradas y Salidas (Donde está la hoja "Centros")
const ID_LIBRO_BASE = "1PchIxXq617RRL556vHui4ImG7ms2irxiY3fPLIoqcQc"; 

GEO.CONFIG = {
  SHEET_RESPUESTAS: "Respuestas", // Nombre de la hoja principal (Libro actual)
  SHEET_CENTROS: "Centros",      // Nombre de la hoja de referencia (Libro Externo)
  
  // ✅ MAPEO DE COLUMNAS (Basado en estructura original solicitada por usuario)
  // Indices 0-based correspondientes a:
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
  // ⚠️ ATENCIÓN: Reemplaza '0f4d...' con tu clave real si aún no la has guardado en Properties
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
 * @description Correcciones lógicas para ciudades ambiguas.
 */
GEO.normalizarCiudadLocal = function(ciudad, lat, lng) {
  if (!ciudad) return "No encontrado";
  const l = ciudad.toString().toLowerCase();
  
  // Reglas específicas para ciudades con varios nombres
  if (l.includes("risaralda") || l.includes("eje") || l.includes("amco") || l.includes("perimetro urbano pereira") || l.includes("perimetro urbano pereira")) {
    return lat >= 4.825 ? "Dosquebradas" : "Pereira";
  }
  if (l.includes("valle") || l.includes("pac") || l.includes("perimetro urbano santiago de cali") || l.includes("cali")) return "Cali";
  if (l.includes("cartago")) return "Cartago";
  if (l.includes("bogot") || l.includes("cundinamarca")) return "Bogotá";
  if (l.includes("medellín") || l.includes("perimetro urbano medellín")) return "Medellín";
  
  // Elimina prefijos comunes si no hay regla específica
  return ciudad.replace(/^(Per[íi]metro Urbano|AMCO|Area Metropolitana Centro Occidente|.*,\s*)+/gi, '').split(',')[0].trim() || "No encontrado";
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
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
};

/**
 * @summary Busca dinámicamente el índice de una columna por nombre.
 * @description Fallback robusto para encontrar índices aunque el orden varíe ligeramente.
 */
GEO.findHeaderIndex = function(headers, names) {
  if (!headers || !headers.length) return -1;
  const lower = headers.map(h => (h || "").toString().trim().toLowerCase());
  for (const cand of names) {
    const idx = lower.indexOf(cand.toString().trim().toLowerCase());
    if (idx !== -1) return idx + 1;
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
  const r = UrlFetchApp.fetch(url, { headers: { 'User-Agent': 'NASE-Geocoder/1.0 (contact: soporte@nasecolombia.com.co)' } });
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
    const ciudad = GEO.normalizarCiudadLocal(components.city || components.town || '', lat, lng);
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

  // 1. Intentar Maps.co (Rápido y fiable)
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

  // 2. Si falla, intentar Nominatim (OpenStreetMap)
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
 
  // 3. Si fallan, intentar OpenCage (último recurso)
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
 * @description Lee la fila seleccionada, conecta al libro EXTERNO de centros,
 *              calcula la distancia y escribe los datos en el libro actual.
 */
function geocodificarFilaActiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRes = ss.getActiveSheet();
  
  if (hojaRes.getName() !== GEO.CONFIG.SHEET_RESPUESTAS) {
    SpreadsheetApp.getUi().alert("Por favor, ejecuta esta función desde la hoja 'Respuestas'.");
    return;
  }

  // ----------------------------------------------------------------------
  // 1. OBTENER DATOS DE LA FILA ACTIVA (Libro Actual)
  // ----------------------------------------------------------------------
  const fila = hojaRes.getActiveCell().getRow();
  
  if (!fila || fila < 2) { 
    SpreadsheetApp.getUi().alert("Selecciona una fila válida en la hoja Respuestas."); 
    return; 
  }

  const headers = hojaRes.getRange(1, 1, 1, hojaRes.getLastColumn()).getValues()[0] || [];
  const getCol = (name) => GEO.findHeaderIndex(headers, [name]);

  const filaVals = hojaRes.getRange(fila, 1, 1, hojaRes.getLastColumn()).getValues()[0];  
  const centro = (filaVals[getCol("centro") - 1] || "").toString().trim();
  const ciudadCentro = (filaVals[getCol("ciudad") - 1] || "").toString().trim();
  
  // Normalizar coordenadas (Manejo de comas separador de miles o decimales)
  const latEmp = GEO.normalizarCoord(filaVals[getCol("lat") - 1]); 
  const lngEmp = GEO.normalizarCoord(filaVals[getCol("lng") - 1]); 
  const direccion = (filaVals[getCol("direccion") - 1] || "").toString();

  // Validar datos mínimos
  if (isNaN(latEmp) || isNaN(lngEmp)) {
    SpreadsheetApp.getUi().alert("⚠️ Coordenadas inválidas o vacías en la fila seleccionada.");
    return;
  }
  if (!centro) {
    SpreadsheetApp.getUi().alert("⚠️ La fila seleccionada no tiene un Centro asignado en la columna 'Centro'.");
    return;
  }

  // ----------------------------------------------------------------------
  // 2. BUSCAR INFORMACIÓN DEL CENTRO EN EL LIBRO EXTERNO
  // ----------------------------------------------------------------------
  
  let latCentro = null, lngCentro = null, radio = 30, urlImagenCentro = "";
  
  try {
    // ABRIR LIBRO EXTERNO (Base Operativa)
    const ssExt = SpreadsheetApp.openById(ID_LIBRO_BASE);
    
    // Buscar hoja "Centros" en el libro externo
    let hojaCentros = ssExt.getSheetByName("Centros");
    if (!hojaCentros) {
       // Intento secundario por si el nombre difiere
       hojaCentros = ssExt.getSheetByName("BASE_CENTROS");
    }

    if (!hojaCentros) {
      SpreadsheetApp.getUi().alert("⚠️ No se encontró la hoja 'Centros' ni 'BASE_CENTROS' en el Libro Base Operativa.");
      return;
    }

    const dataCentros = hojaCentros.getDataRange().getValues();
    if (!dataCentros || dataCentros.length < 2) return;

    const headersCentros = dataCentros[0];
  
    // Función auxiliar para buscar en la hoja Centros
    const getColC = (name) => GEO.findHeaderIndex(headersCentros, [name]);

    // Iterar para encontrar el centro que coincida (Nombre + Ciudad)
    for (let i = 1; i < dataCentros.length; i++) {
      const rowC = dataCentros[i];
      const nombreC = (rowC[getColC("centro") - 1] || "").toString().trim();
      const ciudadC = (rowC[getColC("ciudad") - 1] || "").toString().trim();
     
      // Coincidencia insensible a mayúsculas
      if (nombreC.toUpperCase() === centro.toUpperCase() && ciudadC.toUpperCase() === ciudadCentro.toUpperCase()) {
        // Leer coordenadas del centro externo
        latCentro = GEO.normalizarCoord(rowC[getColC("lat ref") - 1]);
        lngCentro = GEO.normalizarCoord(rowC[getColC("lng ref") - 1]);
        radio = rowC[getColC("radio") - 1] ? Number(rowC[getColC("radio") - 1]) : 30;
        
        // Leer dirección e imagen si existen en el libro externo
        const idxDir = getColC("direccion");
        const idxImg = getColC("link_imagen");
        
        direccion = (idxDir > -1) ? (rowC[idxDir - 1] || "").toString() : "";
        
        if (idxImg > -1) {
           urlImagenCentro = (rowC[idxImg - 1] || "").toString().trim();
        }
        
        break;
      }
    }
  } catch (errExt) {
    SpreadsheetApp.getUi().alert("❌ Error leyendo datos externos del Centro: " + errExt.toString());
    return;
  }

  // Validar que se encontró el centro de referencia
  if (isNaN(latCentro) || isNaN(lngCentro)) {
    SpreadsheetApp.getUi().alert(`⚠️ No se encontró el centro "${centro}" en la ciudad "${ciudadCentro}" en el Libro Base.`);
    return;
  }

  // ----------------------------------------------------------------------
  // 3. CÁLCULO DE DISTANCIA Y ESTADO (Dentro/Fuera)
  // ----------------------------------------------------------------------
  const distancia = GEO.calcularDistancia(latCentro, lngCentro, latEmp, lngEmp); 
  const dentro = distancia <= radio;

  // ----------------------------------------------------------------------
  // 4. OBTENER GEOCODIFICACIÓN DE LA UBICACIÓN DEL EMPLEADO
  // ----------------------------------------------------------------------
  const resultadoGeo = GEO.reverseGeocodeInternal(latEmp, lngEmp);

  // ----------------------------------------------------------------------
  // 5. ESCRIBIR RESULTADOS EN LA HOJA "RESPUESTAS" (Libro Actual)
  // ----------------------------------------------------------------------
  
  // Preparar rangos para escritura rápida
  const rangoCiudad = hojaRes.getRange(fila, GEO.CONFIG.R_COL_CIUDAD_GEO + 1);
  const rangoDir = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DIR_GEO + 1);
  const rangoDentro = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DENTRO_CENTRO + 1);
  const rangoDistancia = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DISTANCIA + 1);
  const rangoAccuracy = hojaRes.getRange(fila, GEO.CONFIG.R_COL_ACCURACY + 1);

  rangoCiudad.setValue(resultadoGeo.ciudad);
  rangoDir.setValue(resultadoGeo.direccion);
  rangoAccuracy.setValue(resultadoGeo.fuente); // Guardar fuente de la geocodificación

  rangoDentro.setValue(dentro ? "Sí" : "No");
  rangoDistancia.setValue(Number(distancia.toFixed(2)));

  // Alerta final
  if (dentro) {
    SpreadsheetApp.getUi().alert(`✅ La ubicación está DENTRO del centro "${centro}".\n\n🏢 Ciudad: ${resultadoGeo.ciudad}\n📍 Dirección: ${resultadoGeo.direccion}\n📏 Distancia: ${distancia.toFixed(2)}m`);
  } else {
    SpreadsheetApp.getUi().alert(`📍 La ubicación está FUERA del centro "${centro}".\n\n🏢 Ciudad: ${resultadoGeo.ciudad}\n📍 Dirección: ${resultadoGeo.direccion}\n📏 Distancia: ${distancia.toFixed(2)}m\n\nSe ha registrado la geocodificación.`);
  }
}
