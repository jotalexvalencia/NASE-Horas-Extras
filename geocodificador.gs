// =================================================================
// 📁 geocodificador.gs – Módulo de Geocodificación y Distancia (NASE 2026)
// =================================================================
/**
 * @summary Módulo inteligente para geocodificación y control de asistencia por ubicación.
 * @description Administra la conversión de coordenadas GPS (Lat/Lng) a direcciones y ciudades.
 *              Calcula distancias (Haversine) entre el empleado y su centro asignado.
 *              Implementa un sistema de optimización: Si está dentro del radio, NO consume APIs.
 *              Gestiona cuotas diarias de la API de OpenCage (fallback).
 * 
 * @author NASE Team
 * @version 2.0 (Optimizado con OpenCage y Propiedades de Script)
 */

// =================================================================
// 1. CONFIGURACIÓN DEL SISTEMA
// =================================================================

// Objeto global GEO que encapsula configuración y funciones (Namespace Pattern)
if (typeof GEO === 'undefined') var GEO = {};

GEO.CONFIG = {
  SHEET_RESPUESTAS: "Respuestas", // Nombre de la hoja principal
  SHEET_CENTROS: "Centros",      // Nombre de la hoja de referencia de centros
  
  // Mapeo de columnas en la hoja 'Respuestas' (Índices 0-based para Arrays)
  R_COL_LAT: 6,            // Columna F: Latitud GPS
  R_COL_LNG: 7,            // Columna G: Longitud GPS
  R_COL_CIUDAD_GEO: 9,    // Columna I: Ciudad Geocodificada
  R_COL_DIR_GEO: 10,       // Columna J: Dirección Geocodificada
  R_COL_ACCURACY: 11,     // Columna K: Precisión GPS
  R_COL_DENTRO_CENTRO: 12, // Columna L: ¿Está dentro del centro?
  R_COL_DISTANCIA: 13,     // Columna M: Distancia al centro en metros
 
  // Configuración de OpenCage (API Pagada de respaldo)
  OPENCAGE_DAILY_LIMIT: 2500, // Límite de peticiones gratuitas diarias (aproximado para este script)
  OPENCAGE_API_PROP: 'OPENCAGE_API_KEY', // Nombre de la propiedad donde se guarda la API Key
  OPENCAGE_QUOTA_PROP: 'OPENCAGE_QUOTA',    // Propiedad para contar uso diario
  
  // Retrasos para APIs gratuitas (evitar bloqueos por Rate Limiting)
  REQUEST_DELAY_MS: 300
};

/**
 * @summary Guarda la API Key de OpenCage de forma segura.
 * @description Esta función debe ejecutarse UNA SOLA VEZ manualmente desde el editor.
 *              Guarda la clave en `PropertiesService` para que no esté expuesta en el código.
 */
function guardarApiKeyOpenCage() {
  // ⚠️ ATENCIÓN: Reemplaza 'TU_CLAVE_API_DE_OPENCAGE' con la clave real (ej: abc123...)
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
 * @description Detecta ciudades complejas como Pereira (donde una lat/long puede ser 
 *              Dosquebradas o Pereira) basándose en umbrales de latitud.
 * @param {String} ciudad - Nombre devuelto por la API.
 * @param {Number} lat - Latitud del usuario.
 * @param {Number} lng - Longitud del usuario.
 * @returns {String} Nombre de la ciudad corregido.
 */
GEO.normalizarCiudadLocal = function(ciudad, lat, lng) {
  if (!ciudad) return "No encontrado";
  const l = ciudad.toString().toLowerCase();
  
  // Reglas específicas para Área Metropolitana
  if (l.includes("risaralda") || l.includes("eje") || l.includes("amco") || l.includes("perimetro urbano pereira") || l.includes("perímetro urbano pereira")) {
    // Umbrales: Lat >= 4.825 suele ser Pereira, < 4.825 Dosquebradas
    return lat >= 4.825 ? "Dosquebradas" : "Pereira";
  }
  if (l.includes("valle") || l.includes("pac") || l.includes("perímetro urbano santiago de cali") || l.includes("cali")) return "Cali";
  if (l.includes("cartago")) return "Cartago";
  if (l.includes("bogot") || l.includes("cundinamarca")) return "Bogotá";
  if (l.includes("medellín") || l.includes("perímetro urbano medellín")) return "Medellín";
  
  // Fallback: Limpieza de texto básico
  return ciudad.replace(/^(Per[íi]metro urbano|AMCO|Area Metropolitana Centro Occidente|RAP Eje Cafetero|.*,\s*)+/gi, '').split(',')[0].trim() || "No encontrado";
};

/**
 * @summary Limpia la dirección eliminando código postal y país.
 * @description Quita elementos comunes en direcciones geocodificadas que no sirven para este sistema.
 */
GEO.limpiarDireccion = function(dir) {
  if (!dir) return "Sin dirección específica";
  return dir.split(',').map(p => p.trim()).filter(p => p && !/^\d{5,}$/i.test(p) && !/colombia|rap eje|risaralda|valle|cundinamarca|antioquia|66000|16000|05000|76000/i.test(p.toLowerCase())).join(', ').replace(/\s+/g, ' ').trim();
};

/**
 * @summary Calcula la distancia entre dos puntos GPS (Fórmula Haversine).
 * @description Utilizada para saber si el empleado está dentro del radio del centro.
 * @param {Number} lat1, lon1, lat2, lon2 - Coordenadas decimales.
 * @returns {Number} Distancia en metros.
 */
GEO.calcularDistancia = function(lat1, lon1, lat2, lon2) {
  const R = 6371000; // Radio de la Tierra en metros
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
 * @description Robusto contra cambios de orden en las hojas de cálculo.
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

/**
 * @summary Verifica si se ha excedido el límite diario de OpenCage.
 * @description Usa `PropertiesService` para persistir el contador de uso diario.
 *              Evita gasto innecesario si el límite gratuito se alcanza.
 * @returns {Boolean} `true` si se permite la llamada, `false` si se alcanzó el límite.
 */
GEO.checkAndUpdateOpenCageQuota = function() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date().toISOString().slice(0, 10); // Fecha actual YYYY-MM-DD
  
  let quotaData = JSON.parse(scriptProperties.getProperty(GEO.CONFIG.OPENCAGE_QUOTA_PROP) || '{"date": "", "count": 0}');

  // Si cambió el día, resetear contador
  if (quotaData.date !== today) {
    quotaData.date = today;
    quotaData.count = 0;
  }

  // Si superó el límite, denegar acceso
  if (quotaData.count >= GEO.CONFIG.OPENCAGE_DAILY_LIMIT) {
    Logger.log(`❌ LÍMITE DIARIO DE OPENCAGE ALCANZADO (${GEO.CONFIG.OPENCAGE_DAILY_LIMIT}).`);
    return false; 
  }

  // Si todo bien, incrementar y guardar
  quotaData.count++;
  scriptProperties.setProperty(GEO.CONFIG.OPENCAGE_QUOTA_PROP, JSON.stringify(quotaData));
  Logger.log(`✅ Llamada a OpenCage permitida. (${quotaData.count}/${GEO.CONFIG.OPENCAGE_DAILY_LIMIT}) hoy.`);
  return true; 
};

// =================================================================
// 4. FUENTES DE GEOCODIFICACIÓN (Wrappers de API)
// =================================================================

/**
 * @summary Obtiene dirección usando API de Maps.co (Gratuito).
 * @description Primera opción. Es gratuito y tiene buena cobertura en Colombia.
 */
GEO.getFromMapsCo = (lat, lng) => {
  const url = `https://geocode.maps.co/reverse?lat=${lat}&lon=${lng}&accept-language=es`;
  const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return JSON.parse(r.getContentText());
};

/**
 * @summary Obtiene dirección usando Nominatim OpenStreetMap (Gratuito).
 * @description Segunda opción. Código abierto. Se usa para verificar o complementar Maps.co.
 */
GEO.getFromNominatim = (lat, lng) => {
  const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${lat}&lon=${lng}&addressdetails=1&accept-language=es`;
  const r = UrlFetchApp.fetch(url, { headers: { 'User-Agent': 'NASE-Geocoder/1.0 (contact: analistaoperaciones@nasecolombia.com.co)' } });
  return JSON.parse(r.getContentText());
};

/**
 * @summary Obtiene dirección usando API de OpenCage (Pago/Cuota controlada).
 * @description Tercera opción (Fallback). Se usa solo si las APIs gratis fallan o si
 *              el empleado está "Fuera" y se requiere alta precisión.
 *              Controla cuotas para no exceder el límite gratuito diario.
 */
GEO.getFromOpenCage = (lat, lng) => {
  // Verificar si se puede usar OpenCage (cuotas)
  if (!GEO.checkAndUpdateOpenCageQuota()) {
    return { error: 'Límite diario de OpenCage alcanzado.' };
  }
  
  // Obtener API Key guardada en propiedades
  const apiKey = PropertiesService.getScriptProperties().getProperty(GEO.CONFIG.OPENCAGE_API_PROP);
  if (!apiKey) return { error: 'API Key de OpenCage no configurada.' };
  
  const url = `https://api.opencagedata.com/geocode/v1/json?q=${lat}+${lng}&key=${apiKey}&language=es&limit=1`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(response.getContentText());

  if (response.getResponseCode() === 200 && json.results && json.results.length > 0) {
    const result = json.results[0];
    const components = result.components;
    
    // Normalizar ciudad (Aplicar reglas de negocio locales)
    const ciudad = GEO.normalizarCiudadLocal(components.city || components.town || components.county || '', lat, lng);
    const direccion = GEO.limpiarDireccion(result.formatted);
    
    return { ciudad: ciudad, direccion: direccion, fuente: 'OpenCage' };
  }
  return null; // No encontró nada
};

// =================================================================
// 5. LÓGICA PRINCIPAL DE GEOCODIFICACIÓN
// =================================================================

/**
 * @summary Función interna principal que orquesta la geocodificación.
 * @description Implementa una cascada de intentos (Cascade Fallback).
 *              Intenta orden: Maps.co -> Nominatim -> OpenCage.
 *              Devuelve el primer resultado exitoso.
 * 
 * @param {Number} lat - Latitud GPS del usuario.
 * @param {Number} lng - Longitud GPS del usuario.
 * @returns {Object} { ciudad: String, direccion: String, fuente: String }
 */
GEO.reverseGeocodeInternal = function(lat, lng) {
  let result = null;
  let source = '';

  // ----------------------------------------------------------------
  // 1. Intentar con Maps.co (Rápido y Gratuito)
  // ----------------------------------------------------------------
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

  // ----------------------------------------------------------------
  // 2. Intentar con Nominatim si Maps.co falló
  // ----------------------------------------------------------------
  if (!result) {
    try {
      const j = GEO.getFromNominatim(lat, lng);
      if (j && j.display_name) {
        const a = j.address || {};
        const ciudad = GEO.normalizarCiudadLocal(a.city || a.town || a.county || '', lat, lng);
        const direccion = GEO.limpiarDireccion(j.display_name);
        result = { ciudad: ciudad, direccion: direccion };
        source = 'Nominatim';
        Utilities.sleep(1000); // Retraso cortés para ser amable con la API gratuita
      }
    } catch (e) { Logger.log(e); }
  }
 
  // ----------------------------------------------------------------
  // 3. Intentar con OpenCage como último recurso
  // ----------------------------------------------------------------
  if (!result) {
    const openCageResult = GEO.getFromOpenCage(lat, lng);
    if (openCageResult && !openCageResult.error) {
      result = { ciudad: openCageResult.ciudad, direccion: openCageResult.direccion };
      source = 'OpenCage';
    } else if (openCageResult && openCageResult.error) {
       return { ciudad: 'Error', direccion: openCageResult.error, fuente: 'OpenCage' };
    }
  }

  // ----------------------------------------------------------------
  // 4. Fallback por defecto si todo falla
  // ----------------------------------------------------------------
  if (!result) {
    result = { ciudad: 'No encontrado', direccion: 'Sin dirección específica' };
    source = 'Ninguna';
  }
 
  // Metadata de fuente para auditoría
  result.fuente = source;
  result.accuracy = 0; // La precisión no la calcularemos por ahora para simplificar
 
  // Si la fuente no es Nominatim (que es lento), ponemos un pequeño delay para no saturar
  if (!source.includes('Nominatim')) {
    Utilities.sleep(GEO.CONFIG.REQUEST_DELAY_MS);
  }
 
  return result;
};

// =================================================================
// 6. FUNCIÓN PÚBLICA PRINCIPAL (Trigger de Usuario)
// =================================================================

/**
 * @summary Función principal que se ejecuta desde el Menú o Fórmula.
 * @description Lee la fila activa (con coordenadas), busca el centro de referencia en la hoja
 *              'Centros', calcula la distancia y decide si geocodificar o no.
 * 
 * LÓGICA DE OPTIMIZACIÓN:
 * - Si la distancia es MENOR al radio del centro (Dentro):
 *   Escribe el nombre del centro y "Sí" en la hoja.
 *   NO consume APIs (Ahorro de costo).
 * - Si la distancia es MAYOR al radio (Fuera):
 *   Llama a `reverseGeocodeInternal` para consultar APIs externas.
 *   Escribe la ciudad y dirección halladas.
 *   Escribe "No" en la hoja.
 */
function geocodificarFilaActiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRes = ss.getActiveSheet();
  
  // Validar contexto
  if (hojaRes.getName() !== GEO.CONFIG.SHEET_RESPUESTAS) {
    SpreadsheetApp.getUi().alert("Por favor, ejecuta esta función desde la hoja 'Respuestas'.");
    return;
  }

  // ----------------------------------------------------------------
  // 1. OBTENER DATOS DE LA FILA ACTIVA
  // ----------------------------------------------------------------
  const fila = hojaRes.getActiveCell().getRow();
  
  // No ejecutar en el encabezado
  if (fila === 1) { 
    SpreadsheetApp.getUi().alert("Selecciona una fila con datos, no el encabezado."); 
    return; 
  }

  const headers = hojaRes.getRange(1, 1, 1, hojaRes.getLastColumn()).getValues()[0];
  
  // Función auxiliar para obtener índice de columna de forma dinámica
  const getCol = (name) => GEO.findHeaderIndex(headers, [name]);

  const lat = GEO.normalizarCoord(hojaRes.getRange(fila, getCol("lat")).getValue());
  const lng = GEO.normalizarCoord(hojaRes.getRange(fila, getCol("lng")).getValue());
  const nombreCentro = (hojaRes.getRange(fila, getCol("centro")).getValue() || "").toString().trim();
  const ciudadCentro = (hojaRes.getRange(fila, getCol("ciudad")).getValue() || "").toString().trim();

  // Validar datos mínimos
  if (lat == null || lng == null || !nombreCentro) {
    SpreadsheetApp.getUi().alert("Faltan datos clave en la fila: Latitud, Longitud o Centro.");
    return;
  }

  // ----------------------------------------------------------------
  // 2. BUSCAR EL CENTRO EN LA HOJA 'Centros' PARA OBTENER REFERENCIA Y RADIO
  // ----------------------------------------------------------------
  const hojaCentros = ss.getSheetByName(GEO.CONFIG.SHEET_CENTROS);
  if (!hojaCentros) { SpreadsheetApp.getUi().alert("La hoja 'Centros' no existe."); return; } 
  
  const dataCentros = hojaCentros.getDataRange().getValues();
  const headersCentros = dataCentros[0];
  const getColC = (name) => GEO.findHeaderIndex(headersCentros, [name]);

  let centroRef = null;
  
  // Buscar la fila que coincida con el Nombre del Centro y la Ciudad
  for (let i = 1; i < dataCentros.length; i++) {
    const rowC = dataCentros[i];
    const nombreC = (rowC[getColC("centro") - 1] || "").toString().trim();
    const ciudadC = (rowC[getColC("ciudad") - 1] || "").toString().trim();
   
    if (nombreC.toUpperCase() === nombreCentro.toUpperCase() && ciudadC.toUpperCase() === ciudadCentro.toUpperCase()) {
      centroRef = {
        lat: GEO.normalizarCoord(rowC[getColC("lat ref") - 1]),
        lng: GEO.normalizarCoord(rowC[getColC("lng ref") - 1]),
        radio: Number(rowC[getColC("radio") - 1]) || 30, // Radio por defecto 30m
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

  // ----------------------------------------------------------------
  // 3. CALCULAR DISTANCIA Y DETERMINAR SI ESTÁ DENTRO O FUERA
  // ----------------------------------------------------------------
  const distancia = GEO.calcularDistancia(lat, lng, centroRef.lat, centroRef.lng);
  const estaDentro = distancia <= centroRef.radio;

  // ----------------------------------------------------------------
  // 4. ESCRIBIR RESULTADOS EN LA HOJA
  // ----------------------------------------------------------------
  const rangoCiudad = hojaRes.getRange(fila, GEO.CONFIG.R_COL_CIUDAD_GEO);
  const rangoDir = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DIR_GEO);
  const rangoDentro = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DENTRO_CENTRO);
  const rangoDistancia = hojaRes.getRange(fila, GEO.CONFIG.R_COL_DISTANCIA);

  if (estaDentro) {
    // ----------------------------------------------------------------
    // CASO 1: ESTÁ DENTRO -> NO USAR APIs (Optimización de costos)
    // ----------------------------------------------------------------
    rangoCiudad.setValue(centroRef.ciudad);
    rangoDir.setValue(centroRef.nombre);
    rangoDentro.setValue("Sí");
    rangoDistancia.setValue(Number(distancia.toFixed(2))); // 2 decimales para metros
    
    SpreadsheetApp.getUi().alert(`✅ La ubicación está DENTRO del centro "${centroRef.nombre}".\nNo se consumieron cuotas de geocodificación.`);
  } else {
    // ----------------------------------------------------------------
    // CASO 2: ESTÁ FUERA -> USAR APIs (Consumo de cuotas)
    // ----------------------------------------------------------------
    rangoDentro.setValue("No");
    rangoDistancia.setValue(Number(distancia.toFixed(2)));
    
    SpreadsheetApp.getUi().alert('📍 La ubicación está FUERA del centro. Geocodificando con APIs externas... Por favor, espere.');
    
    // Llamada principal al motor de geocodificación
    const resultado = GEO.reverseGeocodeInternal(lat, lng);
   
    // Escribir resultados
    rangoCiudad.setValue(resultado.ciudad);
    rangoDir.setValue(resultado.direccion);
    hojaRes.getRange(fila, GEO.CONFIG.R_COL_ACCURACY).setValue(resultado.fuente);
   
    SpreadsheetApp.getUi().alert(`✔ Geocodificación completada.\n\nCiudad: ${resultado.ciudad}\nDirección: ${resultado.direccion}\nFuente: ${resultado.fuente}`);
  }
}
