// =================================================================
// 📁 geocodificador.gs – Módulo de Geocodificación (NASE 2026 - Horas Extras)
// =================================================================
/**
 * @summary Módulo inteligente para geocodificación y control de asistencia.
 * @description Lee la hoja "Centros" desde un libro EXTERNO.
 * @version 2.3 (Corregido - Errores de variables y fórmula Haversine)
 */

// =================================================================
// 1. CONFIGURACIÓN DEL SISTEMA
// =================================================================

if (typeof GEO === 'undefined') var GEO = {};

// ID DEL LIBRO EXTERNO donde está la hoja "Centros"
const ID_LIBRO_CENTROS_EXTERNO = "1PchIxXq617RRL556vHui4ImG7ms2irxiY3fPLIoqcQc"; 

GEO.CONFIG = {
  SHEET_RESPUESTAS: "Respuestas",
  SHEET_CENTROS: "Centros",
  
  // ÍNDICES 1-BASED (para usar directamente con getRange)
  // Estructura: A=Cédula, B=Centro, C=Ciudad, D=Lat, E=Lng, F=Acepto, 
  //             G=Ciudad_Geo, H=Barrio/Dir, I=Accuracy, J=Dentro, K=Distancia
  COL: {
    CEDULA: 1,
    CENTRO: 2,
    CIUDAD: 3,
    LAT: 4,
    LNG: 5,
    ACEPTO: 6,
    CIUDAD_GEO: 7,
    DIR_GEO: 8,
    ACCURACY: 9,
    DENTRO: 10,
    DISTANCIA: 11
  },
  
  OPENCAGE_DAILY_LIMIT: 2500,
  OPENCAGE_API_PROP: 'OPENCAGE_API_KEY',
  OPENCAGE_QUOTA_PROP: 'OPENCAGE_QUOTA',
  REQUEST_DELAY_MS: 300
};

/**
 * @summary Guarda la API Key de OpenCage de forma segura.
 */
function guardarApiKeyOpenCage() {
  var apiKey = '0f4d42a072704ffc8ad51d03a21fcea0'; 
  PropertiesService.getScriptProperties().setProperty(GEO.CONFIG.OPENCAGE_API_PROP, apiKey);
  Logger.log('✅ API Key de OpenCage guardada.');
  SpreadsheetApp.getUi().alert('✅ API Key guardada correctamente.');
}

// =================================================================
// 2. FUNCIONES AUXILIARES
// =================================================================

GEO.normalizarCoord = function(v) {
  if (v === null || v === undefined || v === '') return null;
  var s = String(v).replace(/,/g, '.').trim();
  var num = parseFloat(s);
  return isNaN(num) ? null : num;
};

GEO.normalizarCiudadLocal = function(ciudad, lat, lng) {
  if (!ciudad) return "No encontrado";
  var l = ciudad.toString().toLowerCase();
  
  if (l.indexOf("risaralda") !== -1 || l.indexOf("eje") !== -1 || l.indexOf("amco") !== -1 || l.indexOf("pereira") !== -1) {
    return lat >= 4.825 ? "Dosquebradas" : "Pereira";
  }
  if (l.indexOf("valle") !== -1 || l.indexOf("cali") !== -1) return "Cali";
  if (l.indexOf("cartago") !== -1) return "Cartago";
  if (l.indexOf("bogot") !== -1 || l.indexOf("cundinamarca") !== -1) return "Bogotá";
  if (l.indexOf("medell") !== -1 || l.indexOf("antioquia") !== -1) return "Medellín";
  if (l.indexOf("cartagena") !== -1 || l.indexOf("bolivar") !== -1) return "Cartagena";
  if (l.indexOf("ibagu") !== -1 || l.indexOf("tolima") !== -1) return "Ibagué";
  if (l.indexOf("neiva") !== -1 || l.indexOf("huila") !== -1) return "Neiva";
  if (l.indexOf("armenia") !== -1 || l.indexOf("quind") !== -1) return "Armenia";
  
  return ciudad.split(',')[0].trim() || "No encontrado";
};

GEO.limpiarDireccion = function(dir) {
  if (!dir) return "Sin dirección específica";
  return dir.split(',')
    .map(function(p) { return p.trim(); })
    .filter(function(p) { 
      return p && !/^\d{5,}$/i.test(p) && 
             !/colombia|rap eje|risaralda|valle|cundinamarca|antioquia/i.test(p.toLowerCase());
    })
    .join(', ')
    .replace(/\s+/g, ' ')
    .trim();
};

/**
 * @summary Calcula distancia Haversine (CORREGIDO)
 */
GEO.calcularDistancia = function(lat1, lon1, lat2, lon2) {
  var R = 6371000; // Radio de la Tierra en metros
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = Math.sin(dLat/2) * Math.sin(dLat/2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLon/2) * Math.sin(dLon/2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));  // ✅ CORREGIDO
  return R * c;
};

/**
 * @summary Busca índice de columna por nombre.
 */
GEO.findHeaderIndex = function(headers, names) {
  if (!headers || !headers.length) return -1;
  for (var i = 0; i < names.length; i++) {
    var name = names[i].toString().toLowerCase().trim();
    for (var j = 0; j < headers.length; j++) {
      if ((headers[j] || "").toString().toLowerCase().trim() === name) {
        return j + 1; // 1-based
      }
    }
  }
  return -1;
};

// =================================================================
// 3. CONTROL DE CUOTAS OPENCAGE
// =================================================================

GEO.checkAndUpdateOpenCageQuota = function() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var today = new Date().toISOString().slice(0, 10);
  
  var quotaStr = scriptProperties.getProperty(GEO.CONFIG.OPENCAGE_QUOTA_PROP) || '{"date": "", "count": 0}';
  var quotaData = JSON.parse(quotaStr);

  if (quotaData.date !== today) {
    quotaData.date = today;
    quotaData.count = 0;
  }

  if (quotaData.count >= GEO.CONFIG.OPENCAGE_DAILY_LIMIT) {
    Logger.log('❌ Límite diario de OpenCage alcanzado.');
    return false; 
  }

  quotaData.count++;
  scriptProperties.setProperty(GEO.CONFIG.OPENCAGE_QUOTA_PROP, JSON.stringify(quotaData));
  return true; 
};

// =================================================================
// 4. FUENTES DE GEOCODIFICACIÓN
// =================================================================

GEO.getFromMapsCo = function(lat, lng) {
  var url = 'https://geocode.maps.co/reverse?lat=' + lat + '&lon=' + lng + '&accept-language=es';
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return JSON.parse(response.getContentText());
};

GEO.getFromNominatim = function(lat, lng) {
  var url = 'https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=' + lat + 
            '&lon=' + lng + '&addressdetails=1&accept-language=es';
  var response = UrlFetchApp.fetch(url, { 
    headers: { 'User-Agent': 'NASE-Geocoder/1.0' },
    muteHttpExceptions: true 
  });
  return JSON.parse(response.getContentText());
};

GEO.getFromOpenCage = function(lat, lng) {
  if (!GEO.checkAndUpdateOpenCageQuota()) {
    return { error: 'Límite diario alcanzado.' };
  }
  
  var apiKey = PropertiesService.getScriptProperties().getProperty(GEO.CONFIG.OPENCAGE_API_PROP);
  if (!apiKey) return { error: 'API Key no configurada.' };
  
  var url = 'https://api.opencagedata.com/geocode/v1/json?q=' + lat + '+' + lng + 
            '&key=' + apiKey + '&language=es&limit=1';
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(response.getContentText());

  if (response.getResponseCode() === 200 && json.results && json.results.length > 0) {
    var result = json.results[0];
    var components = result.components;
    var ciudad = GEO.normalizarCiudadLocal(components.city || components.town || '', lat, lng);
    var direccion = GEO.limpiarDireccion(result.formatted);
    return { ciudad: ciudad, direccion: direccion, fuente: 'OpenCage' };
  }
  return null;
};

// =================================================================
// 5. LÓGICA PRINCIPAL DE GEOCODIFICACIÓN
// =================================================================

GEO.reverseGeocodeInternal = function(lat, lng) {
  var result = null;
  var source = '';

  // 1. Maps.co
  try {
    var j = GEO.getFromMapsCo(lat, lng);
    if (j && j.display_name) {
      var a = j.address || {};
      result = { 
        ciudad: GEO.normalizarCiudadLocal(a.city || a.town || a.county || '', lat, lng),
        direccion: GEO.limpiarDireccion(j.display_name)
      };
      source = 'Maps.co';
    }
  } catch (e) { Logger.log('Maps.co error: ' + e); }

  // 2. Nominatim
  if (!result) {
    try {
      var j2 = GEO.getFromNominatim(lat, lng);
      if (j2 && j2.display_name) {
        var a2 = j2.address || {};
        result = { 
          ciudad: GEO.normalizarCiudadLocal(a2.city || a2.town || a2.county || '', lat, lng),
          direccion: GEO.limpiarDireccion(j2.display_name)
        };
        source = 'Nominatim';
        Utilities.sleep(1000);
      }
    } catch (e) { Logger.log('Nominatim error: ' + e); }
  }
 
  // 3. OpenCage
  if (!result) {
    var openCageResult = GEO.getFromOpenCage(lat, lng);
    if (openCageResult && !openCageResult.error) {
      result = { ciudad: openCageResult.ciudad, direccion: openCageResult.direccion };
      source = 'OpenCage';
    }
  }

  // Fallback
  if (!result) {
    result = { ciudad: 'No encontrado', direccion: 'Sin dirección' };
    source = 'Ninguna';
  }
 
  result.fuente = source;
  
  if (source !== 'Nominatim') {
    Utilities.sleep(GEO.CONFIG.REQUEST_DELAY_MS);
  }
 
  return result;
};

// =================================================================
// 6. FUNCIÓN PRINCIPAL: GEOCODIFICAR FILA ACTIVA
// =================================================================

function geocodificarFilaActiva() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaRes = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // Validar hoja
  if (hojaRes.getName() !== GEO.CONFIG.SHEET_RESPUESTAS) {
    ui.alert("⚠️ Ejecuta desde la hoja '" + GEO.CONFIG.SHEET_RESPUESTAS + "'.");
    return;
  }

  var fila = hojaRes.getActiveCell().getRow();
  if (fila < 2) { 
    ui.alert("⚠️ Selecciona una fila con datos (no el encabezado)."); 
    return; 
  }

  // Leer datos de la fila usando constantes de columna
  var COL = GEO.CONFIG.COL;
  var lat = GEO.normalizarCoord(hojaRes.getRange(fila, COL.LAT).getValue());
  var lng = GEO.normalizarCoord(hojaRes.getRange(fila, COL.LNG).getValue());
  var nombreCentro = String(hojaRes.getRange(fila, COL.CENTRO).getValue() || '').trim();
  var ciudadCentro = String(hojaRes.getRange(fila, COL.CIUDAD).getValue() || '').trim();

  // Validar coordenadas
  if (lat === null || lng === null) {
    ui.alert("❌ Coordenadas inválidas en la fila " + fila);
    return;
  }
  if (!nombreCentro) {
    ui.alert("❌ Falta el nombre del Centro en la fila " + fila);
    return;
  }

  // ----------------------------------------------------------------
  // BUSCAR CENTRO EN LIBRO EXTERNO
  // ----------------------------------------------------------------
  var centroRef = null;
  
  try {
    var ssExterno = SpreadsheetApp.openById(ID_LIBRO_CENTROS_EXTERNO);
    var hojaCentros = ssExterno.getSheetByName("Centros");
    if (!hojaCentros) hojaCentros = ssExterno.getSheetByName("BASE_CENTROS");
    
    if (!hojaCentros) {
      ui.alert("❌ No se encontró la hoja 'Centros' en el libro externo.");
      return;
    }
    
    var dataCentros = hojaCentros.getDataRange().getValues();
    var headersCentros = dataCentros[0].map(function(h) { 
      return String(h).toUpperCase().trim(); 
    });
    
    // Buscar índices
    var idxCentroNombre = headersCentros.indexOf("CENTRO");
    var idxCentroCiudad = headersCentros.indexOf("CIUDAD");
    var idxCentroLat = -1, idxCentroLng = -1, idxCentroRadio = -1;
    
    for (var i = 0; i < headersCentros.length; i++) {
      if (headersCentros[i] === "LAT REF" || headersCentros[i] === "LAT") idxCentroLat = i;
      if (headersCentros[i] === "LNG REF" || headersCentros[i] === "LNG") idxCentroLng = i;
      if (headersCentros[i] === "RADIO") idxCentroRadio = i;
    }
    
    Logger.log("Headers Centros: " + headersCentros.join(" | "));
    Logger.log("Índices: Centro=" + idxCentroNombre + " Ciudad=" + idxCentroCiudad + 
               " Lat=" + idxCentroLat + " Lng=" + idxCentroLng);
    
    if (idxCentroNombre === -1 || idxCentroCiudad === -1 || idxCentroLat === -1 || idxCentroLng === -1) {
      ui.alert("❌ Faltan columnas en la hoja Centros del libro externo.\n" +
               "Se requiere: Ciudad, Centro, LAT REF, LNG REF");
      return;
    }
    
    // Buscar el centro específico
    for (var i = 1; i < dataCentros.length; i++) {
      var rowC = dataCentros[i];
      var nombreC = String(rowC[idxCentroNombre] || '').trim().toUpperCase();
      var ciudadC = String(rowC[idxCentroCiudad] || '').trim().toUpperCase();
      
      if (nombreC === nombreCentro.toUpperCase() && ciudadC === ciudadCentro.toUpperCase()) {
        centroRef = {
          lat: GEO.normalizarCoord(rowC[idxCentroLat]),
          lng: GEO.normalizarCoord(rowC[idxCentroLng]),
          radio: idxCentroRadio !== -1 ? (parseInt(rowC[idxCentroRadio]) || 30) : 30,
          nombre: String(rowC[idxCentroNombre]).trim(),
          ciudad: String(rowC[idxCentroCiudad]).trim()
        };
        break;
      }
    }
    
  } catch (e) {
    ui.alert("❌ Error accediendo al libro externo:\n" + e.toString());
    return;
  }

  if (!centroRef) {
    ui.alert('❌ Centro no encontrado:\n"' + nombreCentro + '" en "' + ciudadCentro + '"');
    return;
  }
  
  if (centroRef.lat === null || centroRef.lng === null) {
    ui.alert('❌ El centro "' + centroRef.nombre + '" no tiene coordenadas válidas.');
    return;
  }

  // Calcular distancia
  var distancia = GEO.calcularDistancia(lat, lng, centroRef.lat, centroRef.lng);
  var estaDentro = distancia <= centroRef.radio;

  // Geocodificar ubicación del empleado
  var resultado = GEO.reverseGeocodeInternal(lat, lng);

  // Escribir resultados
  hojaRes.getRange(fila, COL.CIUDAD_GEO).setValue(resultado.ciudad);
  hojaRes.getRange(fila, COL.DIR_GEO).setValue(resultado.direccion);
  hojaRes.getRange(fila, COL.ACCURACY).setValue(resultado.fuente);
  hojaRes.getRange(fila, COL.DENTRO).setValue(estaDentro ? "Sí" : "No");
  hojaRes.getRange(fila, COL.DISTANCIA).setValue(Math.round(distancia));

  // Mostrar resultado
  if (estaDentro) {
    ui.alert('✅ DENTRO del centro "' + centroRef.nombre + '"\n\n' +
             '📍 Ciudad: ' + resultado.ciudad + '\n' +
             '🏠 Dirección: ' + resultado.direccion + '\n' +
             '📏 Distancia: ' + Math.round(distancia) + 'm (Radio: ' + centroRef.radio + 'm)\n' +
             '📡 Fuente: ' + resultado.fuente);
  } else {
    ui.alert('⚠️ FUERA del centro "' + centroRef.nombre + '"\n\n' +
             '📍 Ciudad: ' + resultado.ciudad + '\n' +
             '🏠 Dirección: ' + resultado.direccion + '\n' +
             '📏 Distancia: ' + Math.round(distancia) + 'm (Radio: ' + centroRef.radio + 'm)\n' +
             '📡 Fuente: ' + resultado.fuente);
  }
}

// =================================================================
// 7. FUNCIÓN DE PRUEBA
// =================================================================

function testGeocodificador() {
  Logger.log("=== PRUEBA GEOCODIFICADOR ===");
  
  // Probar acceso al libro externo
  try {
    var ssExt = SpreadsheetApp.openById(ID_LIBRO_CENTROS_EXTERNO);
    Logger.log("✅ Libro externo accesible: " + ssExt.getName());
    
    var hojaCentros = ssExt.getSheetByName("Centros");
    if (hojaCentros) {
      Logger.log("✅ Hoja 'Centros' encontrada con " + (hojaCentros.getLastRow() - 1) + " registros");
      
      var headers = hojaCentros.getRange(1, 1, 1, hojaCentros.getLastColumn()).getValues()[0];
      Logger.log("📊 Headers: " + headers.join(" | "));
    } else {
      Logger.log("❌ Hoja 'Centros' no encontrada");
    }
  } catch (e) {
    Logger.log("❌ Error: " + e.toString());
  }
  
  // Probar geocodificación
  var lat = 4.710989;
  var lng = -74.072092;
  Logger.log("\n📍 Probando geocodificación: " + lat + ", " + lng);
  
  var resultado = GEO.reverseGeocodeInternal(lat, lng);
  Logger.log("Ciudad: " + resultado.ciudad);
  Logger.log("Dirección: " + resultado.direccion);
  Logger.log("Fuente: " + resultado.fuente);
}
