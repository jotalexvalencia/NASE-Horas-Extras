// ======================================================================
// 📍 MapaComparativo.gs – Visualización (NASE 2026 - Horas Extras)
// ----------------------------------------------------------------------
/**
 * @summary Módulo de visualización geoespacial comparativa.
 * @description Genera un mapa interactivo (Leaflet.js) que compara
 *              la ubicación de un empleado contra su centro de trabajo asignado.
 *
 * @features
 *   - 🔗 **Lectura Externa:** Lee la hoja 'Centros' del Libro Base Operativa.
 *   - 🗺️ Mapa Leaflet: Capas de "Callejero" y "Satélite".
 *   - 🏢 Centro: Marcador del centro de trabajo con popup (Dirección e Imagen).
 *   - 📍 Empleado: Marcador de la ubicación del empleado (Lat/Lng).
 *   - 📏 Distancia: Cálculo de la distancia en metros (Haversine).
 *   - 🎨 Círculo: Dibuja el radio de asistencia del centro (mínimo 80m visual).
 *   - 🚦 Semáforo: Azul si está DENTRO, Rojo si está FUERA del radio.
 *   - 🧭 Línea: Conexión visual entre el centro y el empleado.
 *
 * @author NASE Team
 * @version 1.4 (Corrección Externa + Coordenadas)
 */

// ======================================================================
// 1. FUNCIÓN PRINCIPAL: Generar Mapa Comparativo
// ======================================================================

/**
 * @summary Genera y muestra el mapa comparativo en una ventana modal.
 * @description Lee la fila seleccionada, busca el centro en el LIBRO EXTERNO,
 *              calcula la distancia y construye el código HTML para Leaflet.
 */
function mostrarMapaComparativo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResp = ss.getActiveSheet();
  const fila = hojaResp.getActiveCell().getRow();

  // Validar contexto: Seleccionar una fila con datos (no fila 1)
  if (!fila || fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una fila válida en la hoja Respuestas antes de usar 'Ver mapa comparativo'.");
    return;
  }

  // ================================================================
  // 1️⃣ OBTENER DATOS DEL EMPLEADO (Fila Activa)
  // ================================================================
  const headersResp = hojaResp.getRange(1, 1, 1, hojaResp.getLastColumn()).getValues()[0] || [];  
  const findIdx = (hdrs, candidates) => {
    const low = hdrs.map(h => (h || "").toString().trim().toLowerCase());
    for (const cand of candidates) {
      const idx = low.indexOf(cand.toString().trim().toLowerCase());
      if (idx !== -1) return idx + 1;
    }
    return -1;
  };

  const idxCentro = findIdx(headersResp, ["centro", "centros", "centro de trabajo"]);
  const idxCiudadCentro = findIdx(headersResp, ["ciudad", "city"]);
  const idxLat = findIdx(headersResp, ["lat", "latitude", "latitud"]);
  const idxLng = findIdx(headersResp, ["lng", "lon", "long", "longitude", "longitud"]);
  const idxDir = findIdx(headersResp, ["barrio / dirección", "direccion", "dirección", "address"]);

  const filaVals = hojaResp.getRange(fila, 1, 1, hojaResp.getLastColumn()).getValues()[0];  
  const centro = (filaVals[idxCentro - 1] || "").toString().trim();
  const ciudadCentro = (filaVals[idxCiudadCentro - 1] || "").toString().trim();
  
  // Parsear coordenadas de la hoja Respuestas (manejando comas)
  const latEmp = _parseCoord(filaVals[idxLat - 1]); 
  const lngEmp = _parseCoord(filaVals[idxLng - 1]); 
  const direccion = (filaVals[idxDir - 1] || "").toString();

  // Validar datos mínimos
  if (isNaN(latEmp) || isNaN(lngEmp)) {
    SpreadsheetApp.getUi().alert("⚠️ Coordenadas inválidas o vacías en la fila seleccionada. Verifica Lat/Lng en la hoja Respuestas. (Nota: Usa punto '.' en lugar de coma ',')");
    return;
  }
  if (!centro) {
    SpreadsheetApp.getUi().alert("⚠️ La fila seleccionada no tiene un Centro asignado.");
    return;
  }

  // ================================================================
  // 2️⃣ BUSCAR INFORMACIÓN DEL CENTRO DE REFERENCIA (LIBRO EXTERNO)
  // ================================================================
  // ID DEL LIBRO BASE OPERATIVA (El mismo que se usa para Centros)
  const ID_LIBRO_BASE_OPERATIVA = "1PchIxXq617RRL556vHui4ImG7ms2irxiY3fPLIoqcQc"; 
  
  let latCentro = null, lngCentro = null, radio = 30, urlImagenCentro = "", direccionCentro = "";
  
  try {
    const ssExt = SpreadsheetApp.openById(ID_LIBRO_BASE_OPERATIVA);
    
    // Buscar la hoja. Probamos "Centros" primero.
    let hojaCentros = ssExt.getSheetByName("Centros");
    if (!hojaCentros) {
       // Si no existe, probar con "BASE_CENTROS"
       hojaCentros = ssExt.getSheetByName("BASE_CENTROS");
    }

    if (!hojaCentros) {
      SpreadsheetApp.getUi().alert("⚠️ No se encontró la hoja 'Centros' ni 'BASE_CENTROS' en el Libro Base Operativa.");
      return;
    }

    const dataCentros = hojaCentros.getDataRange().getValues();
    if (!dataCentros || dataCentros.length < 2) return;

    const headersCentros = dataCentros[0];
  
    // Función auxiliar para buscar índice en hoja Centros
    const getColC = (name) => _findHeaderIndex(headersCentros, [name]);

    // Iterar para encontrar el centro (Comparando Nombre + Ciudad)
    for (let i = 1; i < dataCentros.length; i++) {
      const rowC = dataCentros[i];
      const nombreC = (rowC[getColC("centro") - 1] || "").toString().trim();
      const ciudadC = (rowC[getColC("ciudad") - 1] || "").toString().trim();
     
      if (nombreC.toUpperCase() === centro.toUpperCase() && ciudadC.toUpperCase() === ciudadCentro.toUpperCase()) {
        // Leer Coordenadas y Radio (manejando comas en lectura externa también)
        latCentro = _parseCoord(rowC[getColC("lat ref") - 1]);
        lngCentro = _parseCoord(rowC[getColC("lng ref") - 1]);
        radio = _parseCoord(rowC[getColC("radio") - 1]); // Radio como número
        direccionCentro = (rowC[getColC("direccion") - 1] || "").toString();
        
        // Buscar columna de imagen
        const idxImg = getColC("link_imagen");
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

  // Validar que se encontró el centro de referencia con datos numéricos
  if (isNaN(latCentro) || isNaN(lngCentro)) {
    SpreadsheetApp.getUi().alert(`⚠️ No se pudieron leer coordenadas válidas para el centro "${centro}" en el Libro Base.`);
    return;
  }
  if (isNaN(radio) || radio <= 0) radio = 30; // Radio por defecto si está vacío o 0

  // ================================================================
  // 3️⃣ CÁLCULO DE DISTANCIA Y ESTADO (Dentro/Fuera)
  // ================================================================
  const distancia = _distMetros(latCentro, lngCentro, latEmp, lngEmp); 
  const dentro = distancia <= radio;

  // ================================================================
  // 4️⃣ GENERAR HTML DEL MAPA (Leaflet.js + Manejo de Errores)
  // ================================================================
  const leafletCSS = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
  const leafletJS = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";

  // Preparar variables JS inyectadas con valores por defecto para evitar syntax errors
  const safeLatCentro = isNaN(latCentro) ? 4.57 : latCentro;
  const safeLngCentro = isNaN(lngCentro) ? -74.29 : lngCentro;
  const safeLatEmp = isNaN(latEmp) ? 4.57 : latEmp;
  const safeLngEmp = isNaN(lngEmp) ? -74.29 : lngEmp;
  const safeRadio = isNaN(radio) ? 30 : radio;

  const html = `
  <!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <link rel="stylesheet" href="${leafletCSS}" />
    <style>
      body {margin:0;padding:10px;font-family:Arial,sans-serif}
      #map {height:720px;width:100%;border-radius:8px}
      .popup-img {width:100%;max-width:300px;border-radius:5px;margin-top:5px}
      .legend {background:#fff;padding:6px;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.12)}
      .loading-text {text-align:center; padding:20px; font-size:18px; color:#666;}
    </style>
  </head>
  <body>
    <div class="info">🏢 Centro: ${centro}</div>
    <div id="map"></div>
    <script src="${leafletJS}"></script>
    <script>
      // Esperar a que cargue Leaflet para evitar "L is not defined"
      document.addEventListener('DOMContentLoaded', function() {
        try {
          const latC = ${safeLatCentro};
          const latE = ${safeLatEmp};
          const lngC = ${safeLngCentro};
          const lngE = ${safeLngEmp};
          const rd = ${safeRadio};

          // Validar coordenadas antes de crear mapa (Seguridad extra)
          if (typeof latC !== 'number' || typeof latE !== 'number' || typeof lngC !== 'number' || typeof lngE !== 'number') {
             throw new Error("Coordenadas no son números válidos.");
          }

          // Inicializar mapa: Centrado en el punto medio entre Centro y Empleado
          const map = L.map('map').setView([ ((latC + latE) / 2), ((lngC + lngE) / 2) ], 15);
          
          // Capa 1: Callejero (OpenStreetMap)
          const calle = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap'}).addTo(map);
          
          // Capa 2: Satélite (Esri World Imagery)
          const sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{attribution:'© Esri'});
          
          // Control de capas (Switcher)
          L.control.layers({"🗺️ Callejero":calle,"🌍 Satélite":sat}).addTo(map);

          // -------------------------------------------------------------
          // MARCADOR DEL CENTRO DE TRABAJO
          // -------------------------------------------------------------
          const centroMarker = L.marker([latC, lngC]).addTo(map);
          
          let htmlCentro = "<b>🏢 Centro de Trabajo</b><br>${centro}<br><b>Ciudad:</b> ${ciudadCentro}<br><b>Dirección:</b> ${direccionCentro}<br><b>Radio:</b> " + rd + " m";
          const imgUrl = "${urlImagenCentro}";
          if (imgUrl && imgUrl.length > 5) htmlCentro += "<br><img src='" + imgUrl + "' class='popup-img'>";
          
          centroMarker.bindPopup(htmlCentro);

          // -------------------------------------------------------------
          // CÍRCULO DE RADIO (Visualización de Área Permitida)
          // -------------------------------------------------------------
          L.circle([latC, lngC],{
            color:'#2e7d32',fillColor:'#2e7d32',fillOpacity:0.12,
            radius:Math.max(rd, 80) // Mínimo visual de 80m
          }).addTo(map);

          // -------------------------------------------------------------
          // MARCADOR DEL EMPLEADO
          // -------------------------------------------------------------
          const color = ${dentro} ? '#1976d2' : '#d32f2f'; // Azul = Dentro, Rojo = Fuera
          
          const icono = L.divIcon({
            html: \`<svg height="36" width="36"><circle cx="18" cy="18" r="16" fill="\${color}" stroke="white" stroke-width="2"/></svg>\`,
            iconSize:[36,36],iconAnchor:[18,36],popupAnchor:[0,-36]
          });
          
          const emp = L.marker([latE, lngE], {icon:icono}).addTo(map);
          
          const estado = ${dentro} ? '✅ DENTRO del radio' : '❌ FUERA del radio';
          const distStr = (${distancia}).toFixed(1);
          
          emp.bindPopup("<b>📍 Registro Empleado</b><br>${centro}<br>${direccion}<br><b>Lat:</b> " + latE.toFixed(6) + "<br><b>Lng:</b> " + lngE.toFixed(6) + "<br><b>Estado:</b> "+estado+"<br><b>Distancia:</b> " + distStr + " m");

          // -------------------------------------------------------------
          // LÍNEA DE CONEXIÓN (Centro -> Empleado)
          // -------------------------------------------------------------
          L.polyline([[latC, lngC],[latE, lngE]],{color:'gray',weight:2,dashArray:'6,6'}).addTo(map);

          // -------------------------------------------------------------
          // LEYENDA PERSONALIZADA
          // -------------------------------------------------------------
          const legend = L.control({position:'bottomleft'});
          legend.onAdd = function (map) {
            const div = L.DomUtil.create('div', 'legend');
            div.innerHTML = "<b>📋 Leyenda</b><br>🏢 Centro<br>🔵 Radio: " + rd + " m<br>🔵/🔴 Registro(dentro/fuera)<br><b>Distancia:</b> " + distStr + " m";
            return div;
          };
          legend.addTo(map);
          
        } catch (e) {
          document.getElementById('map').innerHTML = '<div class="loading-text">⚠️ Error al cargar el mapa:<br>' + e.message + '</div>';
          console.error(e);
        }
      });
    </script>
  </body>
  </html>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(920).setHeight(780),
    "📍 Mapa Comparativo (Control Horas Extras)"
  );
}

// ======================================================================
// 2. UTILIDADES LOCALES / GLOBALES (Helpers)
// ======================================================================

function _findHeaderIndex(headers, names) {
  if (!headers || !headers.length) return -1;
  const lower = headers.map(h => (h || "").toString().trim().toLowerCase());
  for (const cand of names) {
    const idx = lower.indexOf(cand.toString().trim().toLowerCase());
    if (idx !== -1) return idx + 1;
  }
  return -1;
}

function _parseCoord(v) {
  if (v == null || typeof v === 'undefined') return NaN;
  let s = String(v).trim();
  
  // Manejar textos de error o vacíos
  if (!s || ["[REVISAR]", "NO", "N/A"].includes(s.toUpperCase())) return NaN;  
  
  // REEMPLAZO DE COMA A PUNTO (Manejo explícito solicitado)
  s = s.replace(",", ".");
  
  // Limpiar caracteres que no sean números, puntos o signo negativo
  s = s.replace(/\s+/g, "").replace(/[^0-9.\-]/g, "");  
  
  let n = parseFloat(s);
  
  // Si falla el parseo
  if (isNaN(n)) return NaN;
  
  // Corrección para coordenadas que vengan como grados (ej: lat > 180)
  while (Math.abs(n) > 180) n /= 10;
  
  return n;
}

function _distMetros(aLat, aLng, bLat, bLng) {
  const R = 6371000;
  const dLat = (bLat - aLat) * Math.PI / 180;
  const dLng = (bLng - aLng) * Math.PI / 180;
  const A =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(aLat * Math.PI / 180) * Math.cos(bLat * Math.PI / 180) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  const c = 2 * Math.atan2(Math.sqrt(A), Math.sqrt(1 - A));
  return R * c;
}
