// ======================================================================
// 📍 MapaComparativo.gs – Visualización (NASE 2026 - Horas Extras)
// ----------------------------------------------------------------------
/**
 * @summary Módulo de visualización geoespacial comparativa.
 * @description Genera un mapa interactivo (Leaflet.js) que compara
 *              la ubicación de un empleado contra su centro de trabajo asignado.
 *
 * @features
 *   - 🗺️ Mapa Leaflet: Capas de "Callejero" y "Satélite".
 *   - 🏢 Centro: Marcador del centro de trabajo con popup (Dirección e Imagen).
 *   - 📍 Empleado: Marcador de la ubicación del empleado (Lat/Lng).
 *   - 📏 Distancia: Cálculo de la distancia en metros (Haversine).
 *   - 🎨 Círculo: Dibuja el radio de asistencia del centro (mínimo 80m visual).
 *   - 🚦 Semáforo: Azul si está DENTRO, Rojo si está FUERA del radio.
 *   - 🧭 Línea: Conexión visual entre el centro y el empleado.
 *
 * @author NASE Team
 * @version 1.3 (Corrección Visual - DOMContentLoaded)
 */

// ======================================================================
// 1. FUNCIÓN PRINCIPAL: Generar Mapa Comparativo
// ======================================================================

/**
 * @summary Genera y muestra el mapa comparativo.
 */
function mostrarMapaComparativo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResp = ss.getActiveSheet();
  const fila = hojaResp.getActiveCell().getRow();

  // Validar contexto
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
  
  // Parsear coordenadas (Función auxiliar abajo)
  const latEmp = _parseCoord(filaVals[idxLat - 1]); 
  const lngEmp = _parseCoord(filaVals[idxLng - 1]); 
  const direccion = (filaVals[idxDir - 1] || "").toString();

  // Validar que las coordenadas sean números válidos
  if (isNaN(latEmp) || isNaN(lngEmp)) {
    SpreadsheetApp.getUi().alert("⚠️ Coordenadas inválidas o vacías en la fila seleccionada. Verifica Lat/Lng en la hoja Respuestas.");
    return;
  }
  if (!centro) {
    SpreadsheetApp.getUi().alert("⚠️ La fila seleccionada no tiene un Centro asignado.");
    return;
  }

  // ================================================================
  // 2️⃣ BUSCAR INFORMACIÓN DEL CENTRO DE REFERENCIA (Hoja 'Centros')
  // ================================================================
  const hojaCentros = ss.getSheetByName("Centros");
  if (!hojaCentros) { SpreadsheetApp.getUi().alert("⚠️ La hoja 'Centros' no existe en este libro."); return; } 
  
  const dataCentros = hojaCentros.getDataRange().getValues();
  const headersCentros = dataCentros[0];
  
  const getColC = (name) => _findHeaderIndex(headersCentros, [name]);

  let latCentro = null, lngCentro = null, radio = 30, urlImagenCentro = "", direccionCentro = "";
  
  // Iterar para encontrar el centro
  for (let i = 1; i < dataCentros.length; i++) {
    const rowC = dataCentros[i];
    const nombreC = (rowC[getColC("centro") - 1] || "").toString().trim();
    const ciudadC = (rowC[getColC("ciudad") - 1] || "").toString().trim();
   
    if (nombreC.toUpperCase() === centro.toUpperCase() && ciudadC.toUpperCase() === ciudadCentro.toUpperCase()) {
      latCentro = _parseCoord(rowC[getColC("lat ref") - 1]);
      lngCentro = _parseCoord(rowC[getColC("lng ref") - 1]);
      radio = rowC[getColC("radio") - 1] ? Number(rowC[getColC("radio") - 1]) : 30;
      direccionCentro = rowC[getColC("direccion") - 1] || "";
      urlImagenCentro = (getColC("link_imagen") >= 0 ? (rowC[getColC("link_imagen") - 1] || "").toString().trim() : "");
      break;
    }
  }

  if (isNaN(latCentro) || isNaN(lngCentro))
    return SpreadsheetApp.getUi().alert(`⚠️ Coordenadas del centro "${centro}" no válidas en la hoja 'Centros'.`);

  // ================================================================
  // 3️⃣ CÁLCULO DE DISTANCIA Y ESTADO
  // ================================================================
  const distancia = _distMetros(latCentro, lngCentro, latEmp, lngEmp); 
  const dentro = distancia <= radio;

  // ================================================================
  // 4️⃣ GENERAR HTML DEL MAPA (Con DOMContentLoaded para evitar error)
  // ================================================================
  const leafletCSS = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
  const leafletJS = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";

  const html = `
  <!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <link rel="stylesheet" href="${leafletCSS}" />
    <style>
      body {margin:0;padding:10px;font-family:Arial,sans-serif}
      #map {height:680px;width:100%;border-radius:8px}
      .popup-img {width:100%;max-width:300px;border-radius:5px;margin-top:5px}
      .legend {background:#fff;padding:6px;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.12)}
    </style>
  </head>
  <body>
    <div class="info">🏢 Centro: ${centro}</div>
    <div id="map"></div>
    <script src="${leafletJS}"></script>
    <script>
      // Esperar a que Leaflet cargue para evitar errores "L is not defined"
      document.addEventListener('DOMContentLoaded', function() {
        try {
          const latC = ${latCentro};
          const latE = ${latEmp};
          const lngC = ${lngCentro};
          const lngE = ${lngEmp};
          const rd = ${radio};

          // Validación por seguridad dentro del HTML
          if (isNaN(latC) || isNaN(latE) || isNaN(lngC) || isNaN(lngE)) {
             document.getElementById('map').innerHTML = "⚠️ Coordenadas inválidas. Imposible mostrar mapa.";
             return;
          }

          const map = L.map('map').setView([ ((latC + latE) / 2), ((lngC + lngE) / 2) ], 15);
          
          // Capa 1: Callejero
          const calle = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap'}).addTo(map);
          
          // Capa 2: Satélite
          const sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{attribution:'© Esri'});
          
          // Control de capas
          L.control.layers({"🗺️ Callejero":calle,"🌍 Satélite":sat}).addTo(map);

          // Marcador del Centro
          const centroMarker = L.marker([latC, lngC]).addTo(map);
          
          let htmlCentro = "<b>🏢 Centro de Trabajo</b><br>${centro}<br><b>Ciudad:</b> ${ciudadCentro}<br><b>Dirección:</b> ${direccionCentro}<br><b>Radio:</b> ${rd} m";
          const imgUrl = "${urlImagenCentro}";
          if (imgUrl && imgUrl.length > 5) {
             htmlCentro += "<br><img src='" + imgUrl + "' class='popup-img'>";
          }

          centroMarker.bindPopup(htmlCentro);

          // Círculo del Radio
          L.circle([latC, lngC],{
            color:'#2e7d32',fillColor:'#2e7d32',fillOpacity:0.12,
            radius:Math.max(rd, 80)
          }).addTo(map);

          // Marcador del Empleado
          const color = ${dentro} ? '#2e7d32' : '#d32f2f';
          
          const icono = L.divIcon({
            html:\`<svg height="36" width="36"><circle cx="18" cy="18" r="16" fill="\${color}" stroke="white" stroke-width="2"/></svg>\`,
            iconSize:[36,36],iconAnchor:[18,36],popupAnchor:[0,-36]
          });
          
          const emp = L.marker([latE, lngE], {icon:icono}).addTo(map);
          const estado = ${dentro} ? '✅ DENTRO del radio' : '❌ FUERA del radio';
          const dist = (${distancia}).toFixed(1);

          emp.bindPopup("<b>📍 Registro Empleado</b><br>${centro}<br>${direccion}<br><b>Lat:</b> " + latE.toFixed(6) + "<br><b>Lng:</b> " + lngE.toFixed(6) + "<br><b>Estado:</b> " + estado + "<br><b>Distancia:</b> " + dist + " m");

          // Línea de conexión
          L.polyline([[latC, lngC],[latE, lngE]],{color:'gray',weight:2,dashArray:'6,6'}).addTo(map);

          // Leyenda
          const legend = L.control({position:'bottomleft'});
          legend.onAdd = function (map) {
            const div = L.DomUtil.create('div', 'legend');
            div.innerHTML = "<b>📋 Leyenda</b><br>🏢 Centro<br>🔵 Radio: " + rd + " m<br>🔵/🔴 Registro(dentro/fuera)<br><b>Distancia:</b> " + dist + " m";
            return div;
          };
          legend.addTo(map);

        } catch (e) {
          document.getElementById('map').innerHTML = "⚠️ Error al cargar el mapa: " + e.message;
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
// 2. UTILIDADES LOCALES (Helpers)
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
  if (!s) return NaN;  
  if (["[REVISAR]", "NO", "N/A"].includes(s.toUpperCase())) return NaN;  
  s = s.replace(/\s+/g, "").replace(",", ".");
  s = s.replace(/[^0-9.\-]/g, "");  
  let n = parseFloat(s);
  if (isNaN(n)) return NaN;  
  // Corrección rápida para coordenadas en formato decimal si vienen como 4.7
  if (Math.abs(n) > 180) n /= 10;
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
  const c = 2 * R * Math.atan2(Math.sqrt(A), Math.sqrt(1 - A));
  return R * c;
}
