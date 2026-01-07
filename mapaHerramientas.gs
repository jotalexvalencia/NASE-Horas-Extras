// ======================================================================
// 🗺️ mapaHerramientas.gs – Sistema de Menú y Mapas (NASE 2026 - Horas Extras)
// ----------------------------------------------------------------------
/**
 * @summary Módulo de interfaz de usuario (UI) y Visualización para Admin.
 * @description Gestiona el menú principal de Google Sheets y los mapas.
 *              La lógica de lectura de columnas es dinámica, por lo que es compatible
 *              con las nuevas columnas de aprobación de Horas Extras.
 *
 * @features
 *   - 🎛️ Menú personalizado "NASE - Control Horas Extras".
 *   - 🗺️ Visualización de mapas interactivos (Leaflet).
 *   - 📏 Formato estético automático de hojas.
 *
 * @author NASE Team
 * @version 1.7 (Actualizado para Control Horas Extras)
 */

// ======================================================================
// 1. CONFIGURACIÓN DEL MENÚ PRINCIPAL
// ======================================================================

/**
 * @summary Crea el menú personalizado al abrir la hoja.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  // TÍTULO DEL MENÚ ACTUALIZADO
  const menu = ui.createMenu('🧭 NASE - Control Horas Extras');

  // ---------------------------------------------------------
  // 1. Geocodificación y Mapas
  // ---------------------------------------------------------
  menu.addItem('➡ Geocodificar fila actual', 'geocodificarFilaActiva'); 
  menu.addSeparator();
  menu.addItem('🗺️ Ver mapa del registro', 'mostrarMapaDelRegistro');
  menu.addItem('📍 Ver mapa comparativo (centro vs registro)', 'mostrarMapaComparativo');
  menu.addSeparator();

  // ---------------------------------------------------------
  // 2. Utilidades Generales
  // ---------------------------------------------------------
  // Si existe el módulo SalidasPendientes.gs, esta opción funcionará
  menu.addItem('📋 Lista Salidas Pendientes (Visual)', 'mostrarListadoSalidasPendientes');
  menu.addSeparator();

  // ---------------------------------------------------------
  // 3. Gestión de Asistencia y Registros
  // ---------------------------------------------------------
  menu.addItem('📊 Generar Asistencia', 'generarTablaAsistenciaSinValores');
  menu.addSeparator();

  // ---------------------------------------------------------
  // 4. Configuración
  // ---------------------------------------------------------
  menu.addItem('⚙️ Configurar Horas (Inicio y Fin Recargo)', 'mostrarConfiguracionHorarios');
  menu.addSeparator();

  // ---------------------------------------------------------
  // 5. Mantenimiento y Formato
  // ---------------------------------------------------------
  menu.addItem('🧹 Limpieza Profunda de Triggers', 'limpiezaProfundaTriggers');
  menu.addItem('✨ Formatear Hojas (Estilo Profesional)', 'formatearHojasEstandar');

  menu.addToUi();
}

// ======================================================================
// 2. UTILIDADES INTERNAS (Helpers Robustos)
// ======================================================================

/**
 * @summary Busca el índice de una columna dinámicamente por nombre.
 * @description Compatible con las nuevas columnas de Horas Extras.
 */
function _findHeaderIndex(headers, names) {
  if (!headers || !headers.length) return -1;
  const lower = headers.map(h => (h || "").toString().trim().toLowerCase());
  for (const cand of names) {
    const idx = lower.indexOf(cand.toString().trim().toLowerCase());
    if (idx !== -1) return idx + 1; 
  }
  return -1;
}

/**
 * @summary Parsea coordenadas de texto/número a Float.
 */
function _parseCoord(v) {
  if (v === null || typeof v === 'undefined') return NaN;
  let s = String(v).trim();
  
  if (["[REVISAR]", "NO", "N/A"].includes(s.toUpperCase())) return NaN;
  s = s.replace(",", ".");
  s = s.replace(/[^0-9.\-]/g, "");
  
  let n = parseFloat(s);
  if (isNaN(n)) return NaN;
  
  while (Math.abs(n) > 180) n /= 10;
  return n;
}

// ======================================================================
// 3. VISUALIZACIÓN DE MAPAS (Leaflet)
// ======================================================================

/**
 * @summary Muestra un mapa interactivo con el registro de la fila seleccionada.
 */
function mostrarMapaDelRegistro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRes = ss.getActiveSheet();
  const fila = hojaRes.getActiveCell().getRow();

  if (!fila || fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una fila válida en la hoja Respuestas.");
    return;
  }

  // ----------------------------------------------------------------------
  // 1. LEER DATOS DE LA FILA ACTIVA (Búsqueda dinámica segura)
  // ----------------------------------------------------------------------
  const headers = hojaRes.getRange(1, 1, 1, hojaRes.getLastColumn()).getValues()[0] || [];
  
  const idxCentro = _findHeaderIndex(headers, ["centro"]);
  const idxCiudadCentro = _findHeaderIndex(headers, ["ciudad"]);
  const idxLat = _findHeaderIndex(headers, ["lat"]);
  const idxLng = _findHeaderIndex(headers, ["lng"]);
  const idxDir = _findHeaderIndex(headers, ["barrio / dirección", "direccion", "dirección"]);

  const filaVals = hojaRes.getRange(fila, 1, 1, hojaRes.getLastColumn()).getValues()[0];
  
  const centro = (filaVals[idxCentro - 1] || "").toString();
  const ciudadCentro = (filaVals[idxCiudadCentro - 1] || "").toString();
  
  const latEmp = _parseCoord(filaVals[idxLat - 1]);
  const lngEmp = _parseCoord(filaVals[idxLng - 1]);
  const direccion = (filaVals[idxDir - 1] || "").toString();

  if (isNaN(latEmp) || isNaN(lngEmp)) {
    SpreadsheetApp.getUi().alert("⚠️ Coordenadas inválidas en la fila seleccionada.");
    return;
  }

  // ----------------------------------------------------------------------
  // 2. BUSCAR EL CENTRO EN LA HOJA 'Centros'
  // ----------------------------------------------------------------------
  const hojaCentros = ss.getSheetByName("Centros");
  let latCentro = null, lngCentro = null, radio = 30, urlImagenCentro = "";
  
  if (hojaCentros) {
    const dataC = hojaCentros.getDataRange().getValues();
    const headersC = dataC[0].map(h => (h || "").toString().trim().toUpperCase());
    
    const idxC_Centro = _findHeaderIndex(headersC, ["centro"]) - 1;
    const idxC_Ciudad = _findHeaderIndex(headersC, ["ciudad"]) - 1;
    const idxC_Lat = _findHeaderIndex(headersC, ["lat ref", "latitud"]) - 1;
    const idxC_Lng = _findHeaderIndex(headersC, ["lng ref", "longitud"]) - 1;
    const idxC_Radio = _findHeaderIndex(headersC, ["radio"]) - 1;
    const idxC_UrlImagen = _findHeaderIndex(headersC, ["link_imagen", "url imagen", "imagen"]) - 1;

    for (let i = 1; i < dataC.length; i++) {
      const rowC = dataC[i];
      const nombreC = (rowC[idxC_Centro] || "").toString().trim();
      const ciudadC = (rowC[idxC_Ciudad] || "").toString().trim();
      
      if (nombreC.toUpperCase() === centro.toUpperCase() && ciudadC.toUpperCase() === ciudadCentro.toUpperCase()) {
        latCentro = _parseCoord(rowC[idxC_Lat]);
        lngCentro = _parseCoord(rowC[idxC_Lng]);
        radio = rowC[idxC_Radio] ? Number(rowC[idxC_Radio]) : 30;
        urlImagenCentro = (idxC_UrlImagen >= 0 ? (rowC[idxC_UrlImagen] || "").toString().trim() : "");
        break;
      }
    }
  }

  // ----------------------------------------------------------------------
  // 3. GENERAR HTML DEL MAPA
  // ----------------------------------------------------------------------
  const html = `
  <!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
      body {margin:0;padding:10px;font-family:Arial,sans-serif}
      #map {height:680px;width:100%;border-radius:8px}
      .popup-img {width:100%;max-width:300px;border-radius:5px;margin-top:5px}
      .legend {background:#fff;padding:6px;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.15)}
    </style>
  </head>
  <body>
    <div id="map"></div>
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
      const map = L.map('map').setView([${latEmp}, ${lngEmp}], 17);
      const calle = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap'}).addTo(map);
      const sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{attribution:'© Esri'});
      L.control.layers({"🗺️ Callejero":calle,"🌍 Satélite":sat}).addTo(map);

      const iconoRegistro = L.divIcon({
        html:\`<svg height="36" width="36"><circle cx="18" cy="18" r="16" fill="#1976d2" stroke="white" stroke-width="2"/><image href="https://raw.githubusercontent.com/jotalexvalencia/NASE/main/nase_marcador.png" x="10" y="10" width="16" height="16"/></svg>\`,
        iconSize:[36,36],iconAnchor:[18,36],popupAnchor:[0,-36]
      });
      
      const empMarker=L.marker([${latEmp},${lngEmp}],{icon:iconoRegistro}).addTo(map);
      empMarker.bindPopup("<b>📍 Registro Empleado</b><br>${centro}<br>${direccion}<br><b>Lat:</b> ${latEmp.toFixed(6)}<br><b>Lng:</b> ${lngEmp.toFixed(6)}").openPopup();

      if(${latCentro} && ${lngCentro}){
        const centroMarker=L.marker([${latCentro},${lngCentro}]).addTo(map);
        L.circle([${latCentro},${lngCentro}],{color:'#1e88e5',fillColor:'#1e88e5',fillOpacity:0.12,radius:${radio}}).addTo(map);
        let html="<b>🏢 Centro:</b> ${centro}<br>Radio: ${radio} m";
        if("${urlImagenCentro}") html+="<br><img src='${urlImagenCentro}' class='popup-img'>";
        centroMarker.bindPopup(html);
      }
      L.control.scale().addTo(map);
    </script>
  </body></html>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(880).setHeight(760),
    "🌍 Mapa del Registro (Horas Extras)"
  );
}

// ======================================================================
// 4. FORMATO ESTÁNDAR (Estilo Zebra)
// ======================================================================

function aplicarFormatoEstandar(hoja) {
  if (!hoja) return;
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 1 || ultimaColumna < 1) return;

  const encabezados = hoja.getRange(1, 1, 1, ultimaColumna);
  encabezados
    .setFontWeight("bold")
    .setBackground("#17365D")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  for (let i = 2; i <= ultimaFila; i++) {
    hoja.getRange(i, 1, 1, ultimaColumna)
      .setBackground(i % 2 === 0 ? "#D9E1F2" : "#FFFFFF")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  hoja.getDataRange().setBorder(true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  hoja.autoResizeColumns(1, ultimaColumna);
}

function formatearHojasEstandar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Lista de hojas que deben verse profesionales
  const hojas = ['Respuestas', 'Centros', 'Asistencia_SinValores'];
  
  hojas.forEach(n => {
    const hoja = ss.getSheetByName(n);
    if (hoja) aplicarFormatoEstandar(hoja);
  });
  
  SpreadsheetApp.getUi().alert("✅ Formato aplicado a todas las hojas estándar.");
}

// ======================================================================
// 5. CONFIGURACIÓN DE HORARIOS
// ======================================================================

function mostrarConfiguracionHorarios() {
  const template = HtmlService.createTemplateFromFile('config_horarios');
  const html = template.evaluate()
    .setWidth(700)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showSidebar(html);
}
