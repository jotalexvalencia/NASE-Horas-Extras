// ======================================================================
// 🗺️ mapaHerramientas.gs – Menú y Mapas (NASE 2026 - Horas Extras)
// ======================================================================

// ======================================================================
// 1. MENÚ PRINCIPAL
// ======================================================================

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🧭 NASE - Control Horas Extras');

  menu.addItem('➡ Geocodificar fila actual', 'geocodificarFilaActiva'); 
  menu.addSeparator();
  menu.addItem('🗺️ Ver mapa del registro', 'mostrarMapaDelRegistro');
  menu.addItem('📍 Ver mapa comparativo', 'mostrarMapaComparativo');
  menu.addSeparator();
  menu.addItem('📋 Lista Salidas Pendientes', 'mostrarListadoSalidasPendientes');
  menu.addSeparator();
  menu.addItem('📊 Generar Asistencia', 'generarTablaAsistenciaSinValores');
  menu.addSeparator();
  menu.addItem('🔑 Configurar API Key OpenCage', 'guardarApiKeyOpenCage');
  menu.addItem('🧪 Probar Geocodificador', 'testGeocodificador');
  menu.addSeparator();
  menu.addItem('✨ Formatear Hojas', 'formatearHojasEstandar');

  menu.addToUi();
}

// ======================================================================
// 2. UTILIDADES INTERNAS
// ======================================================================

function _findHeaderIndex(headers, names) {
  if (!headers || !headers.length) return -1;
  for (var i = 0; i < names.length; i++) {
    var name = names[i].toString().toLowerCase().trim();
    for (var j = 0; j < headers.length; j++) {
      if ((headers[j] || "").toString().toLowerCase().trim() === name) {
        return j + 1;
      }
    }
  }
  return -1;
}

function _parseCoord(v) {
  if (v === null || typeof v === 'undefined') return NaN;
  var s = String(v).trim();
  if (s.toUpperCase() === "[REVISAR]" || s.toUpperCase() === "NO" || s.toUpperCase() === "N/A") return NaN;
  s = s.replace(",", ".");
  s = s.replace(/[^0-9.\-]/g, "");
  var n = parseFloat(s);
  if (isNaN(n)) return NaN;
  while (Math.abs(n) > 180) n /= 10;
  return n;
}

// ======================================================================
// 3. VISUALIZACIÓN DE MAPAS (Lee Centros de libro EXTERNO)
// ======================================================================

function mostrarMapaDelRegistro() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaRes = ss.getActiveSheet();
  var fila = hojaRes.getActiveCell().getRow();
  var ui = SpreadsheetApp.getUi();

  if (!fila || fila < 2) {
    ui.alert("Selecciona una fila válida.");
    return;
  }

  // Leer datos de la fila
  var headers = hojaRes.getRange(1, 1, 1, hojaRes.getLastColumn()).getValues()[0];
  var filaVals = hojaRes.getRange(fila, 1, 1, hojaRes.getLastColumn()).getValues()[0];
  
  var idxCentro = _findHeaderIndex(headers, ["centro"]);
  var idxCiudad = _findHeaderIndex(headers, ["ciudad"]);
  var idxLat = _findHeaderIndex(headers, ["lat"]);
  var idxLng = _findHeaderIndex(headers, ["lng"]);
  var idxDir = _findHeaderIndex(headers, ["barrio / dirección", "direccion", "dir_geo"]);

  var centro = (filaVals[idxCentro - 1] || "").toString().trim();
  var ciudadCentro = (filaVals[idxCiudad - 1] || "").toString().trim();
  var latEmp = _parseCoord(filaVals[idxLat - 1]);
  var lngEmp = _parseCoord(filaVals[idxLng - 1]);
  var direccion = (idxDir > 0 ? filaVals[idxDir - 1] : "") || "";

  if (isNaN(latEmp) || isNaN(lngEmp)) {
    ui.alert("⚠️ Coordenadas inválidas.");
    return;
  }

  // Buscar centro en libro EXTERNO
  var latCentro = null, lngCentro = null, radio = 30, urlImagenCentro = "";
  
  try {
    var ssExt = SpreadsheetApp.openById(ID_LIBRO_CENTROS_EXTERNO);
    var hojaCentros = ssExt.getSheetByName("Centros");
    if (!hojaCentros) hojaCentros = ssExt.getSheetByName("BASE_CENTROS");
    
    if (hojaCentros) {
      var dataC = hojaCentros.getDataRange().getValues();
      var headersC = dataC[0].map(function(h) { return (h || "").toString().trim().toUpperCase(); });
      
      var idxC_Centro = headersC.indexOf("CENTRO");
      var idxC_Ciudad = headersC.indexOf("CIUDAD");
      var idxC_Lat = -1, idxC_Lng = -1, idxC_Radio = -1, idxC_UrlImg = -1;
      
      for (var i = 0; i < headersC.length; i++) {
        if (headersC[i] === "LAT REF" || headersC[i] === "LAT") idxC_Lat = i;
        if (headersC[i] === "LNG REF" || headersC[i] === "LNG") idxC_Lng = i;
        if (headersC[i] === "RADIO") idxC_Radio = i;
        if (headersC[i] === "LINK_IMAGEN" || headersC[i] === "URL IMAGEN") idxC_UrlImg = i;
      }
      
      for (var i = 1; i < dataC.length; i++) {
        var rowC = dataC[i];
        var nombreC = (rowC[idxC_Centro] || "").toString().trim().toUpperCase();
        var ciudadC = (rowC[idxC_Ciudad] || "").toString().trim().toUpperCase();
        
        if (nombreC === centro.toUpperCase() && ciudadC === ciudadCentro.toUpperCase()) {
          latCentro = _parseCoord(rowC[idxC_Lat]);
          lngCentro = _parseCoord(rowC[idxC_Lng]);
          radio = idxC_Radio >= 0 ? (Number(rowC[idxC_Radio]) || 30) : 30;
          urlImagenCentro = idxC_UrlImg >= 0 ? (rowC[idxC_UrlImg] || "").toString().trim() : "";
          break;
        }
      }
    }
  } catch (e) {
    Logger.log("Error leyendo libro externo: " + e);
  }

  // Generar HTML del mapa
  var hayCentro = !isNaN(latCentro) && !isNaN(lngCentro);
  
  var html = '<!doctype html><html><head><meta charset="utf-8"/>' +
    '<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>' +
    '<style>body{margin:0;padding:10px;font-family:Arial}#map{height:680px;width:100%;border-radius:8px}</style></head><body>' +
    '<div id="map"></div>' +
    '<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>' +
    '<script>' +
    'var map=L.map("map").setView([' + latEmp + ',' + lngEmp + '],17);' +
    'L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{attribution:"© OpenStreetMap"}).addTo(map);' +
    'var empMarker=L.marker([' + latEmp + ',' + lngEmp + ']).addTo(map);' +
    'empMarker.bindPopup("<b>📍 Registro</b><br>' + centro + '<br>Lat:' + latEmp.toFixed(6) + '<br>Lng:' + lngEmp.toFixed(6) + '").openPopup();';
  
  if (hayCentro) {
    html += 'var centroMarker=L.marker([' + latCentro + ',' + lngCentro + ']).addTo(map);' +
            'L.circle([' + latCentro + ',' + lngCentro + '],{color:"#1e88e5",fillColor:"#1e88e5",fillOpacity:0.12,radius:' + radio + '}).addTo(map);' +
            'centroMarker.bindPopup("<b>🏢 Centro:</b> ' + centro + '<br>Radio: ' + radio + 'm");';
  }
  
  html += 'L.control.scale().addTo(map);</script></body></html>';

  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(880).setHeight(760),
    "🌍 Mapa del Registro"
  );
}

// ======================================================================
// 4. FORMATO ESTÁNDAR
// ======================================================================

function aplicarFormatoEstandar(hoja) {
  if (!hoja) return;
  var ultimaFila = hoja.getLastRow();
  var ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 1 || ultimaColumna < 1) return;

  hoja.getRange(1, 1, 1, ultimaColumna)
    .setFontWeight("bold")
    .setBackground("#17365D")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  for (var i = 2; i <= ultimaFila; i++) {
    hoja.getRange(i, 1, 1, ultimaColumna)
      .setBackground(i % 2 === 0 ? "#D9E1F2" : "#FFFFFF")
      .setHorizontalAlignment("center");
  }

  hoja.autoResizeColumns(1, ultimaColumna);
}

function formatearHojasEstandar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = ['Respuestas', 'Asistencia_SinValores'];
  
  for (var i = 0; i < hojas.length; i++) {
    var hoja = ss.getSheetByName(hojas[i]);
    if (hoja) aplicarFormatoEstandar(hoja);
  }
  
  SpreadsheetApp.getUi().alert("✅ Formato aplicado.");
}
