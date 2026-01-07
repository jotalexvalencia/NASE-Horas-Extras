// ======================================================================
// 📋 SalidasPendientes.gs – Cierre de Turnos (NASE 2026 - Horas Extras)
// ----------------------------------------------------------------------
/**
 * @summary Módulo de Administración para cierre de turnos (Entrada/Salida).
 * @description Este es el módulo administrativo "Visual y Robusto". Centraliza la
 *              gestión de turnos abiertos en una interfaz Sidebar.
 * 
 * @features
 *   - 🎨 **Formato Limpio:** Convierte fechas tipo "Fri Jan 02..." a "dd/mm/yyyy".
 *   - ✍️ **Validación Manual:** Permite ingresar fecha/hora en formato "dd/mm/yyyy hh:mm".
 *   - 🎯 **Detección de Basura:** Identifica celdas con "Jan 02 00:00" y las trata como pendientes.
 *   - 👁️ **Auditoría:** Registra todos los cambios en una hoja dedicada.
 *   - 🆕 **Compatibilidad Horas Extras:** El mapeo de columnas ahora incluye las 
 *       columnas de aprobación (`Estado HE`, `Aprobado Supervisor`, etc.) para 
 *       asegurar que la estructura de la hoja se mantenga íntegra.
 *
 * @author NASE Team
 * @version 2.1 (Actualizado para Horas Extras)
 */

// ======================================================================
// 1. CONFIGURACIÓN DEL SISTEMA
// ======================================================================

const SP_SHEET_RESPUESTAS = 'Respuestas';
const SP_TZ = "America/Bogota";

// Modo de Auditoría (Sheet = Guarda en hoja, None = No hace nada)
const SP_ADMIN_AUDIT_MODE = "sheet"; 
const SP_AUDIT_SHEET = "Auditoria_Salidas_Admin";

// -------------------------------------------------------------------
// 1.1 ENCABEZADOS OFICIALES HOJA RESPUESTAS (Mapeo de Índices)
// -------------------------------------------------------------------
// Actualizado para incluir columnas de Horas Extras (Índices 21-27)
const SP_HEADERS = [
  "Cédula",
  "Centro",
  "Ciudad",
  "Lat",
  "Lng",
  "Acepto",
  "Ciudad_Geo",
  "Dir_Geo",
  "Accuracy",
  "Dentro",      
  "Distancia",
  "Observaciones",
  "Nombre",
  "Foto",
  "Fecha Entrada", // 14
  "Hora Entrada",   // 15
  "Foto Entrada",
  "Fecha Salida", // 17
  "Hora Salida",   // 18
  "Foto Salida",
  "Dentro Salida", // 20
  // --- NUEVAS COLUMNAS HORAS EXTRAS ---
  "Total Horas Extras",     // 21
  "Total Horas Nocturnas",  // 22
  "Estado HE",             // 23
  "Aprobado Supervisor",   // 24
  "Fecha Aprueba Super",   // 25
  "Aprobado Director",     // 26
  "Fecha Aprueba Director"  // 27
];

/**
 * @summary Mapa de Índices (0-based) para acceso rápido a columnas.
 * @description Alineado con el array `SP_HEADERS` y el sistema completo.
 */
const SP_I = {
  CED: 0,
  CENTRO: 1,
  CIUDAD: 2,
  LAT: 3,
  LNG: 4,
  ACEPTO: 5,
  OBS: 11,
  NOMBRE: 12,
  FOTO: 13,
  FECHA_ENT: 14,
  HORA_ENT: 15,
  FOTO_ENT: 16,
  FECHA_SAL: 17,
  HORA_SAL: 18,
  FOTO_SAL: 19,
  DENTRO_SAL: 20,
  // Nuevos Índices (Necesarios para spAsegurarEncabezados_)
  TOTAL_HE: 21,
  TOTAL_NOCT: 22,
  ESTADO_HE: 23,
  APROB_SUPER: 24,
  FECHA_APROB_SUPER: 25,
  APROB_DIR: 26,
  FECHA_APROB_DIR: 27
};

// ======================================================================
// 2. UTILIDADES DE FORMATEO
// ======================================================================

function spAsegurarEncabezados_(hoja) {
  const lastCol = hoja.getLastColumn();
  if (lastCol === 0) {
    hoja.getRange(1, 1, 1, SP_HEADERS.length).setValues([SP_HEADERS]);
    return;
  }
  const current = hoja.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const currentLen = current.length;
  
  // Si faltan columnas al final, agregarlas (Incluye las nuevas de HE)
  if (currentLen < SP_HEADERS.length) {
    const faltantes = SP_HEADERS.slice(currentLen);
    hoja.getRange(1, currentLen + 1, 1, faltantes.length).setValues([faltantes]);
  }
}

function spFormatearFecha_(valor) {
  if (!valor) return '';  
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, SP_TZ, "dd/MM/yyyy");
  }
  const str = String(valor || '').trim();  
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
    return str;
  }
  try {
    const dt = new Date(str);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, SP_TZ, "dd/MM/yyyy");
    }
  } catch(e) {}  
  return str; 
}

function spFormatearHora_(valor) {
  if (!valor) return '';  
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, SP_TZ, "HH:mm:ss");
  }
  const str = String(valor || '').trim();  
  if (/^\d{2}:\d{2}(:\d{2})?$/.test(str)) {
    return str;
  }
  try {
    const dt = new Date(str);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, SP_TZ, "HH:mm:ss");
    }
  } catch(e) {}  
  return str;
}

function spParsearFechaHora_(textoFechaHora) {
  if (!textoFechaHora) return null;  
  const texto = String(textoFechaHora).trim();  
  const partes = texto.split(' ');
  if (partes.length < 2) return null;  
  const fechaParts = partes[0].split('/'); 
  const horaParts = partes[1].split(':'); 
  if (fechaParts.length !==3 || horaParts.length < 2) return null;  
  const dia = fechaParts[0].padStart(2, '0');
  const mes = fechaParts[1].padStart(2, '0');
  const anio = fechaParts[2];
  const hora = horaParts[0].padStart(2, '0');
  const min = horaParts[1].padStart(2, '0');
  const seg = horaParts[2] ? horaParts[2].padStart(2, '0') : '00';
  const isoStr = `${anio}-${mes}-${dia}T${hora}:${min}:${seg}`;
  const dt = new Date(isoStr);  
  return isNaN(dt.getTime()) ? null : dt;
}

// ======================================================================
// 3. FUNCIÓN PRINCIPAL DEL MENÚ
// ======================================================================

function mostrarListadoSalidasPendientes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SP_SHEET_RESPUESTAS);

  if (!hoja) { ui.alert("❌ No existe la hoja 'Respuestas'."); return; }  
  spAsegurarEncabezados_(hoja);

  const data = getRawPendingList_(); 
  
  if (!data || data.length === 0) {
    ui.alert("✅ No hay salidas pendientes. Todos los turnos están cerrados.");
    return;
  }

  // HTML DEL SIDEBAR
  let html = '<style>';
  html += 'body { font-family: Arial, sans-serif; font-size: 12px; } ';
  html += 'h3 { color: #17365D; font-size: 18px; margin-bottom: 10px; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px; } ';
  html += '.table-container { max-height: 500px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px; } ';
  html += 'table { width: 100%; border-collapse: collapse; } ';
  html += 'th { background-color: #f8f9fa; color: #17365D; padding: 10px; text-align: left; position: sticky; top: 0; border-bottom: 2px solid #eee; } ';
  html += 'td { padding: 10px; border-bottom: 1px solid #f1f3f5; vertical-align: middle; } ';
  html += 'tr:hover { background-color: #f1f8ff; } ';
  html += '.btn-cerrar { background-color: #ffc107; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-weight: bold; } ';
  html += '.btn-cerrar:hover { background-color: #e0a800; } ';
  html += '</style>';

  html += '<h3>📋 Salidas Pendientes (Admin)</h3>';
  html += '<p>Lista de turnos abiertos. Presiona "🔒 Cerrar" para cerrar turno.</p>';
  
  html += '<div class="table-container"><table>';
  html += '<thead><tr><th>Cédula</th><th>Nombre</th><th>Centro</th><th>Entrada (Fecha/Hora)</th><th>Acción</th></tr></thead>';
  html += '<tbody>';

  data.forEach(item => {
    const fechaVisual = `${item.fechaEntrada} ${item.horaEntrada}`; 
    
    html += '<tr>';
    html += `<td><strong>${item.cedula}</strong></td>`;
    html += `<td>${item.nombre}</td>`;
    html += `<td>${item.centro}</td>`;
    html += `<td>${fechaVisual}</td>`; 
    // Botón que llama al backend
    html += `<td><button class="btn-cerrar" onclick="google.script.run.withSuccessHandler(google.script.host.close).corregirDesdeSidebar('${item.cedula}')">🔒 Cerrar</button></td>`;
    html += '</tr>';
  });

  html += '</tbody></table></div>';

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setTitle('Corrección de Turnos - Horas Extras')
    .setWidth(500);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// ======================================================================
// 4. FUNCIÓN DE DATOS (Backend)
// ======================================================================

function getRawPendingList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SP_SHEET_RESPUESTAS);
  if (!hoja) return [];  
  spAsegurarEncabezados_(hoja);
  const datos = hoja.getDataRange().getValues();
  if (!datos || datos.length <=1) return [];

  const listaFinal = [];

  for (let i =1; i < datos.length; i++) {
    const row = datos[i];
    const ced = String(row[SP_I.CED] || '').trim();
    if (!ced) continue;

    // Lógica Robusta de "PENDIENTE"
    let fechaSalStr = '';
    if (!row[SP_I.FECHA_SAL]) fechaSalStr = '';
    else if (row[SP_I.FECHA_SAL] instanceof Date) fechaSalStr = Utilities.formatDate(row[SP_I.FECHA_SAL], SP_TZ, "dd/MM/yyyy");
    else fechaSalStr = String(row[SP_I.FECHA_SAL]).trim();

    const esBasura = fechaSalStr.includes('Jan 02 00:00') || fechaSalStr.includes('Sun Dec 14 00:00');

    if (!fechaSalStr || fechaSalStr === '' || esBasura) {
      
      const fechaFormateada = spFormatearFecha_(row[SP_I.FECHA_ENT]);
      const horaFormateada = spFormatearHora_(row[SP_I.HORA_ENT]);
      
      listaFinal.push({
        cedula: ced,
        nombre: String(row[SP_I.NOMBRE] || 'Desconocido'),
        centro: String(row[SP_I.CENTRO] || ''),
        fechaEntrada: fechaFormateada, 
        horaEntrada: horaFormateada    
      });
    }
  }

  listaFinal.sort((a, b) => {
    const parseF = (f) => {
      const p = f.split('/');
      return p.length ===3 ? new Date(`${p[2]}-${p[1]}-${p[0]}`) : new Date(0);
    };
    return parseF(b.fechaEntrada) - parseF(a.fechaEntrada);
  });

  return listaFinal;
}

// ======================================================================
// 5. FUNCIÓN DE CORRECCIÓN (Sidebar)
// ======================================================================

function corregirDesdeSidebar(cedula) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SP_SHEET_RESPUESTAS);
  if (!hoja) { ui.alert("❌ Hoja no encontrada"); return; }

  spAsegurarEncabezados_(hoja);

  const ced = String(cedula).replace(/\D/g, '').trim();
  const finder = hoja.createTextFinder(ced).matchEntireCell(true);
  const hits = finder.findAll(); 
  
  if (!hits || hits.length === 0) {
    ui.alert("❌ Cédula no encontrada.");
    return;
  }

  let rowEntrada = -1;
  hits.sort((a, b) => b.getRow() - a.getRow());

  for (let i =0; i < hits.length; i++) {
    const r = hits[i].getRow();
    if (r <= 1) continue;
    
    const cSalida = String(hoja.getRange(r, SP_I.FECHA_SAL +1).getValue() || '').trim();
    
    if (!cSalida) { 
      rowEntrada = r; 
      break; 
    }
  }

  if (rowEntrada === -1) {
    ui.alert("❌ No se encontró entrada abierta.");
    return; 
  }

  const entradaVals = hoja.getRange(rowEntrada, 1, 1, hoja.getLastColumn()).getValues()[0];
  const nombre = String(entradaVals[SP_I.NOMBRE] || 'Sin nombre');
  const centro = String(entradaVals[SP_I.CENTRO] || '');
  
  const fechaEntradaStr = spFormatearFecha_(entradaVals[SP_I.FECHA_ENT]);
  const horaEntradaStr = spFormatearHora_(entradaVals[SP_I.HORA_ENT]);
  const entradaCompleta = `${fechaEntradaStr} ${horaEntradaStr}`;  
  const dtEntrada = spParsearFechaHora_(entradaCompleta) || new Date(0);  
  const ahora = new Date();
  const sugerencia = Utilities.formatDate(ahora, SP_TZ, "dd/MM/yyyy HH:mm:ss");
  
  const mensaje = `Empleado: ${nombre}\nCédula: ${ced}\nCentro: ${centro}\nEntrada: ${entradaCompleta}\n\nIngresa fecha y hora de cierre:`;

  const fechaResp = ui.prompt(
    'FECHA Y HORA DE CIERRE',
    `${mensaje}\n\nFormato requerido: DD/MM/YYYY HH:MM:SS\nEjemplo: ${sugerencia}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (fechaResp.getSelectedButton() !== ui.Button.OK) return;

  const textoFecha = String(fechaResp.getResponseText() || '').trim();  
  const cierreDt = spParsearFechaHora_(textoFecha);

  if (!cierreDt) {
    ui.alert('❌ Formato inválido.\n\nUsa: DD/MM/YYYY HH:MM:SS\nEjemplo: 02/01/2026 14:30:00');
    return;
  }
  
  if (dtEntrada.getTime() > 0 && cierreDt <= dtEntrada) {
    ui.alert(`❌ La fecha/hora de cierre debe ser POSTERIOR a la entrada.\n\nEntrada: ${entradaCompleta}\nCierre ingresado: ${textoFecha}`);
    return;
  }

  const fechaDDMMYYYY = Utilities.formatDate(cierreDt, SP_TZ, "dd/MM/yyyy");
  const horaHHMMSS = Utilities.formatDate(cierreDt, SP_TZ, "HH:mm:ss");

  spSetSalidaEnFilaEntrada_(hoja, rowEntrada, fechaDDMMYYYY, horaHHMMSS, '', '');
  
  // Auditoría
  if (SP_ADMIN_AUDIT_MODE === "sheet") {
    const audit = spEnsureAuditSheet_(ss);
    const usuario = (Session.getActiveUser && Session.getActiveUser().getEmail) ? (Session.getActiveUser().getEmail() || 'Desconocido') : 'ScriptUser';
    const obsActual = String(hoja.getRange(rowEntrada, SP_I.OBS +1).getValue() || '');
    const obsFinal = obsActual ? `[ADMIN] ${obsActual}` : 'Cerrado desde Sidebar';
    hoja.getRange(rowEntrada, SP_I.OBS +1).setValue(obsFinal);

    audit.appendRow([
      new Date(),
      ced,
      nombre,
      centro,
      rowEntrada,
      fechaDDMMYYYY,
      horaHHMMSS,
      "Cerrado desde Sidebar Admin",
      usuario
    ]);
  }

  ui.alert(`✅ Turno cerrado correctamente.\n\nCédula: ${ced}\nEmpleado: ${nombre}\nFila: ${rowEntrada}\nCierre: ${fechaDDMMYYYY} ${horaHHMMSS}`);
}

// ======================================================================
// 6. UTILIDAD DE ESCRITURA
// ======================================================================

function spSetSalidaEnFilaEntrada_(hoja, rowEntrada, fechaDDMMYYYY, horaHHMMSS, fotoSalida, dentroSalida) {
  const colFechaSal = SP_I.FECHA_SAL +1;
  const colHoraSal = SP_I.HORA_SAL +1;
  const colFotoSal = SP_I.FOTO_SAL +1;
  const colDentroSal = SP_I.DENTRO_SAL +1;

  hoja.getRange(rowEntrada, colFechaSal).setValue(fechaDDMMYYYY);
  hoja.getRange(rowEntrada, colHoraSal).setValue(horaHHMMSS);
  hoja.getRange(rowEntrada, colFotoSal).setValue(fotoSalida || '');
  hoja.getRange(rowEntrada, colDentroSal).setValue(dentroSalida || '');
}

// ======================================================================
// 7. AUDITORÍA
// ======================================================================

function spEnsureAuditSheet_(ss) {
  if (SP_ADMIN_AUDIT_MODE !== "sheet") return null; 
  let sh = ss.getSheetByName(SP_AUDIT_SHEET);
  if (!sh) sh = ss.insertSheet(SP_AUDIT_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "TimestampAdmin", "Cedula", "Nombre", "Centro",
      "FilaEntradaCerrada", "FechaSalida", "HoraSalida",
      "Motivo", "UsuarioAdmin"
    ]);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f3f3f3');
  }
  return sh;
}
