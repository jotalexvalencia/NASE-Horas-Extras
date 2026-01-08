// ===================================================================
// 🧠 Code.gs – NASE2025 – SISTEMA HORAS EXTRAS (MODIFICADO)
// ===================================================================
/**
 * @summary Archivo principal backend - Modificado para Control de Horas Extras.
 * @description Agrega lógica de aprobación en cascada (Supervisor -> Director).
 */

// -------------------------------------------------------------------
// 1. CONFIGURACIÓN GENERAL
// -------------------------------------------------------------------
const SHEET_ID = '12G2vLux_pG0rBFZXodyNqo01ruEHkynSL-Kgl8Yhbl0'; // <--- CAMBIAR ESTE ID AL DEL NUEVO LIBRO
const SHEET_NAME = 'Respuestas';
const SHEET_CENTROS = 'Centros';
const SHEET_ASISTENCIA = 'Asistencia_SinValores';
const ID_LIBRO_BASE = '1bU-lyiQzczid62n8timgUguW6UxC3qZN8Vehnn26zdY';
const ID_CARPETA_FOTOS = '1J2-204Chw5zg2xVOYKwUK1KIvVQOX3ij';
const TZ = "America/Bogota";
const RADIO_DEFAULT = 30;

const CHUNK_SIZE = 95 * 1024;
const CACHE_DURATION = 21600;

let centrosDataCache = null;
let empleadosCacheRAM = null;
let ultimaActualizacionCache = 0;
const CACHE_LOCAL_DURATION = 21600000;

// -------------------------------------------------------------------
// 1.1 ENCABEZADOS OFICIALES (ACTUALIZADOS CON HORAS EXTRAS)
// -------------------------------------------------------------------
const RESP_HEADERS = [
  "Cédula",      "Centro",      "Ciudad",      "Lat",         "Lng",         "Acepto",      
  "Ciudad_Geo",  "Dir_Geo",     "Accuracy",    "Dentro",      "Distancia",   
  "Observaciones","Nombre",      "Foto",        "Fecha Entrada","Hora Entrada",
  "Foto Entrada", "Fecha Salida", "Hora Salida", "Foto Salida","Dentro Salida",
  // --- NUEVAS COLUMNAS PARA HORAS EXTRAS ---
  "Total Horas Extras",      // Columna 21
  "Total Horas Nocturnas",   // Columna 22
  "Estado HE",               // Columna 23 (Pendiente Supervisor, Pendiente Director, Aprobado)
  "Aprobado Supervisor",     // Columna 24 (Email)
  "Fecha Aprueba Super",     // Columna 25
  "Aprobado Director",       // Columna 26 (Email)
  "Fecha Aprueba Director"   // Columna 27
];

const RESP_I = {
  CEDULA: 0,      CENTRO: 1,      CIUDAD: 2,      LAT: 3,         LNG: 4,         ACEPTO: 5,
  CIUDAD_GEO: 6,  DIR_GEO: 7,     ACCURACY: 8,    DENTRO: 9,      DISTANCIA: 10,
  OBS: 11,        NOMBRE: 12,     FOTO: 13,       FECHA_ENT: 14,  HORA_ENT: 15,
  FOTO_ENT: 16,   FECHA_SAL: 17,  HORA_SAL: 18,   FOTO_SAL: 19,   DENTRO_SAL: 20,
  // Índices Nuevos
  TOTAL_HE: 21,   TOTAL_NOCT: 22, ESTADO: 23,     APROB_SUPER: 24, FECHA_APROB_SUPER: 25,
  APROB_DIR: 26,  FECHA_APROB_DIR: 27
};

function diagnosticarSistema() {
  Logger.log('=== DIAGNÓSTICO SISTEMA NASE HE ===');
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Acceso al libro: ' + ss.getName());
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      Logger.log('ERROR: Hoja Respuestas NO EXISTE');
      return;
    }
    Logger.log('Hoja encontrada: ' + SHEET_NAME);
    const filtros = {
      fechaInicio: Utilities.formatDate(new Date(2024, 0, 1), TZ, 'yyyy-MM-dd'),
      fechaFin: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')
    };
    const resultado = obtenerRegistros(filtros);
    Logger.log('Resultado: ' + JSON.stringify(resultado));
  } catch (e) {
    Logger.log('ERROR CRÍTICO: ' + e.toString());
  }
} 

// -------------------------------------------------------------------
// 1.2 CONFIGURACIÓN DE PERMISOS (SEGURIDAD)
// -------------------------------------------------------------------
// Lista de Supervisores (Pueden aprobar el primer nivel)
const PERMISOS_CONSULTA = [
  "supervisorbogota1@nasecolombia.com.co",
  "supervisorbogota2@nasecolombia.com.co",
  "supervisorbogota3@nasecolombia.com.co",
  "supervisorcali1@nasecolombia.com.co",
  "supervisorcali2@nasecolombia.com.co",
  "supervisorcartagena@nasecolombia.com.co",
  "supervisoribague@nasecolombia.com.co",
  "supervisorinterno@nasecolombia.com.co",
  "supervisormedellin@nasecolombia.com.co",
  "supervisorneiva@nasecolombia.com.co",
  "supervisorpereira1@nasecolombia.com.co",
  "supervisorpereira2@nasecolombia.com.co",
  "supervisorpereira3@nasecolombia.com.co",
  "supervisorpereira4@nasecolombia.com.co",
  "supervisorquindio@nasecolombia.com.co",
  "administraciondigital@nasecolombia.com.co",
  "directorctt@nasecolombia.com.co",
  "analistaprogramador@nasecolombia.com.co"
];

// Lista de Director Nacional de Operaciones (Puede aprobar el segundo nivel)
const PERMISOS_DIRECTOR = [
  "directornacionaloperaciones@nasecolombia.com.co",
  "analistaprogramador@nasecolombia.com.co"
];

// Los permisos de asistencia y centros se mantienen igual
const PERMISOS_ASISTENCIA = [
  "analistanomina@nasecolombia.com.co",
  "lidernomina@nasecolombia.com.co",
  "administraciondigital@nasecolombia.com.co",
  "directorctt@nasecolombia.com.co",
  "analistaprogramador@nasecolombia.com.co"
];
const PERMISOS_CENTROS = PERMISOS_CONSULTA;

// -------------------------------------------------------------------
// 2. WEB APP ROUTING (CON CONTROL DE ACCESO)
// -------------------------------------------------------------------
function doGet(e) {
  var emailUsuario = Session.getActiveUser().getEmail();
  var page = e.parameter.page || 'form';

  // -----------------------------------------------------------------
  // PÁGINA 1: CONSULTA (Solo Supervisores/Directores)
  // -----------------------------------------------------------------
  if (page === 'consulta') {
    if (PERMISOS_CONSULTA.indexOf(emailUsuario) === -1 && PERMISOS_DIRECTOR.indexOf(emailUsuario) === -1) {
      return generarPaginaAccesoDenegado(emailUsuario, "Consulta Horas Extras");
    }
    return HtmlService.createTemplateFromFile('consulta')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setTitle('NASE - Consulta Horas Extras');

  } 
  // -----------------------------------------------------------------
  // PÁGINA 2: ASISTENCIA (Solo Directores/RRHH)
  // -----------------------------------------------------------------
  else if (page === 'asistencia') {
    if (PERMISOS_ASISTENCIA.indexOf(emailUsuario) === -1) {
      return generarPaginaAccesoDenegado(emailUsuario, "Nómina y Asistencia");
    }
    return HtmlService.createTemplateFromFile('asistencia')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setTitle('NASE - Consulta Asistencia');
  } 
  
  // -----------------------------------------------------------------
  // ❌ ELIMINADO: BLOQUE 'ACTUALIZAR_CENTROS'
  // Ya no necesitamos este bloque porque estamos leyendo los centros
  // directamente del libro externo en la función 'getCentrosData' arriba.
  // -----------------------------------------------------------------

  // -----------------------------------------------------------------
  // PÁGINA 3: FORMULARIO (Registro de Entradas/Salidas) - PÁGINA POR DEFECTO
  // -----------------------------------------------------------------
  else {
    // Esta es la página que carga si escribes la URL sin ?page=...
    // Carga el form.html y le inyecta los centros.
    
    const template = HtmlService.createTemplateFromFile('form');
    try {
      // Llamamos a getCentrosData (que ya tiene la lógica del libro externo)
      const dataObj = getCentrosData();
      
      // Estructurar datos para el HTML
      const structured = (dataObj && dataObj.structured) ? dataObj.structured : {};
      const centrosJson = JSON.stringify(structured);
      
      // Inyectar en la plantilla
      // replace reemplaza barras invertidas dobles para evitar errores JSON
      template.centrosInyectados = centrosJson.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
      
    } catch (err) {
      // Si falla la carga de centros, inyectar vacío para no romper la página
      Logger.log("Error cargando centros en doGet: " + err);
      template.centrosInyectados = "{}";
    }
    
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setTitle('NASE - Registro Horas Extras');
  }
}

function asegurarEncabezadosRespuestas_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) {
    sh.getRange(1, 1, 1, RESP_HEADERS.length).setValues([RESP_HEADERS]);
    return;
  }
  const current = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const currentLen = current.length;
  const allBlank = current.every(function(v) { return String(v || "").trim() === ""; });
  if (allBlank) {
    sh.getRange(1, 1, 1, RESP_HEADERS.length).setValues([RESP_HEADERS]);
    return;
  }
  // Si faltan columnas nuevas (agregadas para Horas Extras), las agrega
  if (currentLen < RESP_HEADERS.length) {
    const faltantes = RESP_HEADERS.slice(currentLen);
    sh.getRange(1, currentLen + 1, 1, faltantes.length).setValues([faltantes]);
  }
}

// ===================================================================
// 3. LÓGICA DE APROBACIÓN (NUEVO MÓDULO)
// ===================================================================

// ===================================================================
// 3. LÓGICA DE APROBACIÓN (DETECCIÓN AUTOMÁTICA DE ROL)
// ===================================================================

/**
 * @summary Aprueba un registro de Horas Extras (Backend decide el Rol).
 * @description Determina el rol del usuario basado en su correo en las constantes.
 *              Valida la secuencia de aprobación (Supervisor -> Director).
 *
 * @permissions
 *   - Director Nacional: Puede aprobar registros en estado "Pendiente Supervisor" (pasa a Pendiente Dir)
 *                          y registros en estado "Pendiente Director" (pasa a Aprobado).
 *   - Supervisor: Solo puede aprobar registros en estado "Pendiente Supervisor".
 *
 * @param {Number} rowIndex - Índice de la fila (1-based).
 * @returns {Object} { status: 'ok'|'error', message: String }
 */
function aprobarHorasExtras(rowIndex) {
  // 1. BLOQUEO Y CARGA
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) return { status: 'error', message: 'El sistema está ocupado. Intente de nuevo.' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { status: 'error', message: 'Hoja no encontrada' };

    const fila = Number(rowIndex);
    if (fila < 2) return { status: 'error', message: 'Fila inválida' };

    // 2. LEER ESTADO ACTUAL DEL REGISTRO
    const estadoActual = sh.getRange(fila, RESP_I.ESTADO + 1).getValue();
    const estadoNormalizado = String(estadoActual || '').trim().toLowerCase();

    // 3. DETERMINAR ROL DEL USUARIO AUTOMÁTICAMENTE
    const emailUsuario = Session.getActiveUser().getEmail();
    var rolUsuario = '';

    // Prioridad: Si eres Director, tienes poder de Supervisor también.
    if (PERMISOS_DIRECTOR.indexOf(emailUsuario) !== -1) {
      rolUsuario = 'director';
    } else if (PERMISOS_CONSULTA.indexOf(emailUsuario) !== -1) {
      rolUsuario = 'supervisor';
    } else {
      return { status: 'error', message: 'Acceso Denegado: Tu correo no tiene permisos de aprobación.' };
    }

    // --- LOGS DE DEPURACIÓN ---
    Logger.log("============================================");
    Logger.log("APROBACIÓN AUTOMÁTICA (Rol Detectado)");
    Logger.log("Usuario: " + emailUsuario);
    Logger.log("Rol Detectado: " + rolUsuario);
    Logger.log("Estado Actual Fila " + fila + ": " + estadoActual);

    // 4. VALIDACIÓN DE PERMISOS (Seguridad de Niveles)
    if (rolUsuario === 'supervisor') {
      // El Supervisor SOLO puede actuar si el estado es "Pendiente Supervisor".
      // No puede mover de "Pendiente Director" a "Aprobado".
      if (estadoNormalizado !== '' && 
          estadoNormalizado !== 'pendiente supervisor') {
             Logger.log("❌ NEGADO: Supervisor intentando aprobar estado '" + estadoActual + "'");
             return { status: 'error', message: 'Este registro está a la espera del Director Nacional. Supervisores solo pueden aprobar el primer nivel.' };
      }
    } 
    // El Director puede hacer cualquier cosa, no hay restricción para él excepto que el registro esté completado.

    // 5. LÓGICA DE ACTUALIZACIÓN
    if (estadoNormalizado === 'pendiente supervisor' || estadoNormalizado === '') {
      // Primer paso: Pendiente -> Director
      sh.getRange(fila, RESP_I.ESTADO + 1).setValue('Pendiente Director');
      sh.getRange(fila, RESP_I.APROB_SUPER + 1).setValue(emailUsuario);
      sh.getRange(fila, RESP_I.FECHA_APROB_SUPER + 1).setValue(new Date());
      
      Logger.log("✅ Aprobado Supervisor. Pasando a Pendiente Director.");
      return { status: 'ok', message: 'Aprobado correctamente (Nivel Supervisor -> Nivel Director).' };
      
    } else if (estadoNormalizado === 'pendiente director') {
      // Segundo paso: Director -> Aprobado
      // Validamos que solo Director llegue aquí (Ya hecho arriba por el IF del supervisor, pero doble check):
      if (rolUsuario !== 'director') {
         Logger.log("❌ NEGADO: Supervisor intentando aprobar nivel Director.");
         return { status: 'error', message: 'Error Crítico: Solo el Director Nacional puede dar la aprobación final.' };
      }

      sh.getRange(fila, RESP_I.ESTADO + 1).setValue('Aprobado');
      sh.getRange(fila, RESP_I.APROB_DIR + 1).setValue(emailUsuario);
      sh.getRange(fila, RESP_I.FECHA_APROB_DIR + 1).setValue(new Date());
      
      Logger.log("✅ Aprobado Director. Registro finalizado.");
      return { status: 'ok', message: 'Aprobación final completada.' };
      
    } else if (estadoNormalizado === 'aprobado') {
      Logger.log("⚠️ Registro ya aprobado anteriormente.");
      return { status: 'error', message: 'Este registro ya se encuentra aprobado.' };
    } else {
      Logger.log("❌ Estado desconocido: " + estadoActual);
      return { status: 'error', message: 'Estado del registro inválido: ' + estadoActual + '.' };
    }

  } catch (e) {
    Logger.log("❌ Error en aprobarHorasExtras: " + e.toString());
    return { status: 'error', message: 'Error interno: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// 3. CONSULTA Y EXPORTACIÓN
// ===================================================================

function obtenerRegistros(filtros) {
  if (!filtros) filtros = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { status: 'error', registros: [] };
    asegurarEncabezadosRespuestas_(sh);
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return { status: 'ok', registros: [] };
    const tz = TZ;
    
    const fInicio = filtros.fechaInicio ? new Date(filtros.fechaInicio + 'T00:00:00') : null;
    const fFin = filtros.fechaFin ? new Date(filtros.fechaFin + 'T23:59:59') : null;
    const fCedula = (filtros.cedula || '').toLowerCase();
    const fNombre = (filtros.nombre || '').toLowerCase();
    const fCentro = (filtros.centro || '').toLowerCase();

    const registros = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || !row[RESP_I.CEDULA]) continue;
      
      // --- FECHAS ---
      let fechaEntradaStr = '';
      const fechaEntradaRaw = row[RESP_I.FECHA_ENT];
      if (fechaEntradaRaw instanceof Date && !isNaN(fechaEntradaRaw.getTime())) {
        fechaEntradaStr = Utilities.formatDate(fechaEntradaRaw, tz, "dd/MM/yyyy");
      } else if (fechaEntradaRaw) {
        const strTemp = String(fechaEntradaRaw).trim();
        if (strTemp.includes('Mon') || strTemp.includes('Sun') || strTemp.includes('1899')) {
           const tempDate = new Date(strTemp);
           if (!isNaN(tempDate.getTime()) && tempDate.getFullYear() > 1900) fechaEntradaStr = Utilities.formatDate(tempDate, tz, "dd/MM/yyyy");
           else fechaEntradaStr = strTemp;
        } else { fechaEntradaStr = strTemp; }
      }

      let horaEntradaStr = '';
      const horaEntradaRaw = row[RESP_I.HORA_ENT];
      if (horaEntradaRaw instanceof Date) horaEntradaStr = Utilities.formatDate(horaEntradaRaw, tz, "HH:mm:ss");
      else if (horaEntradaRaw) {
         const matchHora = String(horaEntradaRaw).match(/(\d{1,2}:\d{2}(:\d{2})?)/);
         horaEntradaStr = matchHora ? matchHora[1] : String(horaEntradaRaw);
      }
      if (!horaEntradaStr) horaEntradaStr = '00:00:00';
      
      let tsEntrada = null;
      if (fechaEntradaStr) {
        const parts = fechaEntradaStr.split('/');
        if (parts.length === 3 && parts[2].length === 4) tsEntrada = new Date(parts[2] + '-' + parts[1] + '-' + parts[0] + 'T' + horaEntradaStr);
      }
      if (!tsEntrada || isNaN(tsEntrada)) continue; 

      const tiempoRegistro = tsEntrada.getTime();
      if (fInicio && tiempoRegistro < fInicio.getTime()) continue;
      if (fFin && tiempoRegistro > fFin.getTime()) continue;

      // --- FILTROS TEXTO ---
      const cedulaVal = String(row[RESP_I.CEDULA] || '').trim();
      const nombreVal = String(row[RESP_I.NOMBRE] || '').trim();
      const centroVal = String(row[RESP_I.CENTRO] || '').trim();

      if (fCedula && cedulaVal.toLowerCase().indexOf(fCedula) === -1) continue;
      if (fNombre && nombreVal.toLowerCase().indexOf(fNombre) === -1) continue;
      if (fCentro && centroVal.toLowerCase().indexOf(fCentro) === -1) continue;

      // --- SALIDA ---
      let horaSalidaStr = '';
      const horaSalidaRaw = row[RESP_I.HORA_SAL];
      if (horaSalidaRaw instanceof Date) horaSalidaStr = Utilities.formatDate(horaSalidaRaw, tz, "HH:mm:ss");
      else if (horaSalidaRaw) {
         const strHoraSal = String(horaSalidaRaw).trim();
         if (strHoraSal.includes('1899') || strHoraSal.includes('GMT')) {
            const matchH = strHoraSal.match(/(\d{1,2}:\d{2}(:\d{2})?)/);
            horaSalidaStr = matchH ? matchH[1] : '-';
         } else { horaSalidaStr = strHoraSal || '-'; }
      } else { horaSalidaStr = '-'; }
      
      let fechaSalidaStr = '';
      const fechaSalidaRaw = row[RESP_I.FECHA_SAL];
      if (fechaSalidaRaw instanceof Date && !isNaN(fechaSalidaRaw.getTime())) {
        fechaSalidaStr = Utilities.formatDate(fechaSalidaRaw, tz, "dd/MM/yyyy");
      } else if (fechaSalidaRaw) {
         const strTemp = String(fechaSalidaRaw).trim();
         if (strTemp.includes('Mon') || strTemp.includes('Sun')) {
            const tempDate = new Date(strTemp);
            if (!isNaN(tempDate.getTime()) && tempDate.getFullYear() > 1900) fechaSalidaStr = Utilities.formatDate(tempDate, tz, "dd/MM/yyyy");
            else fechaSalidaStr = strTemp;
         } else { fechaSalidaStr = strTemp; }
      }

      let lat = row[RESP_I.LAT];
      let lng = row[RESP_I.LNG];
      if (lat instanceof Date) lat = lat.getTime().toString();
      if (lng instanceof Date) lng = lng.getTime().toString();

      // --- DATOS DE APROBACIÓN (NUEVO) ---
      const estadoHE = String(row[RESP_I.ESTADO] || '').trim();
      const aprobSuper = String(row[RESP_I.APROB_SUPER] || '').trim();
      const aprobDir = String(row[RESP_I.APROB_DIR] || '').trim();
      // Cálculos de horas (leer lo que haya, si está vacío es 0)
      const totalHE = row[RESP_I.TOTAL_HE] || 0;
      const totalNoct = row[RESP_I.TOTAL_NOCT] || 0;

      const reg = {
        fila: i + 1, // IMPORTANTE: Enviamos el número de fila para poder actualizarlo después
        timestamp: tsEntrada.toISOString(),
        timestampRaw: tiempoRegistro,
        cedula: cedulaVal,
        nombre: nombreVal,
        centro: centroVal,
        ciudad: String(row[RESP_I.CIUDAD] || '').trim(),
        fotoUrl: String(row[RESP_I.FOTO] || '').trim(),
        dentroCentro: String(row[RESP_I.DENTRO] || 'No').trim(),
        distancia: row[RESP_I.DISTANCIA],
        accuracy: row[RESP_I.ACCURACY],
        lat: lat,
        lng: lng,
        dirGeo: String(row[RESP_I.DIR_GEO] || '').trim(),
        observaciones: String(row[RESP_I.OBS] || '').trim(),
        fechaEntrada: fechaEntradaStr,
        horaEntrada: horaEntradaStr,
        fotoEntrada: String(row[RESP_I.FOTO_ENT] || '').trim(),
        fechaSalida: fechaSalidaStr,
        horaSalida: horaSalidaStr,
        fotoSalida: String(row[RESP_I.FOTO_SAL] || '').trim(),
        dentroCentroSal: String(row[RESP_I.DENTRO_SAL] || '').trim() || '-',
        // Campos nuevos para el Frontend
        estadoHE: estadoHE,
        totalHorasExtras: totalHE,
        totalNocturnas: totalNoct
      };
      registros.push(reg);
    }
    registros.sort(function(a, b) { return b.timestampRaw - a.timestampRaw; });
    return { status: 'ok', registros: registros };
  } catch (e) {
    Logger.log('Error en obtenerRegistros: ' + e.toString());
    return { status: 'error', message: 'Error interno: ' + e.toString() };
  }
}

function exportarRegistrosExcel(filtros) {
  if (!filtros) filtros = {};
  const resultado = obtenerRegistros(filtros);
  if (resultado.status !== 'ok') return { status: 'error', message: 'Error datos' };
  
  let csv = 'Cedula,Nombre,Centro,Ciudad,Entrada,Salida,Total HE,Total Noct,Estado,Aprob Supervisor,Aprob Director\n';

  const escape = function(str) { return '"' + String(str || '').replace(/"/g, '""') + '"'; };

  for (let i = 0; i < resultado.registros.length; i++) {
    const r = resultado.registros[i];
    const csvLinea = [
      escape(r.cedula), escape(r.nombre), escape(r.centro), escape(r.ciudad),
      escape(r.fechaEntrada + ' ' + r.horaEntrada),
      escape(r.fechaSalida && r.fechaSalida !== '-' ? r.fechaSalida + ' ' + r.horaSalida : '-'),
      escape(r.totalHorasExtras), escape(r.totalNocturnas),
      escape(r.estadoHE), escape(''), escape('') // Emails de aprobación se pueden agregar si se desea exportar
    ].join(',') + '\n';
    csv += csvLinea;
  }
  
  return {
    status: 'ok',
    filename: 'Reporte_NASE_HE_' + Date.now() + '.csv',
    csvContent: csv
  };
}

// -------------------------------------------------------------------
// 4. MÓDULO EMPLEADOS (Sin cambios mayores, solo revisión rápida)
// -------------------------------------------------------------------
function actualizarCacheEmpleados() {
  try {
    const ssBase = SpreadsheetApp.openById(ID_LIBRO_BASE);
    const hoja = ssBase.getSheetByName('BASE OPERATIVA');
    if (!hoja) return "Error: Hoja BASE OPERATIVA no encontrada";
    const data = hoja.getDataRange().getValues();
    if (data.length < 2) return "Error: Sin datos";
    const headers = data[0].map(function(h) { return String(h).toUpperCase().trim(); });
    const idxEstado = headers.findIndex(function(h) { return h.includes('ESTADO'); });
    const idxCedula = headers.findIndex(function(h) { return h.includes('DOCUMENTO') || h.includes('IDENTIDAD'); });
    const idxNombre = headers.findIndex(function(h) { return h.includes('NOMBRE'); });
    const idxCargo = headers.findIndex(function(h) { return h.includes('CARGO'); });
    const idxCentro = headers.findIndex(function(h) { return h.includes('SUBCENTRO') || h.includes('CENTRO'); });
    const idxCiudad = headers.findIndex(function(h) { return h.includes('CIUDAD'); });
    if (idxCedula === -1 || idxNombre === -1) return "Error Columnas";
    const empleados = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const estado = idxEstado > -1 ? String(row[idxEstado] || '').toUpperCase().trim() : 'A';
      const doc = String(row[idxCedula] || '').replace(/\D/g, '').trim();
      if (doc) {
        if (estado === 'A' || estado === 'ACTIVO' || estado === '') {
          const nombre = String(row[idxNombre] || '').trim();
          const centro = idxCentro > -1 ? String(row[idxCentro] || '').trim() : '';
          const ciudad = idxCiudad > -1 ? String(row[idxCiudad] || '').trim() : '';
          const cargo = idxCargo > -1 ? String(row[idxCargo] || '').trim() : '';
          const tipo = cargo.toUpperCase().includes('SUPERNUMERARIO') ? 'super' : 'fijo';
          empleados[doc] = { nombre: nombre, centro: centro, cargo: cargo, tipo: tipo, ciudad: ciudad };
        }
      }
    }
    const json = JSON.stringify(empleados);
    const cache = CacheService.getScriptCache();
    const totalChunks = Math.ceil(json.length / CHUNK_SIZE);
    cache.put('empleadosBase_chunks', String(totalChunks), CACHE_DURATION);
    for (let i = 0; i < totalChunks; i++) {
      cache.put('empleadosBase_' + i, json.substr(i * CHUNK_SIZE, CHUNK_SIZE), CACHE_DURATION);
    }
    empleadosCacheRAM = empleados;
    return "OK";
  } catch (e) { return "Error: " + e.toString(); }
}

function buscarEmpleadoPorCedula(cedula) {
  if (!cedula) return { ok: false };
  const cedulaLimpia = String(cedula).replace(/\D/g, '').trim();
  if (empleadosCacheRAM && empleadosCacheRAM[cedulaLimpia]) {
    const emp = empleadosCacheRAM[cedulaLimpia];
    return { ok: true, nombre: emp.nombre, centro: emp.centro, cargo: emp.cargo, tipo: emp.tipo, ciudad: emp.ciudad };
  }
  try {
    const cache = CacheService.getScriptCache();
    const totalChunks = parseInt(cache.get('empleadosBase_chunks') || '0', 10);
    if (totalChunks > 0) {
      let json = '';
      for (let i = 0; i < totalChunks; i++) {
        const chunk = cache.get('empleadosBase_' + i);
        if (!chunk) throw new Error("Cache incompleta");
        json += chunk;
      }
      empleadosCacheRAM = JSON.parse(json);
      if (empleadosCacheRAM[cedulaLimpia]) return { ok: true, ...empleadosCacheRAM[cedulaLimpia] };
    }
  } catch (err) {}
  return buscarEmpleadoPorCedulaEnLibro(cedulaLimpia);
}

function buscarEmpleadoPorCedulaEnLibro(cedula) {
  try {
    const ssBase = SpreadsheetApp.openById(ID_LIBRO_BASE);
    const hoja = ssBase.getSheetByName('BASE OPERATIVA');
    if (!hoja) return { ok: false, nombre: 'Error hoja' };
    const finderExact = hoja.createTextFinder(cedula).matchEntireCell(true);
    let results = finderExact.findAll();
    if (!results || results.length === 0) {
      const finderLoose = hoja.createTextFinder(cedula).matchEntireCell(false);
      results = finderLoose.findAll();
    }
    if (!results || results.length === 0) return { ok: false, nombre: 'No encontrado' };
    const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(function(h) { return String(h).toUpperCase().trim(); });
    const idxEstado = headers.findIndex(function(h) { return h.includes('ESTADO'); });
    const idxCedula = headers.findIndex(function(h) { return h.includes('DOCUMENTO') || h.includes('IDENTIDAD'); });
    const idxNombre = headers.findIndex(function(h) { return h.includes('NOMBRE'); });
    const idxCargo = headers.findIndex(function(h) { return h.includes('CARGO'); });
    const idxCentro = headers.findIndex(function(h) { return h.includes('SUBCENTRO') || h.includes('CENTRO'); });
    const idxCiudad = headers.findIndex(function(h) { return h.includes('CIUDAD'); });
    if (idxCedula === -1 || idxNombre === -1) return { ok: false, nombre: 'Estructura incorrecta' };
    results.sort(function(a, b) { return b.getRow() - a.getRow(); });
    for (let i = 0; i < results.length; i++) {
      const row = results[i].getRow();
      const rowValues = hoja.getRange(row, 1, 1, hoja.getLastColumn()).getValues()[0];
      const docEnFila = String(rowValues[idxCedula] || '').replace(/\D/g, '');
      if (docEnFila === cedula) {
        const estado = idxEstado > -1 ? String(rowValues[idxEstado] || '').toUpperCase().trim() : 'A';
        if (estado === 'A' || estado === 'ACTIVO' || estado === '') {
           const cargo = idxCargo > -1 ? String(rowValues[idxCargo] || '').trim() : '';
           return { ok: true, nombre: String(rowValues[idxNombre] || '').trim(), centro: idxCentro > -1 ? String(rowValues[idxCentro] || '').trim() : '', cargo: cargo, ciudad: idxCiudad > -1 ? String(rowValues[idxCiudad] || '').trim() : '', tipo: cargo.toUpperCase().includes('SUPERNUMERARIO') ? 'super' : 'fijo' };
        }
      }
    }
    return { ok: false, nombre: 'No encontrado' };
  } catch (e) { return { ok: false, nombre: 'Error DB' }; }
}

// -------------------------------------------------------------------
// 5. CENTROS
// -------------------------------------------------------------------
// ===================================================================
// 2. DATOS DE CENTROS (Leer de Libro Externo)
// ===================================================================

/**
 * @summary Obtiene los datos de centros del libro principal.
 * @description En lugar de buscar en la hoja actual, abre el libro
 *              externo (ID_LIBRO_CENTROS) y lee la hoja "Centros".
 *              Esto evita duplicar hojas y permite centralizar los datos.
 * 
 * @returns {Object} Estructura JSON para el formulario.
 */
// ===================================================================
// 2. DATOS DE CENTROS (Leer de Libro Externo) - CORREGIDO
// ===================================================================

/**
 * @summary Obtiene los datos de centros del libro externo con coordenadas.
 * @description Lee Ciudad, Centro, LAT REF, LNG REF y RADIO del libro de centros.
 * @returns {Object} Estructura JSON con coordenadas para el formulario.
 */
function getCentrosData() {
  // ID del libro donde están los Centros
  const ID_LIBRO_CENTROS = "1PchIxXq617RRL556vHui4ImG7ms2irxiY3fPLIoqcQc"; 
  
  try {
    // 1. Abrir Libro Externo
    const ssExterno = SpreadsheetApp.openById(ID_LIBRO_CENTROS);
    
    // 2. Buscar la hoja "Centros" (como indicaste que se llama)
    let hojaCentros = ssExterno.getSheetByName("Centros");
    if (!hojaCentros) hojaCentros = ssExterno.getSheetByName("BASE_CENTROS");
    
    if (!hojaCentros) {
      Logger.log("⚠️ No se encontró la hoja 'Centros' en el libro externo.");
      return { structured: {}, data: [] };
    }

    // 3. Leer Datos
    const data = hojaCentros.getDataRange().getValues();
    if (!data || data.length < 2) {
      Logger.log("⚠️ Hoja Centros vacía o sin datos.");
      return { structured: {}, data: [] };
    }

    // 4. Mapear encabezados (convertir a mayúsculas para comparar)
    const headers = data[0].map(function(h) { 
      return String(h).toUpperCase().trim(); 
    });
    const rows = data.slice(1);

    // 5. Buscar índices de columnas
    // Columnas esperadas: Ciudad, Centro, LAT REF, LNG REF, RADIO
    const idxCiudad = headers.findIndex(function(h) { return h === "CIUDAD"; });
    const idxCentro = headers.findIndex(function(h) { return h === "CENTRO"; });
    const idxLat = headers.findIndex(function(h) { 
      return h === "LAT REF" || h === "LAT" || h === "LATITUD"; 
    });
    const idxLng = headers.findIndex(function(h) { 
      return h === "LNG REF" || h === "LNG" || h === "LONGITUD" || h === "LON"; 
    });
    const idxRadio = headers.findIndex(function(h) { 
      return h === "RADIO" || h === "RADIO_M" || h === "METROS"; 
    });

    // Log para debugging
    Logger.log("📊 Encabezados encontrados: " + headers.join(" | "));
    Logger.log("📍 Índices -> Ciudad:" + idxCiudad + " Centro:" + idxCentro + 
               " Lat:" + idxLat + " Lng:" + idxLng + " Radio:" + idxRadio);

    if (idxCiudad === -1 || idxCentro === -1) {
      Logger.log("❌ Error crítico: No se encontraron columnas 'Ciudad' o 'Centro'.");
      return { structured: {}, data: [] };
    }

    if (idxLat === -1 || idxLng === -1) {
      Logger.log("❌ Error crítico: No se encontraron columnas 'LAT REF' o 'LNG REF'.");
      return { structured: {}, data: [] };
    }

    // 6. Generar objeto estructurado para el formulario
    // El frontend espera: { "clave": { ciudad, centro, lat, lng, radio }, ... }
    const structuredData = {};
    let countValid = 0;
    let countInvalid = 0;
    
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const ciudad = row[idxCiudad] ? String(row[idxCiudad]).trim() : "";
      const centro = row[idxCentro] ? String(row[idxCentro]).trim() : "";
      
      // Parsear coordenadas (manejar comas decimales)
      let lat = 0, lng = 0, radio = 30; // Radio por defecto: 30 metros
      
      if (idxLat !== -1 && row[idxLat] !== undefined && row[idxLat] !== "") {
        const latStr = String(row[idxLat]).replace(',', '.');
        lat = parseFloat(latStr);
      }
      
      if (idxLng !== -1 && row[idxLng] !== undefined && row[idxLng] !== "") {
        const lngStr = String(row[idxLng]).replace(',', '.');
        lng = parseFloat(lngStr);
      }
      
      if (idxRadio !== -1 && row[idxRadio] !== undefined && row[idxRadio] !== "") {
        radio = parseInt(row[idxRadio]) || 30;
      }
      
      // Validar que tenga datos mínimos necesarios
      if (ciudad && centro && !isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0) {
        // Crear clave única para evitar duplicados
        const key = ciudad.toUpperCase().replace(/\s+/g, '_') + '|' + 
                    centro.toUpperCase().replace(/\s+/g, '_');
        
        structuredData[key] = { 
          ciudad: ciudad, 
          centro: centro,
          lat: lat,
          lng: lng,
          radio: radio
        };
        countValid++;
      } else {
        countInvalid++;
        // Log para debugging de filas problemáticas
        if (i < 5) { // Solo logear las primeras 5 inválidas
          Logger.log("⚠️ Fila " + (i+2) + " inválida: Ciudad='" + ciudad + 
                     "' Centro='" + centro + "' Lat=" + lat + " Lng=" + lng);
        }
      }
    }

    // 7. Log resumen
    Logger.log("✅ Centros cargados: " + countValid + " válidos, " + countInvalid + " omitidos.");
    
    // Log de muestra (primeros 3 centros)
    const keys = Object.keys(structuredData);
    for (let j = 0; j < Math.min(3, keys.length); j++) {
      Logger.log("   📍 " + keys[j] + " -> " + JSON.stringify(structuredData[keys[j]]));
    }
    
    return { structured: structuredData, data: data };

  } catch (e) {
    Logger.log("❌ Error en getCentrosData: " + e.toString());
    return { structured: {}, data: [] };
  }
}

// ===================================================================
// 6. REGISTRO (SECUENCIA + GEO) - ADAPTADO PARA GUARDAR CÁLCULOS
// ===================================================================

function registrarUltra(dataInput) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(4000)) return { status: 'error', message: 'Sistema ocupado, intenta de nuevo.' };
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    asegurarEncabezadosRespuestas_(sh);
    
    const cedulaLimpia = String(dataInput.cedula1 || '').replace(/\D/g, '');
    const tipoIntentoRaw = String(dataInput.tipo || '').trim();
    const tipoActual = tipoNormalizado(tipoIntentoRaw);
    
    if (!cedulaLimpia) return { status: 'error', message: 'Cédula inválida.' };
    if (!tipoActual) return { status: 'error', message: 'Tipo inválido (Entrada/Salida).' };

    const seq = validarSecuenciaRapida(sh, cedulaLimpia, tipoActual);
    if (!seq.esValido) return { status: 'error', message: seq.message };

    let estaDentro = false;
    let distanciaReal = 0;
    const centrosInfo = getCentrosData();
    const key = normaliza(dataInput.ciudad) + '|' + normaliza(dataInput.centro);
    if (centrosInfo.structured[key]) {
      const centroData = centrosInfo.structured[key];
      const latUser = parseFloat(String(dataInput.lat || '').replace(',', '.'));
      const lngUser = parseFloat(String(dataInput.lng || '').replace(',', '.'));
      if (!isNaN(latUser) && !isNaN(lngUser)) {
        distanciaReal = calcularDistanciaHaversine(latUser, lngUser, centroData.lat, centroData.lng);
        estaDentro = distanciaReal <= centroData.radio;
      }
    }

    let nombreFinal = String(dataInput.nombre || '').trim();
    if (!nombreFinal || nombreFinal.includes('No encontrado')) {
      const emp = buscarEmpleadoPorCedula(cedulaLimpia);
      nombreFinal = emp.ok ? emp.nombre : 'NO ENCONTRADO EN BASE';
    }

    const now = new Date();
    const fechaStr = Utilities.formatDate(now, TZ, "dd/MM/yyyy");
    const horaStr = Utilities.formatDate(now, TZ, "HH:mm");
    const fotoEvento = String(dataInput.fotoUrl || '').trim();

    if (tipoActual === 'salida') {
      // AL REGISTRAR SALIDA, DEBERÍAMOS CALCULAR LAS HORAS EXTRAS
      const actualizado = _actualizarUltimaEntradaConSalida_(sh, cedulaLimpia, fechaStr, horaStr, fotoEvento, estaDentro ? 'Sí' : 'No');
      if (actualizado) return { status: 'ok', message: 'Salida registrada exitosamente.' };
      return { status: 'error', message: 'No se encontró entrada abierta para cerrar. Registra ENTRADA primero.' };
    }

    let fila = new Array(sh.getLastColumn());
    fila[RESP_I.CEDULA] = cedulaLimpia; 
    fila[RESP_I.CENTRO] = dataInput.centro || '';
    fila[RESP_I.CIUDAD] = dataInput.ciudad || '';
    fila[RESP_I.LAT] = toFixed5(dataInput.lat); 
    fila[RESP_I.LNG] = toFixed5(dataInput.lng); 
    fila[RESP_I.ACEPTO] = dataInput.acepto ? 'Sí' : 'No';
    fila[RESP_I.CIUDAD_GEO] = ''; 
    fila[RESP_I.DIR_GEO] = '';
    fila[RESP_I.ACCURACY] = '';
    fila[RESP_I.DENTRO] = estaDentro ? 'Sí' : 'No';
    fila[RESP_I.DISTANCIA] = Math.round(distanciaReal); 
    fila[RESP_I.OBS] = "Registro Web";
    fila[RESP_I.NOMBRE] = nombreFinal; 
    fila[RESP_I.FOTO] = fotoEvento;
    
    fila[RESP_I.FECHA_ENT] = fechaStr;
    fila[RESP_I.HORA_ENT] = horaStr;
    fila[RESP_I.FOTO_ENT] = fotoEvento;
    
    fila[RESP_I.FECHA_SAL] = ""; 
    fila[RESP_I.HORA_SAL] = "";
    fila[RESP_I.FOTO_SAL] = "";
    fila[RESP_I.DENTRO_SAL] = estaDentro ? 'Sí' : 'No';
    
    // Inicializar campos de Horas Extras vacíos
    fila[RESP_I.TOTAL_HE] = "";
    fila[RESP_I.TOTAL_NOCT] = "";
    fila[RESP_I.ESTADO] = "Pendiente Supervisor"; // Estado inicial

    sh.appendRow(fila);
    return { status: 'ok', message: 'Entrada registrada exitosamente.' };
  } catch (e) {
    return { status: 'error', message: 'Error interno: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function _actualizarUltimaEntradaConSalida_(sheet, cedula, fechaSalida, horaSalida, fotoSalida, dentroSalida) {
  const finder = sheet.createTextFinder(cedula).matchEntireCell(true);
  const results = finder.findAll();
  if (!results || results.length === 0) return false;
  results.sort(function(a, b) { return b.getRow() - a.getRow(); });

  for (let i = 0; i < results.length; i++) {
    const r = results[i].getRow();
    if (r <= 1) continue;
    const colFechaSal = RESP_I.FECHA_SAL + 1;
    const colHoraSal = RESP_I.HORA_SAL + 1;
    const colFotoSal = RESP_I.FOTO_SAL + 1;
    const colDentroSal = RESP_I.DENTRO_SAL + 1;

    const yaTieneSalida = String(sheet.getRange(r, colFechaSal).getValue() || '').trim() !== '';
    if (yaTieneSalida) continue;

    sheet.getRange(r, colFechaSal).setValue(fechaSalida);
    sheet.getRange(r, colHoraSal).setValue(horaSalida);
    sheet.getRange(r, colFotoSal).setValue(fotoSalida || '');
    sheet.getRange(r, colDentroSal).setValue(dentroSalida || '');
    
    // AQUÍ IRÍA EL LLAMADO A FUNCIÓN DE CÁLCULO DE HORAS
    // calcularYGuardarHorasExtras(sheet, r); 
    // (Haremos esto cuando veamos el archivo de cálculo que mencionas)
    
    return true;
  }
}

function validarSecuenciaRapida(sheet, cedula, tipoActual) {
  const finder = sheet.createTextFinder(cedula).matchEntireCell(true);
  const results = finder.findAll();
  if (!results || results.length === 0) {
    if (tipoActual !== 'entrada') return { esValido: false, message: 'Tu primer registro debe ser una ENTRADA.', tipoSugerido: 'entrada' };
    return { esValido: true };
  }
  results.sort(function(a, b) { return a.getRow() - b.getRow(); });
  const lastRow = results[results.length - 1].getRow();
  const colFechaSal = RESP_I.FECHA_SAL + 1;
  const colHoraSal = RESP_I.HORA_SAL + 1;
  const fechaSalidaExistente = String(sheet.getRange(lastRow, colFechaSal).getValue() || '').trim();
  const horaSalidaExistente = String(sheet.getRange(lastRow, colHoraSal).getValue() || '').trim();

  if (fechaSalidaExistente && horaSalidaExistente) {
    if (tipoActual === 'salida') return { esValido: false, message: 'Tu último turno ya está cerrado. Debes registrar ENTRADA primero.', tipoSugerido: 'entrada' };
    return { esValido: true };
  }
  if (!fechaSalidaExistente && tipoActual === 'entrada') return { esValido: false, message: 'Tienes una ENTRADA abierta. Debes registrar SALIDA.', tipoSugerido: 'salida' };
  return { esValido: true };
}

function validarSecuenciaFront(cedula, tipoIntento) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { esValido: false, message: 'Hoja no encontrada', tipoSugerido: 'entrada' };
    asegurarEncabezadosRespuestas_(sh);
    const cedulaLimpia = String(cedula || '').replace(/\D/g, '').trim();
    const tipoNorm = tipoIntento.includes('ent') ? 'entrada' : 'salida';
    const finder = sh.createTextFinder(cedulaLimpia).matchEntireCell(true);
    const results = finder.findAll();
    if (!results || results.length === 0) {
      if (tipoNorm !== 'entrada') return { esValido: false, message: 'Tu primer registro debe ser una ENTRADA.', tipoSugerido: 'entrada' };
      return { esValido: true };
    }
    results.sort(function(a, b) { return a.getRow() - b.getRow(); });
    const lastRow = results[results.length - 1].getRow();
    const lastRowData = sh.getRange(lastRow, 1, 1, sh.getLastColumn()).getValues()[0];
    const colFechaSal = RESP_I.FECHA_SAL;
    const colHoraSal = RESP_I.HORA_SAL;
    const fechaSalidaExistente = String(lastRowData[colFechaSal] || '').trim();
    const horaSalidaExistente = String(lastRowData[colHoraSal] || '').trim();
    if (fechaSalidaExistente && horaSalidaExistente) {
      if (tipoNorm === 'salida') return { esValido: false, message: 'Tu último turno ya está cerrado. Debes registrar ENTRADA primero.', tipoSugerido: 'entrada' };
      return { esValido: true };
    }
    if (!fechaSalidaExistente && tipoNorm === 'entrada') return { esValido: false, message: 'Tienes una ENTRADA abierta. Debes registrar SALIDA.', tipoSugerido: 'salida' };
    return { esValido: true };
  } catch (e) { return { esValido: false, message: 'Error interno: ' + e.toString(), tipoSugerido: null }; }
}

function tipoNormalizado(v) {
  const s = String(v || '').trim().toLowerCase();
  if (s.includes('ent')) return 'entrada';
  if (s.includes('sal')) return 'salida';
  return '';
}

function obtenerUltimoTipoRegistro(cedula) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { ok: true, ultimoTipo: null, tipoSugerido: 'entrada' };
    const ced = String(cedula).replace(/\D/g, '');
    const finder = sh.createTextFinder(ced).matchEntireCell(true);
    const results = finder.findAll();
    if (!results || results.length === 0) return { ok: true, ultimoTipo: null, tipoSugerido: 'entrada', mensaje: 'Tu primer registro debe ser una ENTRADA' };
    results.sort(function(a, b) { return a.getRow() - b.getRow(); });
    const lastRow = results[results.length - 1].getRow();
    const lastRowData = sh.getRange(lastRow, 1, 1, sh.getLastColumn()).getValues()[0];
    const colFechaSal = RESP_I.FECHA_SAL;
    const colHoraSal = RESP_I.HORA_SAL;
    const fechaSalidaVal = String(lastRowData[colFechaSal] || '').trim();
    const horaSalidaVal = String(lastRowData[colHoraSal] || '').trim();
    if (fechaSalidaVal && horaSalidaVal) return { ok: true, ultimoTipo: 'Salida', tipoSugerido: 'entrada', mensaje: 'Tu último registro fue Salida. Ahora debes registrar Entrada.' };
    else return { ok: true, ultimoTipo: 'Entrada', tipoSugerido: 'salida', mensaje: 'Tu último registro fue Entrada. Ahora debes registrar Salida.' };
  } catch (e) { return { ok: false, ultimoTipo: null, tipoSugerido: null }; }
}

function subirFoto(fotoBase64, cedula) {
  if (!fotoBase64) return '';
  try {
    const folder = DriveApp.getFolderById(ID_CARPETA_FOTOS);
    const base64 = String(fotoBase64).includes(',') ? fotoBase64.split(',')[1] : fotoBase64;
    const imageBytes = Utilities.base64Decode(base64);
    const fileName = 'foto_' + (cedula || 'temp') + '_' + Date.now() + '.jpg';
    const file = folder.createFile(Utilities.newBlob(imageBytes, 'image/jpeg', fileName));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) { return ''; }
}

function calcularDistanciaHaversine(lat1, lon1, lat2, lon2) {
  const R = 6371e3;
  const φ1 = lat1 * Math.PI/180;
  const φ2 = lat2 * Math.PI/180;
  const Δφ = (lat2-lat1) * Math.PI/180;
  const Δλ = (lon2-lon1) * Math.PI/180;
  const a = Math.sin(Δφ/2) * Math.sin(Δφ/2) + Math.cos(φ1) * Math.cos(φ2) * Math.sin(Δλ/2) * Math.sin(Δλ/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

function toFixed5(v) { return v ? Number(Number(String(v).replace(',', '.')).toFixed(5)) : ''; }
function normaliza(txt) { return String(txt || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase(); }

function obtenerSugerencias(query, tipo) {
  if (!query || query.length < 2) return [];
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let data = [];
  if (tipo === 'centro') {
    const hojaCentros = ss.getSheetByName(SHEET_CENTROS);
    if (hojaCentros) {
      const centrosData = hojaCentros.getDataRange().getValues();
      const idxCentro = centrosData[0].map(function(h) { return String(h).trim().toUpperCase(); }).findIndex(function(h) { return h === 'CENTRO' || h.includes('SEDE'); });
      if (idxCentro > -1) { for (let i=1; i<centrosData.length; i++) { if (centrosData[i][idxCentro]) data.push(String(centrosData[i][idxCentro]).trim()); } }
    }
  } else {
    const hoja = ss.getSheetByName(SHEET_NAME);
    if (!hoja) return [];
    const respData = hoja.getDataRange().getValues();
    if (respData.length <= 1) return [];
    const idx = (tipo === 'cedula') ? respData[0].indexOf('Cédula') : respData[0].indexOf('Nombre');
    if (idx > -1) { for (let i=1; i<respData.length; i++) { if (respData[i][idx]) data.push(String(respData[i][idx]).trim()); } }
  }
  const queryLower = query.toLowerCase();
  const sugerencias = [];
  const unique = {};
  for (let i = 0; i < data.length; i++) {
    const item = data[i];
    if (item.toLowerCase().includes(queryLower) && !unique[item]) { sugerencias.push(item); unique[item] = true; }
    if (sugerencias.length >= 10) break;
  }
  return sugerencias;
}

function mantenerSistemaActivo() { console.log("Sistema activo: " + new Date()); } 

// ===================================================================
// GESTIÓN DE ACCESO Y ERRORES
// ===================================================================
function generarPaginaAccesoDenegado(email, modulo) {
  var html = '<!DOCTYPE html><html><head>';
  html += '<meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1">';
  html += '<title>Acceso Denegado - NASE</title>';
  html += '<style>body{font-family:sans-serif;background:#f4f7f6;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}.card{background:#fff;padding:40px;border-radius:8px;box-shadow:0 4px 15px rgba(0,0,0,0.1);text-align:center;max-width:400px}.icon{font-size:50px;color:#dc3545;margin-bottom:20px}h1{color:#dc3545;margin-bottom:10px}p{color:#555;line-height:1.6}</style></head><body>';
  html += '<div class="card"><div class="icon">🚫</div><h1>Acceso Denegado</h1><p>No tienes permisos para acceder al módulo:</p><p><strong>' + modulo + '</strong></p>';
  html += '<hr style="border:0;border-top:1px solid #eee;margin:20px 0">';
  html += '<p style="font-size:12px;color:#888">Usuario: ' + email + '</p></div></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function limpiarCoordenadasEnRespuestas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Respuestas');
  if (!hoja) return;
  const data = hoja.getDataRange().getValues();
  const headers = data[0];
  const idxLat = headers.indexOf('Lat');
  const idxLng = headers.indexOf('Lng');
  if (idxLat === -1 || idxLng === -1) return;
  for (let i = 1; i < data.length; i++) {
    let lat = data[i][idxLat];
    let lng = data[i][idxLng];
    if (typeof lat === 'string' && lat.includes(',')) hoja.getRange(i + 1, idxLat + 1).setValue(parseFloat(lat.replace(',', '.')));
    if (typeof lng === 'string' && lng.includes(',')) hoja.getRange(i + 1, idxLng + 1).setValue(parseFloat(lng.replace(',', '.')));
  }
}
