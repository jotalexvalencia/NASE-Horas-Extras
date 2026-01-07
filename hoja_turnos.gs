// ============================================================
// 📘 hoja_turnos.gs – Motor de Cálculo para Horas Extras (NASE 2026)
// ------------------------------------------------------------
/**
 * @summary Módulo Central de Cálculo de Tiempos Trabajados.
 * @description Orquesta la generación de la hoja "Asistencia_SinValores"
 *              basándose en los registros de Entrada/Salida ("Respuestas").
 *              Este cálculo es la BASE para liquidar las Horas Extras.
 *
 * 🏗️ ARQUITECTURA:
 * - ⚡ Procesamiento por Lotes (Batching) para evitar timeout.
 * - 🧬 Disparadores (Triggers) para continuar el proceso automáticamente.
 * - 🧠 Bloqueo (Lock) para evitar conflictos de datos.
 * - 🧮 Algoritmo de Horas minuto a minuto (Diurnas, Nocturnas, Festivos).
 *
 * 📊 SALIDA (Hoja "Asistencia_SinValores"):
 * - Genera una fila por cada turno.
 * - Desglosa horas en 4 categorías (Necesario para liquidar HE):
 *   1. Horas Diurnas Normales.
 *   2. Horas Nocturnas Normales.
 *   3. Horas Diurnas Domingo/Festivo.
 *   4. Horas Nocturnas Domingo/Festivo.
 *
 * @dependencies
 * - Code.gs (Headers y buscarEmpleadoPorCedula).
 * - ConfigHorarios.gs (Inicio/Fin Nocturno).
 *
 * @author NASE Team
 * @version 4.1 (Adaptado para Control de Horas Extras)
 */

// ===================================================================
// 1. CONFIGURACIÓN DEL SISTEMA
// ===================================================================

/** @summary Nombre de la hoja origen con los registros de entrada/salida. */
const HOJA_ORIGEN = 'Respuestas';

/** @summary Nombre de la hoja destino donde se genera la asistencia. */
const HOJA_DESTINO = 'Asistencia_SinValores';

/** @summary Tamaño del lote de filas a procesar por ejecución. */
const TAMANO_LOTE = 10000;

/** @summary Nombre de la función handler del trigger de continuación. */
const ASIS_TRIGGER_HANDLER = 'continuarProcesoAsistencia';

/** @summary Propiedad donde se guarda la fila actual del proceso. */
const ASIS_PROP_LOTE_INICIO = 'ASIS_LOTE_INICIO';

/** @summary Propiedad flag "0" o "1" para saber si el proceso está corriendo. */
const ASIS_PROP_EN_CURSO = 'ASIS_EN_CURSO';

// ===================================================================
// 2. UTILIDADES DE NOTIFICACIÓN (UI Segura)
// ===================================================================

function _asisNotify_(message, title) {
  var t = title || 'Asistencia NASE';
  var msg = String(message || '');

  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(t, msg, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  } catch (e) {}

  try {
    SpreadsheetApp.getActive().toast(msg, t, 8);
    return;
  } catch (e) {}

  Logger.log('[' + t + '] ' + msg);
}

function _asisHasUi_() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

function _asisToastSafe_(message, title, seconds) {
  try {
    SpreadsheetApp.getActive().toast(String(message || ''), String(title || ''), Number(seconds || 5));
  } catch (e) {}
}

// ===================================================================
// 3. FUNCIÓN PRINCIPAL (Orquestador)
// ===================================================================

/**
 * @summary Inicia el proceso de generación de asistencia.
 * @description Prepara la hoja destino, verifica bloqueos y lanza los lotes.
 */
function generarTablaAsistenciaSinValores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDatos = ss.getSheetByName(HOJA_ORIGEN);

  if (!hojaDatos) {
    _asisNotify_('Error: No se encontró la hoja "' + HOJA_ORIGEN + '".', 'Error');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(3000)) {
    _asisNotify_('Otro proceso está usando el generador. Intenta nuevamente en unos segundos.', 'Sistema ocupado');
    return;
  }

  try {
    if (props.getProperty(ASIS_PROP_EN_CURSO) === '1') {
      _asisNotify_('Ya hay un proceso en curso. Continuará automáticamente.', 'Proceso en curso');
      _asisEnsureUniqueTrigger_(1000);
      return;
    }

    var hojaSalida = ss.getSheetByName(HOJA_DESTINO);
    if (hojaSalida) {
      hojaSalida.clearContents();
    } else {
      hojaSalida = ss.insertSheet(HOJA_DESTINO);
    }

    // Warm-up cache
    if (typeof actualizarCacheEmpleados === 'function') {
      try {
        actualizarCacheEmpleados();
      } catch (e) {
        Logger.log('ASIS: actualizarCacheEmpleados falló: ' + e);
      }
    }

    props.deleteProperty(ASIS_PROP_LOTE_INICIO);
    props.setProperty(ASIS_PROP_LOTE_INICIO, '2');
    props.setProperty(ASIS_PROP_EN_CURSO, '1');

    _asisNotify_('Proceso de cálculo iniciado. Procesando por lotes...', 'Asistencia');

    _procesarLotesAsistencia();

  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// 4. PROCESAMIENTO POR LOTES (Batching)
// ===================================================================

function _procesarLotesAsistencia() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var origen = ss.getSheetByName(HOJA_ORIGEN);
  var destino = ss.getSheetByName(HOJA_DESTINO);

  var inicio = parseInt(props.getProperty(ASIS_PROP_LOTE_INICIO) || '2', 10);
  var total = origen.getLastRow();

  if (inicio > total) {
    _finalizarProcesoAsistencia();
    return;
  }

  var fin = Math.min(inicio + TAMANO_LOTE - 1, total);

  var rango = origen.getRange(inicio, 1, fin - inicio + 1, origen.getLastColumn());
  var datos = rango.getValues();
  var tz = ss.getSpreadsheetTimeZone();

  var datosProcesados = _procesarDatosAsistencia(datos, tz);

  if (datosProcesados.length > 0) {
    var inicioEscritura = destino.getLastRow() > 0 ? destino.getLastRow() + 1 : 1;

    if (inicioEscritura === 1) {
      var encabezados = _generarEncabezados();
      destino.getRange(inicioEscritura, 1, 1, encabezados.length).setValues([encabezados]);
      inicioEscritura++;
    }

    destino.getRange(inicioEscritura, 1, datosProcesados.length, datosProcesados[0].length)
      .setValues(datosProcesados);
  }

  props.setProperty(ASIS_PROP_LOTE_INICIO, String(fin + 1));

  var porcentaje = ((fin / total) * 100).toFixed(1);
  Logger.log('📦 Procesadas ' + fin + '/' + total + ' filas (' + porcentaje + '%)');

  _asisToastSafe_('Procesando: ' + porcentaje + '% completado', 'Cálculo de Horas', 5);

  _asisEnsureUniqueTrigger_(1000);
}

function continuarProcesoAsistencia() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    _asisClearTriggers_();
    _procesarLotesAsistencia();
  } finally {
    lock.releaseLock();
  }
}

function _finalizarProcesoAsistencia() {
  var props = PropertiesService.getScriptProperties();

  props.deleteProperty(ASIS_PROP_LOTE_INICIO);
  props.deleteProperty(ASIS_PROP_EN_CURSO);

  _asisClearTriggers_();

  Logger.log('✅ Cálculo de horas completado.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_DESTINO);
  var ultimaFila = hoja.getLastRow();

  Logger.log('Filas calculadas en Asistencia_SinValores: ' + (ultimaFila - 1));

  if (ultimaFila <= 1) {
    Logger.log('⚠️ Advertencia: No se crearon registros.');
  }

  _asisNotify_('✅ Cálculo completado. Lista para liquidar Horas Extras.', 'Asistencia');
}

// ===================================================================
// 5. GESTIÓN DE TRIGGERS
// ===================================================================

function _asisGetTriggersByHandler_(handler) {
  return ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === handler;
  });
}

function _asisClearTriggers_() {
  var ts = _asisGetTriggersByHandler_(ASIS_TRIGGER_HANDLER);
  ts.forEach(function(t) {
    try {
      ScriptApp.deleteTrigger(t);
    } catch (e) {}
  });
}

function _asisEnsureUniqueTrigger_(delayMs) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return;

  try {
    var existentes = _asisGetTriggersByHandler_(ASIS_TRIGGER_HANDLER);

    if (existentes.length > 0) return;

    ScriptApp.newTrigger(ASIS_TRIGGER_HANDLER)
      .timeBased()
      .after(Math.max(300, Number(delayMs) || 500))
      .create();

  } catch (e) {
    Logger.log('ASIS: no se pudo crear el trigger: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// 6. GENERACIÓN DE ENCABEZADOS
// ===================================================================

function _generarEncabezados() {
  return [
    "Cédula",                           // 0
    "Nombre Empleado",                  // 1
    "Centro",                           // 2
    "Rango de Fecha",                   // 3
    "Fecha",                            // 4
    "Hora Inicio",                      // 5
    "Hora Salida",                      // 6
    "Tipo Día Inicio",                  // 7
    "Tipo Día Fin",                     // 8
    "Horas Trabajadas",                 // 9
    "Horas Diurnas Normales",           // 10
    "Horas Nocturnas Normales",         // 11
    "Horas Diurnas Domingo/Festivo",    // 12
    "Horas Nocturnas Domingo/Festivo"   // 13
  ];
}

// ===================================================================
// 7. MOTOR DE PROCESAMIENTO DE DATOS
// ===================================================================

function _procesarDatosAsistencia(datos, tz) {
  // Mapeo de índices basado en RESP_HEADERS actualizado en Code.gs
  var IDX = {
    CEDULA: 0,
    CENTRO: 1,
    FECHA_ENT: 14,
    HORA_ENT: 15,
    FECHA_SAL: 17,
    HORA_SAL: 18
  };

  var registros = [];

  for (var i = 0; i < datos.length; i++) {
    var r = datos[i];

    if (!r || r.length < 3 || !r[IDX.CEDULA]) continue;

    var cedula = String(r[IDX.CEDULA] || '').trim();
    var centroReal = String(r[IDX.CENTRO] || '').trim();

    var fechaEntrada = r[IDX.FECHA_ENT];
    var horaEntrada = r[IDX.HORA_ENT];
    var fechaSalida = r[IDX.FECHA_SAL];
    var horaSalida = r[IDX.HORA_SAL];

    if (!fechaEntrada || !horaEntrada) continue;

    var tsEntrada = _parsearFechaHora(fechaEntrada, horaEntrada, tz);
    var tsSalida = null;

    if (fechaSalida && horaSalida) {
      tsSalida = _parsearFechaHora(fechaSalida, horaSalida, tz);
    }

    if (!tsEntrada || !tsSalida || tsSalida <= tsEntrada) {
      continue;
    }

    registros.push({
      cedula: cedula,
      centroReal: centroReal,
      inicio: tsEntrada,
      fin: tsSalida,
      completo: true
    });
  }

  if (registros.length === 0) {
    Logger.log('⚠️ No se procesaron registros válidos.');
    return [];
  }

  var cacheFestivos = {};
  var filas = [];

  for (var j = 0; j < registros.length; j++) {
    var turno = registros[j];
    var ced = turno.cedula;
    var centro = turno.centroReal;

    var emp = { ok: false, nombre: 'Sin Nombre' };
    if (typeof buscarEmpleadoPorCedula === 'function') {
      emp = buscarEmpleadoPorCedula(ced);
    }
    var nombre = emp.ok ? emp.nombre : 'NO ENCONTRADO';

    var fechaInicioStr = Utilities.formatDate(turno.inicio, tz, "dd/MM/yyyy");
    var horaInicioStr = Utilities.formatDate(turno.inicio, tz, "HH:mm");
    var fechaFinStr = Utilities.formatDate(turno.fin, tz, "dd/MM/yyyy");
    var horaFinStr = Utilities.formatDate(turno.fin, tz, "HH:mm");

    var rangoCompleto = fechaInicioStr + " " + horaInicioStr + " - " + fechaFinStr + " " + horaFinStr;

    var tipoDiaInicio = _esFestivo(turno.inicio, cacheFestivos);
    var tipoDiaFin = _esFestivo(turno.fin, cacheFestivos);

    var horas = _calcularHorasPorTipo([{ inicio: turno.inicio, fin: turno.fin }], cacheFestivos);

    var fila = [
      ced,                
      nombre,             
      centro,             
      rangoCompleto,      
      fechaInicioStr,     
      horaInicioStr,      
      horaFinStr,         
      tipoDiaInicio,      
      tipoDiaFin,         
      horas.total,        
      horas.normalesDia,  
      horas.normalesNoc,  
      horas.festivosDia,  
      horas.festivosNoc   
    ];

    filas.push(fila);
  }

  return filas;
}

// ===================================================================
// 8. UTILIDADES DE PARSEO DE FECHA/HORA
// ===================================================================

function _parsearFechaHora(fechaVal, horaVal, tz) {
  try {
    var fecha = null;

    if (fechaVal instanceof Date) {
      fecha = new Date(fechaVal);
    } else if (typeof fechaVal === 'string') {
      var strFecha = fechaVal.trim();

      if (strFecha.includes('Mon') || strFecha.includes('Tue') || strFecha.includes('Wed') ||
          strFecha.includes('Thu') || strFecha.includes('Fri') || strFecha.includes('Sat') ||
          strFecha.includes('Sun') || strFecha.includes('GMT')) {
        fecha = new Date(strFecha);
      } else {
        var parts = strFecha.split('/');
        if (parts.length === 3) {
          fecha = new Date(parts[2] + '-' + parts[1] + '-' + parts[0]);
        }
      }
    }

    if (!fecha || isNaN(fecha.getTime())) {
      return null;
    }

    var hh = 0, mm = 0, ss = 0;

    if (horaVal instanceof Date) {
      hh = horaVal.getHours();
      mm = horaVal.getMinutes();
      ss = horaVal.getSeconds();
    } else if (typeof horaVal === 'string') {
      var strHora = horaVal.trim();

      var matchHora = strHora.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
      if (matchHora) {
        hh = parseInt(matchHora[1], 10);
        mm = parseInt(matchHora[2], 10);
        ss = matchHora[3] ? parseInt(matchHora[3], 10) : 0;
      }
    }

    if (isNaN(hh) || isNaN(mm)) {
      return null;
    }

    fecha.setHours(hh, mm, ss, 0);
    return fecha;

  } catch (e) {
    Logger.log('❌ Error parseando fecha/hora: ' + fechaVal + ' ' + horaVal + ' - ' + e.message);
    return null;
  }
}

function _truncarMinutos(fecha) {
  if (!fecha || !(fecha instanceof Date)) return null;
  return new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate(), fecha.getHours(), fecha.getMinutes(), 0, 0);
}

function _formatearFecha(anio, mes, dia) {
  return anio + '-' + String(mes).padStart(2, '0') + '-' + String(dia).padStart(2, '0');
}

// ===================================================================
// 9. DETECCIÓN DE TIPO DE DÍA
// ===================================================================

function _esFestivo(fecha, cacheFestivos) {
  if (!(fecha instanceof Date) || isNaN(fecha.getTime())) {
    return "Normal";
  }

  var anio = fecha.getFullYear();

  if (!cacheFestivos[anio]) {
    cacheFestivos[anio] = _generarFestivos(anio);
  }

  var mes = fecha.getMonth() + 1;
  var dia = fecha.getDate();
  var clave = anio + '-' + String(mes).padStart(2, '0') + '-' + String(dia).padStart(2, '0');

  if (cacheFestivos[anio].has(clave)) {
    return "Festivo";
  }

  if (fecha.getDay() === 0) {
    return "Domingo";
  }

  return "Normal";
}

// ===================================================================
// 10. MOTOR DE CÁLCULO DE HORAS
// ===================================================================

function _calcularHorasPorTipo(intervalos, cacheFestivos) {
  if (!Array.isArray(intervalos)) {
    intervalos = [intervalos];
  }

  var config = { horaInicio: 21, horaFin: 6 }; // Default 9PM - 6AM
  if (typeof obtenerConfiguracionHorarios === 'function') {
    try {
      config = obtenerConfiguracionHorarios();
    } catch (e) {
      Logger.log('Usando configuración nocturna por defecto');
    }
  }
  var horaInicioNoc = config.horaInicio || 21;
  var horaFinNoc = config.horaFin || 6;

  var normalesDia = 0;
  var normalesNoc = 0;
  var festivosDia = 0;
  var festivosNoc = 0;

  for (var idx = 0; idx < intervalos.length; idx++) {
    var intervalo = intervalos[idx];

    if (!intervalo || !intervalo.inicio || !intervalo.fin) continue;

    var inicio = _truncarMinutos(intervalo.inicio);
    var fin = _truncarMinutos(intervalo.fin);

    if (!inicio || !fin || isNaN(inicio.getTime()) || isNaN(fin.getTime()) || fin <= inicio) continue;

    var cursor = new Date(inicio);

    while (cursor < fin) {
      var cursorTime = cursor.getTime();
      var cursorHora = cursor.getHours();

      var nextMidnight = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1, 0, 0, 0);

      var nextInicioNoc = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate(), horaInicioNoc, 0, 0);
      if (cursorTime >= nextInicioNoc.getTime()) {
        nextInicioNoc.setDate(nextInicioNoc.getDate() + 1);
      }

      var nextFinNoc = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate(), horaFinNoc, 0, 0);
      if (cursorTime >= nextFinNoc.getTime()) {
        nextFinNoc.setDate(nextFinNoc.getDate() + 1);
      }

      var limites = [fin.getTime(), nextMidnight.getTime(), nextInicioNoc.getTime(), nextFinNoc.getTime()];
      var limitesValidos = [];
      for (var li = 0; li < limites.length; li++) {
        if (limites[li] > cursorTime) limitesValidos.push(limites[li]);
      }

      var siguienteTs = Math.min.apply(null, limitesValidos);

      if (!siguienteTs || siguienteTs <= cursorTime) {
        cursor = new Date(cursorTime + 60000);
        continue;
      }

      var siguiente = new Date(siguienteTs);

      var delta = (siguiente.getTime() - cursor.getTime()) / (1000 * 60 * 60);

      var esNocturna = false;
      if (horaInicioNoc > horaFinNoc) {
        esNocturna = (cursorHora >= horaInicioNoc || cursorHora < horaFinNoc);
      } else {
        esNocturna = (cursorHora >= horaInicioNoc && cursorHora < horaFinNoc);
      }

      var tipoDelDia = _esFestivo(cursor, cacheFestivos);

      if (tipoDelDia === "Normal") {
        if (esNocturna) {
          normalesNoc += delta;
        } else {
          normalesDia += delta;
        }
      } else {
        if (esNocturna) {
          festivosNoc += delta;
        } else {
          festivosDia += delta;
        }
      }

      cursor = siguiente;
    }
  }

  return {
    total: Number((normalesDia + normalesNoc + festivosDia + festivosNoc).toFixed(2)),
    normalesDia: Number(normalesDia.toFixed(2)),
    normalesNoc: Number(normalesNoc.toFixed(2)),
    festivosDia: Number(festivosDia.toFixed(2)),
    festivosNoc: Number(festivosNoc.toFixed(2))
  };
}

// ===================================================================
// 11. LÓGICA DE FESTIVOS COLOMBIA
// ===================================================================

function _calcularDomingoPascua(anio) {
  var a = anio % 19;
  var b = Math.floor(anio / 100);
  var c = anio % 100;
  var d = Math.floor(b / 4);
  var e = b % 4;
  var f = Math.floor((b + 8) / 25);
  var g = Math.floor((b - f + 1) / 3);
  var h = (19 * a + b - d - g + 15) % 30;
  var i = Math.floor(c / 4);
  var k = c % 4;
  var l = (32 + 2 * e + 2 * i - h - k) % 7;
  var m = Math.floor((a + 11 * h + 22 * l) / 451);
  var mes = Math.floor((h + l - 7 * m + 114) / 31);
  var dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(anio, mes - 1, dia);
}

function _generarFestivos(anio) {
  var festivos = new Set();

  var moverAlLunes = function(fecha) {
    var f = new Date(fecha);
    if (f.getDay() === 1) return f;
    var diasALunes = (8 - f.getDay()) % 7;
    if (diasALunes === 0) diasALunes = 7;
    f.setDate(f.getDate() + diasALunes);
    return f;
  };

  var agregarFestivo = function(fecha) {
    var y = fecha.getFullYear();
    var m = String(fecha.getMonth() + 1).padStart(2, '0');
    var d = String(fecha.getDate()).padStart(2, '0');
    festivos.add(y + '-' + m + '-' + d);
  };

  // 1. FESTIVOS FIJOS
  agregarFestivo(new Date(anio, 0, 1));   
  agregarFestivo(new Date(anio, 4, 1));   
  agregarFestivo(new Date(anio, 6, 20));  
  agregarFestivo(new Date(anio, 7, 7));   
  agregarFestivo(new Date(anio, 11, 8));  
  agregarFestivo(new Date(anio, 11, 25)); 

  // 2. LEY EMILIANI
  agregarFestivo(moverAlLunes(new Date(anio, 0, 6)));   
  agregarFestivo(moverAlLunes(new Date(anio, 2, 19)));  
  agregarFestivo(moverAlLunes(new Date(anio, 5, 29)));  
  agregarFestivo(moverAlLunes(new Date(anio, 7, 15)));  
  agregarFestivo(moverAlLunes(new Date(anio, 9, 12)));  
  agregarFestivo(moverAlLunes(new Date(anio, 10, 1)));  
  agregarFestivo(moverAlLunes(new Date(anio, 10, 11))); 

  // 3. MÓVILES
  var domingoPascua = _calcularDomingoPascua(anio);

  var juevesSanto = new Date(domingoPascua);
  juevesSanto.setDate(domingoPascua.getDate() - 3);
  agregarFestivo(juevesSanto);

  var viernesSanto = new Date(domingoPascua);
  viernesSanto.setDate(domingoPascua.getDate() - 2);
  agregarFestivo(viernesSanto);

  var ascension = new Date(domingoPascua);
  ascension.setDate(domingoPascua.getDate() + 39);
  agregarFestivo(moverAlLunes(ascension));

  var corpus = new Date(domingoPascua);
  corpus.setDate(domingoPascua.getDate() + 60);
  agregarFestivo(moverAlLunes(corpus));

  var sagrado = new Date(domingoPascua);
  sagrado.setDate(domingoPascua.getDate() + 68);
  agregarFestivo(moverAlLunes(sagrado));

  return festivos;
}

// ===================================================================
// 12. FUNCIONES DE RESET MANUAL
// ===================================================================

function asistenciaResetManual() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(ASIS_PROP_LOTE_INICIO);
  props.deleteProperty(ASIS_PROP_EN_CURSO);
  _asisClearTriggers_();

  _asisNotify_('🔄 Generador de asistencia reseteado.', 'Asistencia');
}

function resetearProcesoAsistencia() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(ASIS_PROP_LOTE_INICIO);
  props.deleteProperty(ASIS_PROP_EN_CURSO);
  _asisClearTriggers_();

  Logger.log('✅ Proceso de asistencia reseteado manualmente');

  try {
    SpreadsheetApp.getUi().alert('✅ Sistema desbloqueado. Puedes volver a generar la hoja.');
  } catch (e) {
    Logger.log('Sistema desbloqueado (sin UI disponible)');
  }
}

// ===================================================================
// 13. BACKEND PARA HTML
// ===================================================================

function obtenerDataAsistencia(filtros) {
  filtros = filtros || {};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_DESTINO);

  if (!hoja) {
    return { status: 'error', message: 'La hoja Asistencia_SinValores no existe.' };
  }

  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return { status: 'ok', registros: [] };

  var datos = hoja.getRange(2, 1, lastRow - 1, hoja.getLastColumn()).getDisplayValues();

  var fInicio = filtros.fechaInicio ? new Date(filtros.fechaInicio + 'T00:00:00') : null;
  var fFin = filtros.fechaFin ? new Date(filtros.fechaFin + 'T23:59:59') : null;  
  var fCedula = (filtros.cedula || '').toLowerCase().trim();
  var fNombre = (filtros.nombre || '').toLowerCase().trim();
  var fTipoDia = (filtros.tipoDia || ''); 

  var registros = [];

  for (var i = 0; i < datos.length; i++) {
    var fila = datos[i];

    var fechaFila = null;
    if (fila[4]) {
      var parts = fila[4].split('/');
      if (parts.length === 3) {
        fechaFila = new Date(parts[2] + '-' + parts[1] + '-' + parts[0] + 'T00:00:00');
      }
    }

    if (fInicio && fechaFila && fechaFila < fInicio) continue;
    if (fFin && fechaFila && fechaFila > fFin) continue;

    if (fCedula && String(fila[0]).toLowerCase().indexOf(fCedula) === -1) continue;

    if (fNombre && String(fila[1]).toLowerCase().indexOf(fNombre) === -1) continue;

    if (fTipoDia) {
      var tInicio = String(fila[7] || '');
      var tFin = String(fila[8] || '');

      if (tInicio !== fTipoDia && tFin !== fTipoDia) continue;
    }

    var turnosDetalle = fila[3] || "Sin registro"; 

    registros.push({
      cedula: fila[0],
      nombre: fila[1] || "Sin Nombre",
      centro: fila[2] || "Sin Centro",
      fecha: fila[4], 
      turnosDetalle: turnosDetalle,
      horaInicio: fila[5],
      horaSalida: fila[6],
      tipoDiaInicio: fila[7],
      tipoDiaFin: fila[8],
      horasTotal: parseFloat(fila[9] || 0),
      hDiurNorm: parseFloat(fila[10] || 0),
      hNocNorm: parseFloat(fila[11] || 0),
      hDiurFest: parseFloat(fila[12] || 0),
      hNocFest: parseFloat(fila[13] || 0)
    });
  }

  return { status: 'ok', registros: registros };
}

function exportarAsistenciaCSV(filtros) {
  var dataObj = obtenerDataAsistencia(filtros);
  if (dataObj.status !== 'ok') return dataObj;

  var registros = dataObj.registros;

  var csv = "Cédula;Nombre;Centro;Fecha;Hora Inicio;Hora Salida;Rango Completo;Tipo Día Inicio;Tipo Día Fin;Total Horas;H.Ord.Diurna;H.Ord.Nocturna;H.Fest.Diurna;H.Fest.Nocturna\n";

  var fmtNum = function(n) {
    return String(n).includes('.') ? String(n).replace('.', ',') : String(n);
  };

  var escape = function(v) {
    if (v == null) return '""';
    var s = String(v);
    return '"' + s.replace(/"/g, '""') + '"';
  };

  for (var i = 0; i < registros.length; i++) {
    var r = registros[i];
    csv += [
      escape(r.cedula),
      escape(r.nombre),
      escape(r.centro),
      escape(r.fecha),
      escape(r.horaInicio),
      escape(r.horaSalida),
      escape(r.turnosDetalle),
      escape(r.tipoDiaInicio),
      escape(r.tipoDiaFin),
      fmtNum(r.horasTotal),
      fmtNum(r.hDiurNorm),
      fmtNum(r.hNocNorm),
      fmtNum(r.hDiurFest),
      fmtNum(r.hNocFest)
    ].join(';') + "\n";
  }

  var fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");

  return {
    status: 'ok',
    csvContent: csv,
    filename: 'Reporte_Asistencia_' + fechaActual + '.csv'
  };
}

// ===================================================================
// 14. FUNCIONES DE PRUEBA
// ===================================================================

function testFestivos2026() {
  var festivos2026 = _generarFestivos(2026);

  Logger.log('=== FESTIVOS COLOMBIA 2026 ===');
  Logger.log('Total: ' + festivos2026.size);

  var lista = Array.from(festivos2026).sort();
  for (var i = 0; i < lista.length; i++) {
    Logger.log('  ' + lista[i]);
  }

  var cache = {};
  cache[2026] = festivos2026;

  Logger.log('=== PRUEBAS DE FECHAS ===');
  Logger.log('01/01/2026: ' + _esFestivo(new Date(2026, 0, 1), cache));   
  Logger.log('02/01/2026: ' + _esFestivo(new Date(2026, 0, 2), cache));   
  Logger.log('04/01/2026: ' + _esFestivo(new Date(2026, 0, 4), cache));   
  Logger.log('05/01/2026: ' + _esFestivo(new Date(2026, 0, 5), cache));   
  Logger.log('12/01/2026: ' + _esFestivo(new Date(2026, 0, 12), cache));  
  Logger.log('25/12/2026: ' + _esFestivo(new Date(2026, 11, 25), cache)); 
}

function testCalculoHoras() {
  var cache = {};

  var turno = {
    inicio: new Date(2026, 0, 2, 7, 0, 0),  
    fin: new Date(2026, 0, 2, 17, 0, 0)     
  };

  var resultado = _calcularHorasPorTipo([turno], cache);

  Logger.log('=== TEST CÁLCULO HORAS ===');
  Logger.log('Turno: 02/01/2026 07:00 - 17:00');
  Logger.log('Tipo día: ' + _esFestivo(turno.inicio, cache));
  Logger.log('Total: ' + resultado.total);
  Logger.log('Diurnas Normales: ' + resultado.normalesDia);
  Logger.log('Nocturnas Normales: ' + resultado.normalesNoc);
  Logger.log('Diurnas Festivo: ' + resultado.festivosDia);
  Logger.log('Nocturnas Festivo: ' + resultado.festivosNoc);

  var turnoNoc = {
    inicio: new Date(2026, 0, 2, 22, 0, 0),  
    fin: new Date(2026, 0, 3, 6, 0, 0)       
  };

  var resultadoNoc = _calcularHorasPorTipo([turnoNoc], cache);

  Logger.log('');
  Logger.log('Turno: 02/01/2026 22:00 - 03/01/2026 06:00');
  Logger.log('Total: ' + resultadoNoc.total);
  Logger.log('Diurnas Normales: ' + resultadoNoc.normalesDia);
  Logger.log('Nocturnas Normales: ' + resultadoNoc.normalesNoc);
}
