// ===================================================================
// 📁 install_triggers.gs – Disparadores NASE 2026 (Horas Extras)
// -------------------------------------------------------------------
/**
 * @summary Módulo de Automatización (Cron Jobs).
 * @description Gestiona la instalación, eliminación y verificación de disparadores
 *              (Triggers) para el sistema de Control de Horas Extras.
 * 
 * @tasks Automatizadas Instaladas:
 * - 📊 Generación de resumen mensual (Día 1, 9:00 AM).
 * - 🧹 Limpieza bimestral de Asistencia (Día 1, 3:00 PM).
 * - 🗑️ Limpieza bimestral de Respuestas con Archivo (Día 1, 4:00 PM).
 * - 🔄 Actualización de caché de empleados (Cada 2 horas).
 * - 💓 Mantenimiento de sistema activo (Cada 10 min).
 *
 * @author NASE Team
 * @version 1.2
 */

// ===================================================================
// 1. FUNCIÓN PRINCIPAL DE INSTALACIÓN
// ===================================================================

/**
 * @summary Instala y configura todos los disparadores automáticos.
 */
function installTriggers() {
  var current = ScriptApp.getProjectTriggers();

  // ----------------------------------------------------------------
  // 1.1 SEGURIDAD: Si hay demasiados triggers, borrar todo
  // ----------------------------------------------------------------
  if (current.length >= 18) {
    wipeAllTriggers();
  } else {
    // ----------------------------------------------------------------
    // 1.2 MANTENIMIENTO: Eliminar triggers duplicados o específicos
    // ----------------------------------------------------------------
    var toRemove = {
      // Funciones que queremos reinstalar limpiamente
      "generarTablaAsistenciaSinValores": true,
      "actualizarCacheEmpleados": true,
      "mantenerSistemaActivo": true
    };
    
    for (var i = 0; i < current.length; i++) {
      var fn = current[i].getHandlerFunction();
      if (toRemove[fn]) ScriptApp.deleteTrigger(current[i]);
    }
  }

  // ----------------------------------------------------------------
  // 1.3 INSTALACIÓN DE TRIGGERS
  // ----------------------------------------------------------------

  // 1. 📊 Generación de Resumen Mensual
  //    Se ejecuta el día 1 de cada mes a las 9:00 AM
  ensureTimeTrigger("generarTablaAsistenciaSinValores", function() {
    ScriptApp.newTrigger("generarTablaAsistenciaSinValores")
      .timeBased()
      .onMonthDay(1)
      .atHour(9)
      .create();
  });

  // 2. 🔄 Actualización de Caché de Empleados
  //    Se ejecuta cada 2 horas (para mantener actualizada la base de RRHH)
  ensureTimeTrigger("actualizarCacheEmpleados", function() {
    ScriptApp.newTrigger("actualizarCacheEmpleados")
      .timeBased()
      .everyHours(2)
      .create();
  });

  // 3. 💓 Mantener Sistema Activo (Keep-Alive)
  //    Se ejecuta cada 10 minutos para evitar que el script se "apague"
  ensureTimeTrigger("mantenerSistemaActivo", function() {
    ScriptApp.newTrigger("mantenerSistemaActivo")
      .timeBased()
      .everyMinutes(10)
      .create();
  });

  // 4. ✅ Archivo Mensual de Asistencia (Guarda solo en Drive)
  //    Se ejecuta el día 1 de cada mes a las 12:00 PM
  ensureTimeTrigger("generarArchivoMensualAsistencia", function() {
    ScriptApp.newTrigger("generarArchivoMensualAsistencia")
      .timeBased()
      .onMonthDay(1)
      .atHour(12)
      .create();
  });

  // 5. ✅ Limpieza Bimestral de Asistencia (Sin respaldo en Spreadsheet)
  //    Se ejecuta el día 1 impar (en meses impares) a las 3:00 PM
  ensureTimeTrigger("limpiarAsistenciaBimestral", function() {
    ScriptApp.newTrigger("limpiarAsistenciaBimestral")
      .timeBased()
      .onMonthDay(1)
      .atHour(15)
      .create();
  });

  // 6. ✅ Limpieza Bimestral de Respuestas (Con Archivo en Drive)
  //    Se ejecuta el día 1 impar (en meses impares) a las 4:00 PM
  ensureTimeTrigger("limpiarRespuestasBimestral", function() {
    ScriptApp.newTrigger("limpiarRespuestasBimestral")
      .timeBased()
      .onMonthDay(1)
      .atHour(16)
      .create();
  });

  Logger.log("✅ Triggers de Horas Extras instalados. Total actuales: " + ScriptApp.getProjectTriggers().length);
}

// ===================================================================
// 2. UTILIDADES DE CONTROL (Eliminación y Listado)
// ===================================================================

/**
 * @summary Elimina absolutamente todos los disparadores activos.
 */
function wipeAllTriggers() {
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    ScriptApp.deleteTrigger(all[i]);
  }
  return true;
}

/**
 * @summary Lista en el log todos los disparadores activos.
 */
function listTriggers() {
  var all = ScriptApp.getProjectTriggers();
  Logger.log("Triggers actuales: " + all.length);
  for (var i = 0; i < all.length; i++) {
    Logger.log("#" + (i + 1) + 
               " handler=" + all[i].getHandlerFunction() +
               ", source=" + all[i].getTriggerSource() +
               ", event=" + all[i].getEventType());
  }
}

// ===================================================================
// 3. FUNCIÓN AUXILIAR (Gestión Inteligente de Triggers)
// ===================================================================

/**
 * @summary Verifica y asegura que exista un único trigger por función.
 * @description Previene la creación de múltiples triggers duplicados.
 */
function ensureTimeTrigger(handlerName, createFn) {
  // Buscar si ya existe un trigger con ese nombre de función
  var found = ScriptApp.getProjectTriggers().filter(function(t){
    return t.getHandlerFunction() === handlerName; 
  });

  // Caso A: No existe -> Crearlo
  if (found.length === 0) {
    createFn();
    Logger.log("Trigger creado: " + handlerName);
  } 
  // Caso B: Existe más de uno (Duplicados) -> Eliminar extras y dejar 1
  else if (found.length > 1) {
    for (var i = 1; i < found.length; i++) {
      ScriptApp.deleteTrigger(found[i]);
    }
    Logger.log("Duplicados eliminados para: " + handlerName);
  } 
  // Caso C: Ya existe 1 -> No hacer nada
  else {
    Logger.log("Trigger ya existe: " + handlerName);
  }
}
