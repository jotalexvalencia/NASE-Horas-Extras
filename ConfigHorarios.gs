/**
 * ============================================================
 * ⚙️ ConfigHorarios.gs – Gestor de Configuración (NASE 2026)
 * ======================================================================
 * @summary Almacén de configuración de recargos nocturnos.
 * @description Este archivo actúa como el "almacén de configuración" del sistema.
 *              Gestiona los rangos horarios para definir cuándo aplica el
 *              Recargo Nocturno (necesario para liquidar Horas Extras).
 *
 * @features
 *   - 🧹 **Código Limpio:** Se eliminaron configuraciones de nómina (%), solo tiempos.
 *   - 💾 **Persistencia:** Usa `ScriptProperties` (memoria del script) para guardar
 *     las preferencias. Es más rápido y seguro que escribir en una hoja.
 *   - 🕰️ **Defecto Legal:** Por defecto define Nocturno como 19:00 a 06:00
 *     (conforme a Ley 2025).
 *
 * @dependencies
 *   - `hoja_turnos.gs`: Utiliza `obtenerConfiguracionHorarios()` para decidir
 *     si una hora trabajada cuenta como "Nocturna" (recargo).
 *   - `config_horarios.html` (si existe): Utiliza `actualizarConfiguracionHorarios()`
 *     para guardar cambios desde la interfaz.
 *
 * @author NASE Team
 * @version 2.1 (Simplificado - Solo Tiempos Horas Extras)
 */

// ======================================================================
// 1. CONSTANTES (Nombres de Propiedades)
// ======================================================================

const CONFIG_PROPS = {
  // Claves en `ScriptProperties` para guardar las horas de inicio/fin nocturno
  HORA_NOCTURNA_INICIO: 'HORA_NOCTURNA_INICIO', // Default: 19 (7 PM)
  HORA_NOCTURNA_FIN: 'HORA_NOCTURNA_FIN',         // Default: 6  (6 AM)
  
  // Nota: Se eliminaron propiedades de porcentajes monetarios.
  // El cálculo de nómina se maneja externamente o en reportes.
};

// ======================================================================
// 2. LECTURA DE CONFIGURACIÓN
// ======================================================================

/**
 * @summary Obtiene la configuración actual de recargo nocturno.
 * @description Lee las propiedades de `ScriptProperties`.
 *              Si no existen (primera ejecución), devuelve los valores por defecto
 *              establecidos por la Ley (19:00 - 06:00).
 * 
 * @returns {Object} Objeto con:
 *   - `horaInicio` (Number): Hora de inicio del recargo nocturno (Ej: 19).
 *   - `horaFin` (Number): Hora de fin del recargo nocturno (Ej: 6).
 */
function obtenerConfiguracionHorarios() {
  const props = PropertiesService.getScriptProperties();
  
  // Si no existe valor guardado, usa el default (19 y 6)
  return {
    horaInicio: parseInt(props.getProperty(CONFIG_PROPS.HORA_NOCTURNA_INICIO) || '19', 10),
    horaFin: parseInt(props.getProperty(CONFIG_PROPS.HORA_NOCTURNA_FIN) || '6', 10)
  };
}

// ======================================================================
// 3. ESCRITURA DE CONFIGURACIÓN
// ======================================================================

/**
 * @summary Actualiza (Guarda) la configuración de recargos.
 * @description Se ejecuta desde el formulario de configuración HTML.
 *              Guarda las horas de inicio y fin en `ScriptProperties` para que
 *              persistan entre ejecuciones.
 * 
 * @param {Object} config - Objeto con:
 *   - `horaInicio` (Number): Nueva hora de inicio (0-23).
 *   - `horaFin` (Number): Nueva hora de fin (0-23).
 * 
 * @returns {Object} { status: 'ok', message: String }
 */
function actualizarConfiguracionHorarios(config) {
  const props = PropertiesService.getScriptProperties();
  
  // Guardar hora de inicio
  if (config.horaInicio !== undefined) {
    props.setProperty(CONFIG_PROPS.HORA_NOCTURNA_INICIO, String(config.horaInicio));
  }
  
  // Guardar hora de fin
  if (config.horaFin !== undefined) {
    props.setProperty(CONFIG_PROPS.HORA_NOCTURNA_FIN, String(config.horaFin));
  }
  
  return { status: 'ok', message: 'Configuración de recargo nocturno actualizada correctamente.' };
}

// ======================================================================
// 4. RESET DE CONFIGURACIÓN
// ======================================================================

/**
 * @summary Restablece los valores por defecto.
 * @description Función de seguridad para volver al estado original del sistema.
 *              Borra las propiedades personalizadas y fuerza el uso de 19:00 - 06:00.
 * 
 * @returns {Object} { status: 'ok', message: String }
 */
function restablecerConfiguracionPorDefecto() {
  const props = PropertiesService.getScriptProperties();
  
  // Sobrescribir con valores por defecto (19 y 6)
  props.setProperty(CONFIG_PROPS.HORA_NOCTURNA_INICIO, '19');
  props.setProperty(CONFIG_PROPS.HORA_NOCTURNA_FIN, '6');
  
  return { status: 'ok', message: 'Recargo nocturno restablecido a ley (19:00 - 06:00).' };
}
