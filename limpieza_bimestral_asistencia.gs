// ===================================================================
// 🧹 limpieza_bimestral_asistencia.gs – Limpieza Automática (NASE 2026)
// -------------------------------------------------------------------
/**
 * @summary Módulo de Limpieza Automática (Ciclo Bimestral).
 * @description Gestiona la limpieza masiva de la hoja "Asistencia_SinValores"
 *              para evitar que el archivo de Google Sheets crezca indefinidamente.
 * 
 * @logic
 *   - ⚡ **Trigger:** Se programa para ejecutarse el Día 1 de cada mes a las 15:00.
 *   - 📅 **Ciclo Bimestral:** El script solo se ejecuta en meses impares del calendario
 *     (Enero, Marzo, Mayo, Julio, Septiembre, Noviembre).
 *     - ¿Por qué? La lógica es `if (mes % 2 !== 0) return;`.
 *     - Esto filtra para que la limpieza ocurra al INICIO de cada periodo bimestral.
 *     - Enero: Limpia (Inicio). Febrero: Mantiene. Marzo: Limpia (Inicio). Abril: Mantiene...
 *   - 🗑️ **Acción de Limpieza:** Borra todo el contenido de las filas de datos (dejando el encabezado).
 *   - 🛡️ **Seguridad (Sin Respaldo Local):**
 *       Este archivo NO crea respaldos dentro del Spreadsheet actual.
 *       Se asume que `archivo_mensual_asistencia.gs` (que corre el día 1 a las 12:00 PM)
 *       ya ha creado una copia de seguridad en Google Drive.
 *       Esto asegura que los datos del mes anterior no se pierden antes de limpiar.
 *   - 🔁 **Ciclo:**
 *       Enero (Limpia) -> Febrero (No limpia) -> Marzo (Limpia) -> ...
 *
 * @dependencies
 *   - `install_triggers.gs` (Función `ensureTimeTrigger`).
 *   - `archivo_mensual_asistencia.gs` (Debe ejecutarse 3 horas antes para respaldar).
 *
 * @author NASE Team
 * @version 1.2 (Actualizado para Horas Extras)
 */

// ===================================================================
// 1. INSTALACIÓN DEL DISPARADOR (TRIGGER)
// ===================================================================

/**
 * @summary Instala el disparador bimestral de limpieza.
 * @description Función de instalación (manual o al desplegar).
 *              Utiliza `ensureTimeTrigger` para verificar si ya existe.
 * 
 * @schedule
 *   - Día del mes: 1 (Primero de cada mes).
 *   - Hora: 15 (3:00 PM).
 *   - Frecuencia: Mensual, pero la función interna tiene un filtro de meses impares.
 */
function instalarTriggersLimpiezaBimestral() {
  // Wrapper de seguridad para crear trigger si no existe
  ensureTimeTrigger("limpiarAsistenciaBimestral", function () {
    ScriptApp.newTrigger("limpiarAsistenciaBimestral")
      .timeBased()
      .onMonthDay(1) // Se ejecuta el día 1
      .atHour(15)    // A las 15:00 (3 PM)
      .create();
  });
  Logger.log("✅ Trigger bimestral limpieza Asistencia_SinValores instalado (NASE 2026).");
}

// ===================================================================
// 2. LÓGICA DE LIMPIEZA (Ciclo Bimestral)
// ===================================================================

/**
 * @summary Limpia la hoja de asistencia si corresponde al mes.
 * @description Función principal que se ejecuta automáticamente por el Trigger.
 *              Realiza lo siguiente:
 *   1. Obtiene la fecha actual del sistema.
 *   2. Verifica si el mes es impar (Enero, Marzo, Mayo...).
 *   3. Si es impar, limpia la hoja "Asistencia_SinValores".
 *   4. Muestra un Toast en la hoja y un mensaje en Log.
 * 
 * @safety
 *   - Al ser bimestral (Cada 2 meses), el archivo permanece limpio por dos meses.
 *   - Se recomienda que el archivo mensual (`archivo_mensual_asistencia.gs`) corra
 *     siempre el día 1 a las 12:00 PM, 3 horas ANTES de esta limpieza, para respaldar.
 * 
 * @note NO crea respaldo interno. El respaldo es el archivo mensual en Drive.
 */
function limpiarAsistenciaBimestral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Asistencia_SinValores");
  
  // Validación básica: Si no existe hoja o está vacía, no hacer nada
  if (!hoja || hoja.getLastRow() <= 1) return;

  const hoy = new Date();
  const mes = hoy.getMonth(); // 0=Enero, 1=Febrero, ... 11=Diciembre

  // -----------------------------------------------------------------
  // 1. FILTRO DE FRECUENCIA (Solo meses impares del calendario)
  // -----------------------------------------------------------------
  // La lógica `if (mes % 2 !== 0)` significa:
  // - Si el resto de la división por 2 NO es cero (es impar), salte (return).
  // - Se ejecuta solo si el resultado es 0 (Par).
  // En 0-indexing (0=En, 1=Feb, 2=Mar...), los pares son Ene(0), Mar(2), Mayo(4).
  // Que corresponden a los meses 1, 3, 5 del calendario (Impares).
  // Por tanto, la limpieza se ejecuta en Enero, Marzo, Mayo...
  if (mes % 2 !== 0) return; 

  // -----------------------------------------------------------------
  // 2. ACCIÓN DE LIMPIEZA (Borrar Filas)
  // -----------------------------------------------------------------
  
  // ✅ Solo limpiar, sin crear respaldo interno
  // El respaldo se confía al archivo mensual generado anteriormente
  const lastRow = hoja.getLastRow();
  
  if (lastRow > 1) {
    // Borra desde la fila 2 hasta la última fila, todas las columnas
    // Mantiene los encabezados (fila 1)
    hoja.getRange(2, 1, lastRow - 1, hoja.getLastColumn()).clearContent();
  }

  // -----------------------------------------------------------------
  // 3. FEEDBACK VISUAL (Toast y Log)
  // -----------------------------------------------------------------
  
  // Mostrar mensaje en la hoja para el usuario
  SpreadsheetApp.getActive().toast(
    `✅ Limpieza bimestral Asistencia_SinValores completada.`,
    "Sistema Limpiado",
    8 // Segundos visibles
  );

  Logger.log(`✅ Limpieza bimestral Asistencia_SinValores completada. Sin respaldo interno.`);
}
