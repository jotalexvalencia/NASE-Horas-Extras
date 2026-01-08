// ===================================================================
// 📂 archivo_mensual_asistencia.gs – Archivo Histórico (NASE 2026)
// -------------------------------------------------------------------
/**
 * @summary Módulo de Archivo Automático de Nómina y Horas Extras.
 * @description Este archivo automatiza la generación de copias de seguridad
 *              de la hoja de nómina "Asistencia_SinValores" al final de cada mes.
 *
 * @workflow
 * - 🔁 **Trigger Automático:** Se ejecuta el día 1 de cada mes a las 12:00 PM.
 * - 📅 **Target:** Archiva los datos del *mes anterior*.
 *   Ejemplo: Si se ejecuta el 1 de Febrero, archiva los datos de Enero.
 * - 📁 **Ubicación:** Crea un nuevo archivo de Google Sheets y lo guarda en una
 *   carpeta específica de Drive: "Archivos Asistencia Mensual NASE".
 * - ✅ **Inclusivo:** El archivo archivado contiene TODAS las columnas de `Asistencia_SinValores`,
 *   incluyendo las nuevas columnas de "Total Horas Extras" y aprobaciones.
 *
 * @constraints
 *   - ⛔ NO limpia la hoja original (`Asistencia_SinValores`).
 *   - ⛔ NO crea respaldos internos en el Spreadsheet actual.
 *   - ✅ Crea archivos nuevos por mes en Google Drive.
 *
 * @author NASE Team
 * @version 1.2 (Actualizado para Horas Extras)
 */

// ===================================================================
// 1. INSTALACIÓN DE DISPARADOR (TRIGGER)
// ===================================================================

/**
 * @summary Instala el disparador mensual de archivo.
 * @description Función de instalación (manual o inicial).
 *              Utiliza `ensureTimeTrigger` (utility de `install_triggers`)
 *              para evitar duplicados y configurar la ejecución.
 * 
 * @schedule Día 1 de cada mes a las 12:00 PM.
 */
function instalarTriggersAsistenciaMensual() {
  // Wrapper de seguridad para crear trigger si no existe
  ensureTimeTrigger("generarArchivoMensualAsistencia", function () {
    ScriptApp.newTrigger("generarArchivoMensualAsistencia")
      .timeBased()
      .onMonthDay(1) // Se ejecuta el día 1 del mes
      .atHour(12)    // A las 12:00 PM
      .create();
  });
  Logger.log("✅ Trigger mensual Asistencia_SinValores instalado (NASE 2026).");
}

// ===================================================================
// 2. LÓGICA DE ARCHIVO
// ===================================================================

/**
 * @summary Genera el archivo histórico del mes anterior.
 * @description Función principal que se ejecuta automáticamente.
 *              1. Lee la hoja "Asistencia_SinValores".
 *              2. Calcula la fecha del mes anterior.
 *              3. Crea un nuevo Spreadsheet en Drive.
 *              4. Copia los datos al nuevo archivo (incluyendo columnas HE).
 *              5. Mueve el archivo a la carpeta histórica.
 * 
 * @returns {void} Escribe logs en consola.
 */
function generarArchivoMensualAsistencia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Asistencia_SinValores');
  
  // Validar que exista la hoja de origen (con los cálculos de horas)
  if (!hoja) {
    Logger.log("❌ No se encontró la hoja 'Asistencia_SinValores' para archivar.");
    return;
  }

  // -----------------------------------------------------------
  // 1. CALCULAR FECHA DEL MES ANTERIOR (Contexto)
  // -----------------------------------------------------------
  const ahora = new Date();
  // (Año actual, Mes actual - 1, Día 1)
  const mesAnterior = new Date(ahora.getFullYear(), ahora.getMonth() - 1, 1);
  
  // Formatear nombre del mes (Ej: "enero", "febrero")
  const nombreMes = mesAnterior.toLocaleString('es-ES', { month: 'long', year: 'numeric' });
  
  // Crear nombre del archivo (Ej: "Asistencia_enero_2026")
  const nombreArchivo = `Asistencia_${nombreMes.replace(' ', '_')}`;

  // -----------------------------------------------------------
  // 2. OBTENER O CREAR CARPETA DE DRIVE
  // -----------------------------------------------------------
  const folder = obtenerOCrearCarpeta('Archivos Asistencia Mensual NASE');

  // -----------------------------------------------------------
  // 3. CREAR NUEVO ARCHIVO SPREADSHEET
  // -----------------------------------------------------------
  const archivo = SpreadsheetApp.create(nombreArchivo);
  
  // Mover el archivo recién creado a la carpeta específica
  DriveApp.getFileById(archivo.getId()).moveTo(folder);

  // -----------------------------------------------------------
  // 4. COPIAR DATOS (Incluye columnas de Horas Extras)
  // -----------------------------------------------------------
  // `copyTo` copia la hoja entera, incluidas las fórmulas y valores calculados
  const hojaCopia = hoja.copyTo(archivo);
  
  // Renombrar la hoja dentro del archivo nuevo para mantener consistencia
  hojaCopia.setName('Asistencia_' + nombreMes);

  // -----------------------------------------------------------
  // 5. LIMPIEZA DE ARCHIVO NUEVO
  // -----------------------------------------------------------
  // Al crear un Spreadsheet, se crea por defecto una hoja llamada "Hoja 1".
  // Eliminamos esa hoja predeterminada para dejar solo la copia que traemos.
  const hojas = archivo.getSheets();
  if (hojas.length > 1) {
    hojas.forEach(h => {
      if (h.getName() !== hojaCopia.getName()) {
        archivo.deleteSheet(h);
      }
    });
  }

  Logger.log(`✅ Archivo mensual generado con cálculos de Horas Extras: ${nombreArchivo}`);
}

// ===================================================================
// 3. UTILIDAD DE CARPETAS (DRIVE API)
// ===================================================================

/**
 * @summary Busca una carpeta por nombre en Drive. Si no existe, la crea.
 * @description Utiliza `getFoldersByName` para verificar existencia.
 *              Usa `createFolder` para generar la carpeta si falta.
 * 
 * @param {String} nombre - Nombre exacto de la carpeta en Drive.
 * @returns {Folder} Objeto Carpeta de Google Drive.
 * @private
 */
function obtenerOCrearCarpeta(nombre) {
  // Buscar carpetas con ese nombre exacto
  const folders = DriveApp.getFoldersByName(nombre);
  
  // Si existe alguna, retornar la primera
  if (folders.hasNext()) {
    return folders.next();
  }
  
  // Si no existe, crearla
  return DriveApp.createFolder(nombre);
}
