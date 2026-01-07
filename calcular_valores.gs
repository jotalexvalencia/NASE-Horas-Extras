// ============================================================
// 💰 calcular_valores.gs – Formateo Horas Extras (NASE 2026)
// ------------------------------------------------------------
/**
 * @summary Módulo de Formateo y Limpieza para Asistencia.
 * @description 
 * Especializado en la presentación visual de datos para el proyecto de Horas Extras:
 * 
 * ✅ FUNCIONES ACTUALES:
 * - Formateo de columnas de tiempo: "Total Horas Extras", "Horas Nocturnas", etc.
 * - Eliminación de rastros de columnas monetarias de versiones antiguas.
 * 
 * 🎨 Lógica de Formateo:
 * - Busca columnas clave (Entrada, Salida, Horas Extras).
 * - Aplica formato #,##0.00.
 * - Limpia datos sucios.
 *
 * @author NASE Team
 * @version 2.1 (Adaptado para Horas Extras)
 */

// ============================================================
// 1. FUNCIÓN PRINCIPAL
// ============================================================

/**
 * @summary Procesa y limpia la hoja "Asistencia_SinValores".
 * @description 
 * 1. Aplica formato numérico a las columnas de Tiempo/Horas Extras.
 * 2. Elimina columnas monetarias obsoletas.
 */
function agregarValoresAsistencia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Asistencia_SinValores");
  
  if (!hoja) {
    Logger.log("❌ No existe la hoja 'Asistencia_SinValores'. Ejecuta primero la generación de turnos.");
    SpreadsheetApp.getUi().alert("⚠️ No existe la hoja 'Asistencia_SinValores'.");
    return;
  }

  const ultimaFila = hoja.getLastRow();
  
  if (ultimaFila < 2) {
    Logger.log("⚠️ No hay datos para procesar en Asistencia_SinValores.");
    return;
  }

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];

  // -----------------------------------------------------------
  // 1. IDENTIFICACIÓN DE COLUMNAS DE TIEMPO
  // -----------------------------------------------------------
  const findHeaderIndex = nombre =>
    encabezados.findIndex(h => String(h).trim().toLowerCase() === nombre.trim().toLowerCase());

  // Columnas estándar de Asistencia
  const colHorasTotales = findHeaderIndex("Total Horas");
  const colHorasDiurnas = findHeaderIndex("Horas Diurnas");
  const colHorasNocturnas = findHeaderIndex("Horas Nocturnas Normales");
  const colDomDiurnas = findHeaderIndex("Horas Diurnas Domingo/Festivo");
  const colDomNocturnas = findHeaderIndex("Horas Nocturnas Domingo/Festivo");

  // --- NUEVAS COLUMNAS DE HORAS EXTRAS ---
  const colTotalHE = findHeaderIndex("Total Horas Extras");
  const colTotalNoct = findHeaderIndex("Total Horas Nocturnas");

  // Lista de índices de columnas de TIEMPO encontradas (incluyendo HE)
  const indicesTiempo = [
    colHorasTotales, colHorasDiurnas, colHorasNocturnas, 
    colDomDiurnas, colDomNocturnas, 
    colTotalHE, colTotalNoct 
  ].filter(idx => idx !== -1); // Filtrar los que no se encontraron

  // -----------------------------------------------------------
  // 2. ACCIÓN: APLICAR FORMATO NUMÉRICO
  // -----------------------------------------------------------
  // Aplica formato #,##0.00 (miles con punto, decimales con coma)
  // a todas las columnas de tiempo detectadas.
  indicesTiempo.forEach(idx => {
    hoja.getRange(2, idx + 1, ultimaFila, 1).setNumberFormat("#,##0.00");
  });

  // -----------------------------------------------------------
  // 3. ACCIÓN: LIMPIEZA DE COLUMNAS MONETARIAS ANTIGUAS
  // -----------------------------------------------------------
  const colsMonetariasABorrar = [
    "Valor Diurno Domingo/Festivo",
    "Valor Nocturno Día Ordinario",
    "Valor Nocturno Domingo/Festivo",
    "Total Valores"
  ];

  let borradas = 0;
  // Recorrer de derecha a izquierda
  for (let i = encabezados.length - 1; i >= 0; i--) {
    const headerName = String(encabezados[i]).trim();
    if (colsMonetariasABorrar.includes(headerName)) {
      hoja.deleteColumn(i + 1);
      borradas++;
      Logger.log(`🗑️ Columna monetaria eliminada: ${headerName}`);
    }
  }

  // -----------------------------------------------------------
  // 4. RESULTADOS Y LOGGING
  // -----------------------------------------------------------
  Logger.log(`✅ Formateo Horas Extras finalizado. ${indicesTiempo.length} columnas formateadas. ${borradas} columnas monetarias eliminadas.`);
  
  try {
    SpreadsheetApp.getActive().toast("Formateo Horas Extras completado.", "Sistema", 5);
  } catch(e) {
    // Ignorar si no hay UI
  }
}
