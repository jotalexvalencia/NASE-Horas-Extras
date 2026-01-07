// ============================================================
// ⏱️ Traer_horas_laborales.gs – Carga de Base Horaria (NASE 2026)
// ------------------------------------------------------------
/**
 * @summary Módulo de Sincronización de Datos Laborales para Horas Extras.
 * @description Esta función conecta el sistema NASE con el libro externo 
 *              de "Base Operativa" (RRHH) para traer las horas pactadas semanales.
 * 
 * IMPORTANTE PARA HORAS EXTRAS:
 * Para calcular correctamente las horas extra diarias o semanales, el sistema
 * necesita saber cuántas horas debe trabajar el colaborador. Este script 
 * trae ese dato ("Horas Laborales por Semana") y lo inyecta en la hoja
 * de asistencia/reportes.
 * 
 * @features
 * - 🔗 **Conexión Externa:** Abre el libro de RRHH.
 * - 📂 **Gestión de Columnas:** Crea "Horas Laborales por Semana" si falta.
 * - 🧠 **Mapa de Memoria:** Elige el contrato más reciente si hay duplicados.
 * - ✅ **Actualización Masiva:** Cruza y actualiza la hoja de reportes.
 *
 * @author NASE Team
 * @version 1.3 (Adaptado para Horas Extras)
 */

// ======================================================================
// FUNCIÓN PRINCIPAL
// ======================================================================

/**
 * @summary Sincroniza las "Horas Laborales por Semana" desde RRHH.
 * @description 
 * 1. Abre el libro "Base Operativa".
 * 2. Busca el registro más reciente por cédula.
 * 3. Crea la columna en Asistencia si falta.
 * 4. Cruza y actualiza los datos.
 * 
 * @returns {void} Escribe en `Logger` y muestra alerta.
 */
function insertarHorasLaboralesPorCedula() {
  // -----------------------------------------------------------
  // 1. CONFIGURACIÓN Y APERTURA DE LIBROS
  // -----------------------------------------------------------
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAsistencia = ss.getSheetByName("Asistencia_SinValores");
  
  // Validar hoja destino (Hoja de Reportes)
  if (!hojaAsistencia) {
    throw new Error("❌ No se encontró la hoja 'Asistencia_SinValores' (Reportes).");
  }

  // ID del Libro de RRHH (Base Operativa) - Fuente de datos
  const ID_BASE_OPERATIVA = "1bU-lyiQzczid62n8timgUguW6UxC3qZN8Vehnn26zdY";
  const libroBase = SpreadsheetApp.openById(ID_BASE_OPERATIVA);
  const hojaBase = libroBase.getSheetByName("BASE OPERATIVA");

  if (!hojaBase) {
    throw new Error("❌ No se encontró la hoja 'BASE OPERATIVA' en la Base Operativa.");
  }

  // -----------------------------------------------------------
  // 2. PREPARACIÓN DE HOJA DESTINO (Asistencia)
  // -----------------------------------------------------------
  
  const headersAsistencia = hojaAsistencia.getRange(1, 1, 1, hojaAsistencia.getLastColumn()).getValues()[0];
  const colCedulaAsistencia = headersAsistencia.findIndex(h => String(h).trim().toLowerCase() === "cédula") + 1;
  
  if (colCedulaAsistencia === 0) {
    throw new Error("⚠️ No se encontró la columna 'Cédula' en Asistencia_SinValores.");
  }

  const nombreColumnaNueva = "Horas Laborales por Semana";
  let colNueva = headersAsistencia.findIndex(h => String(h).trim() === nombreColumnaNueva) + 1;

  // -----------------------------------------------------------
  // 3. GESTIÓN DE COLUMNAS (Crear si falta)
  // -----------------------------------------------------------
  
  // Si la columna NO existe, insertarla justo después de la columna 'Cédula'
  if (colNueva === 0) {
    hojaAsistencia.insertColumnAfter(colCedulaAsistencia);
    hojaAsistencia.getRange(1, colCedulaAsistencia + 1).setValue(nombreColumnaNueva);
    colNueva = colCedulaAsistencia + 1;
  }

  // -----------------------------------------------------------
  // 4. LECTURA Y PROCESAMIENTO DE DATOS ORIGEN (Base Operativa)
  // -----------------------------------------------------------
  const dataBase = hojaBase.getDataRange().getValues();
  const headersBase = dataBase[0];
  const headersBaseUpper = headersBase.map(h => (h || "").toString().trim().toUpperCase());

  const idxCedulaBase = headersBaseUpper.indexOf("DOCUMENTO DE IDENTIDAD");
  const idxHorasBase = headersBaseUpper.indexOf("HORAS LABORALES POR SEMANA");
  const idxFechaBase = headersBaseUpper.indexOf("FECHA DE INGRESO");

  if ([idxCedulaBase, idxHorasBase, idxFechaBase].includes(-1)) {
    throw new Error("⚠️ Faltan columnas requeridas en Base Operativa.");
  }

  // -----------------------------------------------------------
  // 5. CREAR MAPA DE MEMORIA { Cédula -> { Horas, Fecha } }
  // -----------------------------------------------------------
  const mapaHoras = {};

  for (let i = 1; i < dataBase.length; i++) {
    const fila = dataBase[i];
    const cedula = String(fila[idxCedulaBase]).replace(/\D/g, "").trim();
    
    if (!cedula) continue;

    const horas = fila[idxHorasBase];
    const fechaIngreso = fila[idxFechaBase];
    
    if (!fechaIngreso) continue;

    const fecha = fechaIngreso instanceof Date ? fechaIngreso : new Date(fechaIngreso);
    
    if (!fecha || isNaN(fecha)) continue;

    // Seleccionar el contrato más reciente
    if (!mapaHoras[cedula] || fecha > mapaHoras[cedula].fecha) {
      mapaHoras[cedula] = { 
        horas: horas, 
        fecha: fecha 
      };
    }
  }

  // -----------------------------------------------------------
  // 6. ACTUALIZACIÓN DE HOJA DESTINO (Asistencia)
  // -----------------------------------------------------------
  const ultimaFila = hojaAsistencia.getLastRow();
  
  if (ultimaFila < 2) return Logger.log("⚠️ No hay registros en Asistencia_SinValores.");

  const cedulas = hojaAsistencia.getRange(2, colCedulaAsistencia, ultimaFila - 1, 1).getValues();
  
  // Crear array de valores para escribir
  const valores = cedulas.map(([cedula]) => {
    const c = String(cedula || "").replace(/\D/g, "").trim();
    return [mapaHoras[c] ? mapaHoras[c].horas : ""];
  });

  hojaAsistencia.getRange(2, colNueva, valores.length, 1).setValues(valores);

  Logger.log(`✅ Columna '${nombreColumnaNueva}' actualizada para cálculo de Horas Extras (${valores.length} filas).`);
  
  SpreadsheetApp.getActive().toast("✅ Horas base sincronizadas.", "Sistema Horas Extras", 5);
}
