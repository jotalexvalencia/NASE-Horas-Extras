// ============================================================
// 🧹 limpieza_registros.gs – Reportes de Nómina (NASE 2026 - Horas Extras)
// ------------------------------------------------------------
/**
 * @summary Módulo de Filtrado y Generación de Reportes de Nómina.
 * @description Filtra registros de la hoja "Respuestas" basándose en reglas
 *              temporales y los copia a una hoja "Filtrado" para su exportación.
 * 
 * @safety
 *   - 🛡️ NO BORRA datos: La hoja principal ("Respuestas") permanece intacta.
 *   - 📄 COPIA A REPORTES: Genera una hoja "Filtrado" con los datos del periodo.
 *
 * @criteria (Reglas de Negocio para Nómina)
 *   - 🔹 Criterio 1 (Cierre Mes Anterior): Registros del último día del mes anterior
 *     que ocurrieron entre las 18:00 y las 22:00 (Turnos de cierre).
 *   - 🔹 Criterio 2 (Mes Actual): TODOS los registros del mes actual en curso.
 * 
 * @note Este reporte incluye automáticamente las columnas de aprobación de Horas Extras.
 * 
 * @author NASE Team
 * @version 1.4 (Adaptado para Horas Extras)
 */

// ======================================================================
// FUNCIÓN PRINCIPAL: Filtro y Generación
// ======================================================================

/**
 * @summary Genera reporte de asistencia filtrando por fechas.
 * @description Ejecuta la lógica de doble criterio para extraer registros relevantes
 *              para la liquidación de Horas Extras.
 * 
 * @workflow
 *   1. Abre "Respuestas" (Origen) y "Filtrado" (Destino).
 *   2. Calcula dinámicamente las fechas del mes anterior y actual.
 *   3. Agrupa por Cédula.
 *   4. Aplica Filtro 1: Cierre de mes (Anterior).
 *   5. Aplica Filtro 2: Mes actual (Todos).
 *   6. Escribe en "Filtrado".
 * 
 * @requires Hoja "Respuestas".
 */
function filtrarRegistrosUltimoDiaMesAnteriorYMesActual() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // -----------------------------------------------------------
  // 1. CONFIGURACIÓN DE HOJAS (Origen y Destino)
  // -----------------------------------------------------------
  let hojaDestino = ss.getSheetByName("Filtrado");
  
  // Si no existe la hoja destino, crearla
  if (!hojaDestino) hojaDestino = ss.insertSheet("Filtrado");

  const hojaOrigen = ss.getSheetByName("Respuestas");
  if (!hojaOrigen) {
    ui.alert("❌ No se encontró la hoja 'Respuestas'.");
    return;
  }

  // -----------------------------------------------------------
  // 2. OBTENER INDICES DE COLUMNAS (Manejo Dinámico)
  // -----------------------------------------------------------
  const datos = hojaOrigen.getDataRange().getValues();
  
  if (datos.length < 2) {
    ui.alert("⚠️ La hoja 'Respuestas' está vacía.");
    return;
  }

  // Buscar índices por nombre (Compatible con nuevas columnas HE)
  const encabezados = datos[0];
  
  const idxCedula = encabezados.indexOf("Cédula");
  const idxCentro = encabezados.indexOf("Centro");       
  const idxFechaEnt = encabezados.indexOf("Fecha Entrada");
  const idxHoraEnt = encabezados.indexOf("Hora Entrada");
  const idxFechaSal = encabezados.indexOf("Fecha Salida");
  const idxHoraSal = encabezados.indexOf("Hora Salida");
  const idxDentroSal = encabezados.indexOf("Dentro Salida"); 
  const idxNombre = encabezados.indexOf("Nombre");             

  // Validar columnas esenciales
  if (idxCedula === -1 || idxFechaEnt === -1 || idxHoraEnt === -1) {
    ui.alert("❌ No se encontraron las columnas necesarias ('Cédula', 'Fecha Entrada', 'Hora Entrada').");
    return;
  }

  // -----------------------------------------------------------
  // 3. CÁLCULO DE FECHAS DEL SISTEMA
  // -----------------------------------------------------------
  const hoy = new Date();
  const mesActual = hoy.getMonth(); 
  const anioActual = hoy.getFullYear();
  
  const ultimoDiaMesAnterior = new Date(anioActual, mesActual, 0); 
  const diaUltimoMesAnterior = ultimoDiaMesAnterior.getDate();
  const mesAnterior = ultimoDiaMesAnterior.getMonth();
  const anioMesAnterior = ultimoDiaMesAnterior.getFullYear();

  // -----------------------------------------------------------
  // 4. AGRUPACIÓN DE REGISTROS POR CÉDULA
  // -----------------------------------------------------------
  const mapaCedulas = {};

  for (let i =1; i < datos.length; i++) {
    const fila = datos[i];
    const cedula = fila[idxCedula];
    
    if (!cedula) continue;

    // Reconstrucción robusta de fecha/hora de entrada
    const fechaRaw = fila[idxFechaEnt];
    const horaRaw = fila[idxHoraEnt];
    
    let fecha = null;

    const fechaStr = String(fechaRaw || '').trim();
    const horaStr = String(horaRaw || '').trim();

    if (fechaStr && horaStr) {
       const parts = fechaStr.split('/');
       if (parts.length === 3) {
         fecha = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T${horaStr}`);
       }
    }

    if (cedula && fecha && !isNaN(fecha.getTime())) {
      if (!mapaCedulas[cedula]) mapaCedulas[cedula] = [];
      mapaCedulas[cedula].push(fila); 
    }
  }

  const filasFinales = [];

  // -----------------------------------------------------------
  // 5. LÓGICA DE FILTRADO (Por Empleado)
  // -----------------------------------------------------------
  
  for (const cedula in mapaCedulas) {
    const filasCedula = mapaCedulas[cedula];

    // ---------------------------------------------------------
    // 🔹 CRITERIO 1: Último día del mes anterior (Horario Nocturno)
    // ---------------------------------------------------------
    const registrosUltimoDia = filasCedula.filter(fila => {
      const fechaRaw = fila[idxFechaEnt];
      const horaRaw = fila[idxHoraEnt];
      
      let fecha = null;
      const fechaStr = String(fechaRaw || '').trim();
      const horaStr = String(horaRaw || '').trim();

      if (fechaStr && horaStr) {
         const parts = fechaStr.split('/');
         if (parts.length === 3) fecha = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T${horaStr}`);
      }
      
      if (!fecha) return false;
      
      const h = parseInt(horaStr.split(':')[0], 10);
      const esNoche = (h >= 18 && h <= 22);
      
      return (
        fecha.getFullYear() === anioMesAnterior &&
        fecha.getMonth() === mesAnterior &&
        fecha.getDate() === diaUltimoMesAnterior &&
        esNoche
      );
    });

    if (registrosUltimoDia.length > 0) {
      filasFinales.push(...registrosUltimoDia);
    }

    // ---------------------------------------------------------
    // 🔹 CRITERIO 2: Registros del mes actual (Completo)
    // ---------------------------------------------------------
    const registrosMesActual = filasCedula.filter(fila => {
      const fechaRaw = fila[idxFechaEnt];
      const horaRaw = fila[idxHoraEnt];
      
      let fecha = null;
      const fechaStr = String(fechaRaw || '').trim();
      const horaStr = String(horaRaw || '').trim();

      if (fechaStr && horaStr) {
         const parts = fechaStr.split('/');
         if (parts.length === 3) fecha = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T${horaStr}`);
      }

      if (!fecha) return false;
      
      return fecha.getFullYear() === anioActual && fecha.getMonth() === mesActual;
    });

    filasFinales.push(...registrosMesActual);
  }

  // -----------------------------------------------------------
  // 6. ESCRITURA DE RESULTADO EN HOJA DESTINO
  // -----------------------------------------------------------
  
  if (filasFinales.length === 0) {
    ui.alert("❌ No se encontraron registros válidos para el filtro actual.");
    return;
  }

  hojaDestino.clearContents();
  
  // Escribir encabezados originales (incluye las nuevas columnas de HE)
  hojaDestino.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  
  // Escribir filas filtradas
  hojaDestino.getRange(2, 1, filasFinales.length, encabezados.length).setValues(filasFinales);

  // -----------------------------------------------------------
  // 7. COPIA DE SEGURIDAD (Opcional / Comentado)
  // -----------------------------------------------------------
  const nombreMeses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  const nombreMesAnterior = nombreMeses[mesAnterior];
  
  const nombreRespaldo = `registro_${nombreMesAnterior}_${anioMesAnterior}`;
  
  SpreadsheetApp.getActive().toast(
    `✅ Reporte de Horas Extras generado (${filasFinales.length} registros).`,
    "Nómina",
    5
  );
}
