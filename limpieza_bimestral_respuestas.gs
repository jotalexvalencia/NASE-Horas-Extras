// ===================================================================
// 🗄️ limpieza_bimestral_respuestas.gs – Archivo y Limpieza Inteligente (NASE 2026)
// -------------------------------------------------------------------
/**
 * @summary Módulo de Archivo Inteligente y Limpieza Bimestral de "Respuestas".
 * @description Gestiona el ciclo de vida de los registros crudos de Entrada/Salida.
 *              Evita que la hoja principal se vuelva masiva y lenta.
 *
 * @logic
 * - ⚡ **Trigger:** Se programa para ejecutarse cada 2 meses.
 * - 📅 **Rango Bimestral:** Archiva los 2 meses anteriores (Ej: Mayo archiva Marzo/Abril).
 * - 🗑️ **Limpieza:** Borra de la hoja principal los registros del periodo archivado.
 * - 🛡️ **Preservación Inteligente:** NO borra registros críticos:
 *     1. Registros Futuros (Turnos del día de mañana).
 *     2. Turnos Abiertos del último día del periodo (Ej: Entrada sin Salida del último día),
 *        para permitir que el administrador cierre manualmente esas horas.
 *
 * @correcciones (Versión Final)
 * - ✅ **Sin Timestamp:** Ya no usa ni busca columna 'Timestamp'.
 *     Reconstruye la fecha manualmente desde 'Fecha Entrada' + 'Hora Entrada'.
 * - ✅ **Sin Tipo:** No filtra por columna 'Tipo' (Entrada/Salida).
 *     Detecta "Entrada" verificando si "Fecha Salida" está vacía.
 *
 * @dependencies
 *   - `install_triggers.gs` (Función `ensureTimeTrigger`).
 *   - `Code.gs` (Encabezados compatibles con RESP_HEADERS).
 *
 * @author NASE Team
 * @version 2.1 (Algoritmo de Preservación de Turnos Abiertos)
 */

// ===================================================================
// 1. INSTALACIÓN DEL DISPARADOR (TRIGGER)
// ===================================================================

/**
 * @summary Instala el disparador bimestral para la hoja Respuestas.
 * @description Crea un Time-Based trigger.
 * 
 * @schedule
 *   - Día del mes: 1 (Cada mes 1ro).
 *   - Hora: 16 (4:00 PM).
 *   - Nota: La función interna tiene un filtro para ejecutar solo en meses pares (Feb, Abr...).
 */
function instalarTriggersLimpiezaBimestralRespuestas() {
  // Wrapper de seguridad para crear trigger si no existe
  ensureTimeTrigger("limpiarRespuestasBimestral", function () {
    ScriptApp.newTrigger("limpiarRespuestasBimestral")
      .timeBased()
      .onMonthDay(1)
      .atHour(16)
      .create();
  });
  Logger.log("✅ Trigger bimestral limpieza Respuestas instalado (NASE 2026).");
}

// ===================================================================
// 2. LÓGICA PRINCIPAL (Archivo + Limpieza)
// ===================================================================

/**
 * @summary Archiva el bimestre anterior y limpia la hoja principal.
 * @description Algoritmo complejo en 3 fases:
 *   1. **Fase Cálculo:** Determina qué 2 meses van a ser archivados.
 *   2. **Fase Detección (1er Pasada):** Busca turnos abiertos (Sin Salida)
 *      que ocurrieron en el último día del periodo. Guarda las Cédulas en un Set.
 *   3. **Fase Separación (2da Pasada):** Itera toda la hoja.
 *      - Si está en el rango archivable: Mover al archivo.
 *      - Si es un registro del último día Y ES un turno abierto (Cédula en el Set): Mover al archivo.
 *      - Si es Futuro: MANTENER en la hoja principal (Conservar).
 * 
 * @output
 *   - Archivo en Drive (Carpeta: "Archivos Respuestas Bimestrales").
 *   - Hoja "Respuestas" limpia, conservando solo datos futuros/abiertos.
 */
function limpiarRespuestasBimestral() {
  const hoy = new Date();
  const mes = hoy.getMonth(); // 0=Enero, 1=Febrero...

  // -----------------------------------------------------------------
  // 1. FILTRO DE EJECUCIÓN (Solo Meses Pares)
  // -----------------------------------------------------------------
  // La lógica `if (mes % 2 !== 0)` se ejecuta en Febrero(1), Abril(3), Junio(5)...
  // Es decir, meses IMPARES (según índice 0-based) que son PARES en calendario.
  if (mes % 2 !== 0) return; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResp = ss.getSheetByName('Respuestas');
  
  if (!hojaResp || hojaResp.getLastRow() <= 1) return;

  // -----------------------------------------------------------------
  // 2. CÁLCULO DE RANGOS (Bimestre Anterior)
  // -----------------------------------------------------------------
  // finBimestre: Último día del mes anterior (Ej: Si es May 1st, finBimestre es Abril 30)
  const finBimestre = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 0, 23, 59, 59);
  
  // inicioBimestre: Primer día del mes anterior a ese (Ej: Marzo 1st)
  const inicioBimestre = new Date(finBimestre.getFullYear(), finBimestre.getMonth() - 1, 1, 0, 0, 0);

  Logger.log(`Archivando Respuestas desde ${inicioBimestre.toLocaleDateString()} hasta ${finBimestre.toLocaleDateString()}`);

  // -----------------------------------------------------------------
  // 3. LECTURA DE DATOS (Encabezados + Filas)
  // -----------------------------------------------------------------
  const headers = hojaResp.getRange(1, 1, 1, hojaResp.getLastColumn()).getValues()[0];
  const data = hojaResp.getRange(2, 1, hojaResp.getLastRow() - 1, hojaResp.getLastColumn()).getValues();

  // -----------------------------------------------------------------
  // 4. IDENTIFICACIÓN DE COLUMNAS (Mapeo Dinámico)
  // -----------------------------------------------------------------
  // ✅ CAMBIO CRÍTICO: Dejar de buscar Timestamp ni Tipo.
  // En su lugar usar Fecha Entrada + Hora Entrada para reconstruir fecha.
  const idxCedula = headers.indexOf("Cédula");
  const idxFechaEnt = headers.indexOf("Fecha Entrada");
  const idxHoraEnt = headers.indexOf("Hora Entrada");
  const idxFechaSal = headers.indexOf("Fecha Salida");
  const idxHoraSal = headers.indexOf("Hora Salida");

  if (idxCedula === -1 || idxFechaEnt === -1 || idxHoraEnt === -1) {
    Logger.log('❌ Faltan columnas críticas en Respuestas (Cédula, Fecha Entrada, Hora Entrada)');
    return;
  }

  // -----------------------------------------------------------------
  // 5. ALGORITMO DE SEPARACIÓN (Fase de Análisis)
  // -----------------------------------------------------------------
  const datosBimestre = [headers]; // Array que contendrá lo que se va a archivar
  const datosConservar = [headers]; // Array que se quedará en la hoja (Futuros + Críticos)

  // Identificar entradas del último día sin salida (Registros Críticos)
  const ultimoDia = new Date(finBimestre.getFullYear(), finBimestre.getMonth(), finBimestre.getDate());
  const entradasUltimoDiaSinSalida = new Set(); // Set para guardar Cédulas de turnos abiertos

  // -----------------------------------------------------------------
  // PRIMER PASADA: Detección de Turnos Abiertos
  // -----------------------------------------------------------------
  // Escaneamos toda la hoja buscando registros del "Último Día del Periodo"
  // que NO tengan salida. Estos son los registros que NO debemos borrar.
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // ✅ RECONSTRUIR FECHA (Sin Timestamp)
    const fechaRaw = row[idxFechaEnt];
    const horaRaw = row[idxHoraEnt];
    let ts = null;
    
    // Parseo manual: dd/mm/yyyy HH:mm -> Date Object
    if (fechaRaw && horaRaw) {
       const parts = fechaRaw.split('/');
       if (parts.length === 3) ts = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T${horaRaw}`);
    }

    if (!ts) continue;

    // ✅ DETERMINAR TIPO (Sin Tipo columna)
    // Si no tiene fecha salida ni hora salida -> Es Entrada (Pendiente)
    const fechaSal = String(row[idxFechaSal] || '').trim();
    const horaSal = String(row[idxHoraSal] || '').trim();
    const esEntrada = (!fechaSal && !horaSal);

    // ¿Es el último día del periodo?
    if (ts.getDate() === ultimoDia.getDate() &&
        ts.getMonth() === ultimoDia.getMonth() &&
        ts.getFullYear() === ultimoDia.getFullYear() &&
        esEntrada) {
      // Verificar si tiene salida en la misma fila
      if (!fechaSal) {
        // Si NO tiene salida y ES el último día, es un TURNO ABIERTO CRÍTICO.
        // Guardamos la Cédula para no borrarla después.
        entradasUltimoDiaSinSalida.add(String(row[idxCedula]).trim());
      }
    }
  }

  // -----------------------------------------------------------------
  // SEGUNDA PASADA: Clasificación (A Archivar vs A Conservar)
  // -----------------------------------------------------------------
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // Parsear fecha de la fila
    const fechaRaw = row[idxFechaEnt];
    const horaRaw = row[idxHoraEnt];
    let ts = null;
    
    if (fechaRaw && horaRaw) {
       const parts = fechaRaw.split('/');
       if (parts.length === 3) ts = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T${horaRaw}`);
    }

    if (!ts) continue;

    // Determinar Tipo (Entrada/Salida)
    const fechaSal = String(row[idxFechaSal] || '').trim();
    const horaSal = String(row[idxHoraSal] || '').trim();
    const esEntrada = (!fechaSal && !horaSal);

    const cedula = String(row[idxCedula]).trim();

    // LÓGICA DE CLASIFICACIÓN:
    
    // 1. ARCHIVAR: Si está en el rango del bimestre (Entre inicioBimestre y finBimestre)
    if (ts >= inicioBimestre && ts <= finBimestre) {
      datosBimestre.push(row);
    } 
    // 2. CONSERVAR FUTURO: Si es posterior al fin del bimestre
    else if (ts > finBimestre) {
      datosConservar.push(row);
    } 
    // 3. CONSERVAR CRÍTICO: Si es el último día Y es una entrada abierta Y su cédula está en el Set de la Pasada 1.
    else if (ts.getDate() === ultimoDia.getDate() &&
               ts.getMonth() === ultimoDia.getMonth() &&
               ts.getFullYear() === ultimoDia.getFullYear() &&
               esEntrada &&
               entradasUltimoDiaSinSalida.has(cedula)) {
      // Este registro es antiguo (del periodo), pero es una Entrada Abierta del último día.
      // Lo mantenemos para que el admin pueda cerrarlo manualmente.
      datosConservar.push(row);
    }
  }

  // -----------------------------------------------------------------
  // 6. CREACIÓN DE ARCHIVO EN DRIVE
  // -----------------------------------------------------------------
  // Carpeta específica para históricos de Respuestas
  const folder = obtenerOCrearCarpeta('Archivos Respuestas Bimestrales');
  
  // Nombre del archivo (Ej: Respuestas_Bimestre_2025-03_2025-04)
  const nombreArchivo = `Respuestas_Bimestre_${inicioBimestre.getFullYear()}-${String(inicioBimestre.getMonth() + 1).padStart(2, '0')}_a_${finBimestre.getFullYear()}-${String(finBimestre.getMonth() + 1).padStart(2, '0')}`;
  
  // Crear el nuevo archivo y mover a la carpeta específica
  const archivo = SpreadsheetApp.create(nombreArchivo);
  DriveApp.getFileById(archivo.getId()).moveTo(folder);

  // Escribir los datos archivados en el nuevo archivo
  const hojaArchivo = archivo.getSheets()[0];
  hojaArchivo.setName('Respuestas_Archivadas');
  hojaArchivo.getRange(1, 1, datosBimestre.length, headers.length).setValues(datosBimestre);

  // -----------------------------------------------------------------
  // 7. LIMPIEZA Y RESTAURACIÓN DE HOJA PRINCIPAL
  // -----------------------------------------------------------------
  // Limpiar todo y dejar solo lo que debemos conservar
  hojaResp.clear();
  hojaResp.getRange(1, 1, datosConservar.length, headers.length).setValues(datosConservar);

  Logger.log(`✅ Respuestas archivadas: ${nombreArchivo} con ${datosBimestre.length - 1} registros. Conservados: ${datosConservar.length - 1}`);
}

// ===================================================================
// 3. UTILIDAD DE CARPETAS (DRIVE API)
// ===================================================================

/**
 * @summary Busca una carpeta por nombre en Drive. Si no existe, la crea.
 * @description Función reutilizable para organizar archivos históricos.
 * 
 * @param {String} nombre - Nombre exacto de la carpeta en Drive.
 * @returns {Folder} Objeto Carpeta de Google Drive.
 * @private
 */
function obtenerOCrearCarpeta(nombre) {
  const folders = DriveApp.getFoldersByName(nombre);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(nombre);
}
