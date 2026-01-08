// ===================================================================
// 🧪 test_unificado.gs – Suite de Pruebas Unificada (NASE 2026)
// -------------------------------------------------------------------
/**
 * @summary Módulo de Pruebas y Diagnóstico para Administradores/Desarrolladores.
 * @description Unifica todas las funciones de prueba de versiones anteriores.
 *              Este archivo contiene herramientas para estresar el sistema,
 *              comparar proveedores de geocodificación y diagnosticar datos.
 *
 * @features
 *   - 🏁 **Simulación de Concurrencia:** Genera registros masivos (Ej: 200 empleados,
 *     30 días) para probar la velocidad de escritura y triggers de hoja_turnos.gs.
 *   - 🗺️ **Comparación de APIs:** Consulta simultáneamente Maps.co y Nominatim
 *     para ver cuál devuelve mejor información (Ciudad/Dirección).
 *   - 🔍 **Diagnóstico de Datos:** Inspecciona la hoja "Respuestas" para verificar
 *     que las fechas se guarden como "Strings" (dd/mm/yyyy) y no como "Objects" Date
 *     de Google Sheets (esto evita errores de visualización o fuso horario).
 *   - 🛡️ **Pruebas de Lógica:** Verifica la lógica de secuencias (Entrada->Salida).
 *
 * @usage
 *   Estas funciones están diseñadas para ejecutarse **manualmente desde el Editor de Script**
 *   (Run) o desde una función personalizada. No se recomienda asignarlas a un Trigger
 *   automático por la carga masiva que generan.
 *
 * @author NASE Team
 * @version 1.0 (Unificado)
 */

// ===================================================================
// 1. PRUEBAS DE SIMULACIÓN DE CARGA
// ===================================================================

/**
 * @summary Simula alta concurrencia de registros (Entradas y Salidas con descansos).
 * @description Genera "Empleados Simulados" y registra sus asistencias
 *              en la hoja "Respuestas".
 * 
 * @scenario
 *   - Cantidad de Empleados: 200.
 *   - Cantidad de Registros por día (Turnos): 4 (Ej: Entrada -> Salida -> Entrada -> Salida).
 *   - Días de simulación: 30.
 *   - Retardo entre registros: 250ms (Para saturar el trigger de hoja_turnos.gs).
 * 
 * @workflow
 *   1. Genera datos aleatorios pero consistentes (Cédula, Ciudad, Centro, Coordenadas).
 *   2. Escribe las filas en la hoja "Respuestas".
 *   3. Al final, muestra un Alert con el tiempo total de duración.
 * 
 * @warning No ejecutar en producción con datos reales. Borra o ensucia la hoja "Respuestas".
 *           Si se desea probar con datos reales, eliminar la función de limpieza al inicio.
 */
function simularConcurrenciaConBreaks() {
  const NUM_EMPLEADOS = 200; // Ajustar a 1000 si quieres probar extrema
  const REGISTROS_POR_DIA = 4; // Simula un empleado que entra/sale 2 veces al día
  const DIAS = 30; // Simula 1 mes completo
  const DELAY_MS = 250; // Retardo artificial para estresar el servidor

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Respuestas');
  if (!hoja) {
    SpreadsheetApp.getUi().alert('❌ No existe la hoja Respuestas');
    return;
  }

  const inicio = new Date();
  
  // Datos Mock para simular diferentes centros
  const ciudades = ['PEREIRA', 'DOSQUEBRADAS', 'CALI', 'CARTAGO', 'BOGOTA'];
  const centros = [
    'PEREIRA ADMINISTRACION',
    'PEREIRA MEGABUS',
    'CALI IDIME CAMBULOS',
    'BOGOTA CLINICA NUEVA EL LAGO',
    'CARTAGO CLINICA NUEVA DE CARTAGO'
  ];

  // Generar el pool de empleados
  const empleados = generarEmpleadosConBreaks(NUM_EMPLEADOS, REGISTROS_POR_DIA, DIAS);

  // Bucle de escritura masiva
  for (let i = 0; i < empleados.length; i++) {
    registrarSimulado(hoja, empleados[i]);
    
    // Retraso artificial para saturar el sistema
    Utilities.sleep(DELAY_MS);
  }

  const duracion = ((new Date()) - inicio) / 1000;
  SpreadsheetApp.getUi().alert(
    `✅ Simulación completada\n` +
    `Empleados Simulados: ${empleados.length}\n` +
    `Duración: ${duracion.toFixed(1)} seg`
  );
}

// -------------------------------------------------------------------
// GENERADORES DE DATOS MOCK (HELPER)
// -------------------------------------------------------------------

/**
 * @summary Genera empleados con múltiples registros (simula turnos).
 * @description Crea objetos de empleados con cédulas únicas y arrays
 *              de registros (Entradas/Salidas) distribuidos en 30 días.
 * 
 * @param {Number} cantidad - Cantidad de empleados a generar.
 * @param {Number} registrosPorDia - Cantidad de turnos por día.
 * @param {Number} dias - Cantidad de días a simular.
 * @returns {Array<Object>} Array de objetos empleado con propiedades de turno.
 */
function generarEmpleadosConBreaks(cantidad, registrosPorDia, dias) {
  // Coordenadas base para ciudades (Lat, Lng)
  const coords = {
    PEREIRA: [4.8087, -75.6900],
    DOSQUEBRADAS: [4.8345, -75.6671],
    CALI: [3.4372, -76.5225],
    CARTAGO: [4.7464, -75.9117],
    BOGOTA: [4.7110, -74.0721]
  };

  const arr = [];
  for (let i = 0; i < cantidad; i++) {
    const ciudad = ciudades[i % ciudades.length];
    const centro = centros[i % centros.length];
    const [baseLat, baseLng] = coords[ciudad];

    // Bucle anidado para crear turnos diarios
    for (let d = 0; d < dias; d++) {
      for (let r = 0; r < registrosPorDia; r++) {
        const fecha = new Date();
        fecha.setDate(fecha.getDate() - d); // Restar días para simular pasado
        
        // Simular hora (6 AM + 4 horas * turno) -> 6:00, 10:00, 14:00, 18:00
        const hora = 6 + r * 4; 

        arr.push({
          cedula1: 1000000000 + i, // Cédula falsa única
          tipo: r % 2 === 0 ? 'Entrada' : 'Salida', // Alterna Entrada/Salida
          centro: centro,
          ciudad: ciudad,
          // Añadir ligera variación aleatoria a coordenadas
          lat: baseLat + (Math.random() - 0.5) * 0.01,
          lng: baseLng + (Math.random() - 0.5) * 0.01,
          acepto: "Sí"
        });
      }
    }
  }
  return arr;
}

/**
 * @summary Registra una fila simulada en la hoja Respuestas.
 * @description Crea una fila con estructura compatible con el sistema actual.
 *              Llena columnas vacías con strings vacíos para mantener alineación.
 * 
 * @param {Sheet} hoja - Hoja "Respuestas".
 * @param {Object} emp - Objeto con datos simulados.
 */
function registrarSimulado(hoja, emp) {
  try {
    hoja.appendRow([
      new Date(), // Timestamp
      emp.cedula1,
      emp.tipo,
      emp.centro,
      emp.ciudad,
      emp.lat,
      emp.lng,
      emp.acepto,
      '', '', '', '', '', // Columnas vacías (Geo, Obs, etc.)
      '', '', '', ''  // Más columnas vacías
      // Fecha y Hora Entrada/Salida (se dejan vacíos en esta versión simple,
      // pero se inyectan a continuación si se requiere precisión de fecha simulada)
    ]);
  } catch (e) {
    Logger.log(`❌ Error registrando ${emp.cedula1}: ${e.message}`);
  }
}

// ===================================================================
// 2. PRUEBAS DE GEOCODIFICACIÓN (APIs)
// ===================================================================

/**
 * @summary Compara fuentes de geocodificación (Maps.co vs Nominatim).
 * @description Llama a ambas APIs con las mismas coordenadas para comparar
 *              la calidad del resultado (Nombre de ciudad, Dirección).
 * 
 * @test Caso de prueba: Latitud de Pereira.
 */
function probarFuentesDeGeocodificacion() {
  const lat = 4.8087;
  const lng = -75.6900;

  const resultados = [];

  // Prueba con Maps.co
  try {
    const maps = testFromMapsCo(lat, lng);
    resultados.push({
      fuente: 'Maps.co',
      ciudad: maps?.ciudad || 'No encontrado',
      direccion: maps?.dirOriginal || 'Sin dirección'
    });
  } catch (e) {
    resultados.push({ fuente: 'Maps.co', error: e.message });
  }

  // Prueba con Nominatim
  try {
    const nom = testFromNominatim(lat, lng);
    resultados.push({
      fuente: 'Nominatim',
      ciudad: nom?.ciudad || 'No encontrado',
      direccion: nom?.dirOriginal || 'Sin dirección'
    });
  } catch (e) {
    resultados.push({ fuente: 'Nominatim', error: e.message });
  }

  // Mostrar resultados en consola
  Logger.log('=== RESULTADOS DE GEOCODIFICACIÓN ===');
  resultados.forEach(r => {
    Logger.log(`${r.fuente}: ${r.ciudad} | ${r.direccion || r.error}`);
  });

  SpreadsheetApp.getUi().alert(
    '✅ Prueba de geocodificación completada. Ver registros para resultados.'
  );
}

/**
 * @summary Consulta la API de Maps.co.
 * @private
 */
function testFromMapsCo(lat, lng) {
  const url = `https://geocode.maps.co/reverse?lat=${lat}&lon=${lng}&accept-language=es`;
  const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());

  if (!j || !j.display_name) return null;

  const a = j.address || {};
  // Prioridad: Ciudad > Town > Village > County
  const ciudad = a.city || a.town || a.village || a.county || "No encontrado";
  
  const dirOriginal = [
    a.road, a.house_number, a.suburb, a.neighbourhood, a.name
  ].filter(Boolean).join(", ");

  return { ciudad, dirOriginal };
}

/**
 * @summary Consulta la API de Nominatim.
 * @private
 */
function testFromNominatim(lat, lng) {
  const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${lat}&lon=${lng}&addressdetails=1&accept-language=es`;
  const r = UrlFetchApp.fetch(url, {
    headers: { 'User-Agent': 'NASE-Test/1.0 (contact: analistaoperaciones@nasecolombia.com.co)' },
    muteHttpExceptions: true
  });
  const j = JSON.parse(r.getContentText());
  if (!j || !j.address) return null;

  const a = j.address;
  const ciudad = a.city || a.town || a.village || a.state || "No encontrado";
  const dirOriginal = [a.road, a.house_number, a.suburb, a.neighbourhood].filter(Boolean).join(", ");
  
  return { ciudad, dirOriginal };
}

// ===================================================================
// 3. DIAGNÓSTICO DEL SISTEMA
// ===================================================================

/**
 * @summary Diagnóstico rápido de la estructura de hojas y datos.
 * @description Itera sobre todas las hojas del libro actual y muestra:
 *              - Nombre de la hoja.
 *              - Cantidad de filas.
 *              - Cantidad de columnas.
 *              - Nombres de encabezados.
 *              - Contenido de la primera fila de datos (para inspección visual).
 * 
 * @usage Ejecutar antes de iniciar una migración o cuando se sospechan errores de columnas.
 */
function diagnosticoRapido() {
  Logger.log('=== DIAGNÓSTICO RÁPIDO NASE ===');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('📁 Libro: ' + ss.getName());

    const hojas = ss.getSheets();
    hojas.forEach(hoja => {
      const nombre = hoja.getName();
      const filas = hoja.getLastRow();
      const columnas = hoja.getLastColumn();
      Logger.log(`📄 Hoja: ${nombre} | Filas: ${filas} | Columnas: ${columnas}`);

      if (filas > 1) {
        const headers = hoja.getRange(1, 1, 1, columnas).getValues()[0];
        Logger.log(`   Encabezados: ${headers.join(' | ')}`);
        
        // Mostrar primeras 5 columnas de la primera fila de datos
        const primeraFila = hoja.getRange(2, 1, 1, Math.min(columnas, 5)).getValues()[0];
        Logger.log(`   Primera fila (primeras 5 cols): ${primeraFila.join(' | ')}`);
      }
    });
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
  }

  SpreadsheetApp.getUi().alert('✅ Diagnóstico completado. Ver registros para resultados.');
}

// ===================================================================
// 4. PRUEBAS ESPECÍFICAS (Lógica y Claves)
// ===================================================================

/**
 * @summary Prueba la generación de Clave Temporal (Si existe en el sistema).
 * @description Lee la última fila de la hoja de asistencia y verifica
 *              el formato de la columna auxiliar de clave.
 */
function testClaveTmp() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  const ultima = sh.getRange(sh.getLastRow(), 1, 1, sh.getLastColumn()).getValues()[0];
  Logger.log('Última fila: ' + ultima[21]); // debe mostrar "cedula_fecha_tipo"
}

/**
 * @summary Test de Secuencia de Turnos (Lógica del Sistema).
 * @description Simula la validación del servidor (Backend) para verificar
 *              si rechaza correctamente turnos inconsistentes.
 * 
 * @logic
 * - 1. Intenta validar una 'salida' cuando es la primera vez (Fallo esperado).
 * - 2. Valida una 'entrada' cuando es la primera vez (Éxito esperado).
 */
function testSecuenciaCorregida() {
  const cedula = '12345678';
  
  // Test 1: Intentar salir sin haber entrado (Primera vez -> Debe fallar)
  let res1 = validarSecuenciaFront(cedula, 'salida');
  Logger.log(`Test 1 (primer registro salida): ${res1.esValido ? '❌ FALLO' : '✅ OK'} - ${res1.message}`);
  
  // Test 2: Entrar por primera vez (Primera vez -> Debe ser OK)
  let res2 = validarSecuenciaFront(cedula, 'entrada');
  Logger.log(`Test 2 (primer registro entrada): ${res2.esValido ? '✅ OK' : '❌ FALLO'}`);
}

// ===================================================================
// 5. PRUEBA PROFUNDA DE CONSULTA (Tipos de Datos)
// ===================================================================

/**
 * @summary Test de consulta para inspeccionar tipos de datos (Strings vs Objects).
 * @description Este es el test más importante para depurar errores visuales.
 *              Google Sheets convierte las fechas internamente a Objetos Date.
 *              Si el frontend recibe un Objeto y trata de hacer `.split('/')`, fallará.
 *              Este test verifica qué está llegando al backend.
 * 
 * @steps
 *   1. Lee datos crudos (`getValues()`) para ver qué guarda Sheets.
 *   2. Identifica una fila de prueba (Ej: Cédula 52462638).
 *   3. Analiza el Tipo de la Fecha Entrada (`typeof`).
 *   4. Si es 'object', es una Fecha nativa de Google Sheets.
 *   5. Ejecuta `obtenerRegistros` con un rango específico para ver qué devuelve al HTML.
 */
function testConsulta() {
  Logger.log("================== TEST DE CONSULTA ==================");
  
  try {
    // 1. LEER DATOS CRUDOS DIRECTAMENTE DE LA HOJA
    // Esto nos permite ver QUÉ tipo de datos (String, Date) está guardando Sheets
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('Respuestas');
    
    if (!sh) {
      Logger.log("❌ Hoja 'Respuestas' no encontrada");
      return;
    }
    
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("❌ La hoja está vacía (solo encabezados)");
      return;
    }
    
    // Buscamos tu fila de prueba para inspeccionar los tipos
    // Índice 0 = Cédula, Índice 14 = Fecha Entrada, Índice 15 = Hora Entrada
    const cedulaBuscada = '52462638';
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[0]) === cedulaBuscada) {
        Logger.log(`✅ ENCONTRADA Fila ${i+1}:`);
        Logger.log(`  Cédula: ${row[0]}`);
        
        // --- INSPECCIÓN FECHA ---
        const rawFecha = row[14];
        const tipoFecha = typeof rawFecha;
        Logger.log(`  [Índice 14] Fecha Entrada Tipo: ${tipoFecha}, Valor: ${rawFecha}`);
        
        // ⚠️ ADVERTENCIA CRÍTICA:
        // Si es tipo Date, Sheets lo parsea como Mon Jan 02 2026...
        // Si intentamos hacer row[14].split('/') en el backend sin convertir,
        // el sistema se romperá.
        if (tipoFecha === 'object') {
           Logger.log(`    ⚠️ ADVERTENCIA: La fecha es un OBJETO DATE nativo. Esto puede causar problemas con .split('/')`);
           Logger.log(`    Intentando convertirlo manualmente a dd/mm/yyyy...`);
           const manualStr = Utilities.formatDate(rawFecha, TZ, "dd/MM/yyyy");
           Logger.log(`    Resultado conversión: "${manualStr}"`);
        }

        // --- INSPECCIÓN HORA ---
        const rawHora = row[15];
        const tipoHora = typeof rawHora;
        Logger.log(`  [Índice 15] Hora Entrada Tipo: ${tipoHora}, Valor: ${rawHora}`);
        
        // Si es tipo Date, suele venir con 1899-12-30 si solo es hora
        if (tipoHora === 'object' && String(rawHora).includes('1899')) {
           Logger.log(`    ⚠️ ADVERTENCIA: La hora tiene '1899-12-30'. Esto romperá new Date() si no se limpia.`);
        }
      }
    }

    // 2. EJECUTAR LA FUNCIÓN PROBLEMA (obtenerRegistros)
    // Ponemos un rango amplio de fechas de 2026 para asegurarnos que filtre algo
    const filtros = {
      fechaInicio: '2026-01-01', // 1 Enero 2026
      fechaFin: '2026-12-31'    // 31 Diciembre 2026
    };

    Logger.log(`--- Llamando obtenerRegistros con filtros: ${filtros.fechaInicio} a ${filtros.fechaFin} ---`);
    
    const resultado = obtenerRegistros(filtros);
    
    Logger.log(`--- RESULTADO ---`);
    Logger.log(`Status: ${resultado.status}`);
    Logger.log(`Cantidad Registros: ${resultado.registros ? resultado.registros.length : 'NULL'}`);
    
    if (resultado.registros && resultado.registros.length > 0) {
      Logger.log(`✅ REGISTRO ENCONTRADO: ${resultado.registros[0].nombre} (${resultado.registros[0].fechaEntrada})`);
    } else {
      Logger.log(`❌ NO HAY REGISTROS.`);
      Logger.log(`Revisa los mensajes de advertencia (ADVERTENCIA) arriba en el log.`);
    }
    
    Logger.log("========================================================");
  } catch (e) {
    Logger.log(`❌ ERROR EN TEST CONSULTA: ${e.toString()}`);
  }
}
