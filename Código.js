/**
 * =========================================================================
 * SUNARP - SISTEMA DE CONTROL E INVENTARIO DE TÍTULOS (BACKEND OPTIMIZADO)
 * =========================================================================
 * Arquitectura de Alto Rendimiento:
 *  - Cero I/O superfluo: lecturas y escrituras atómicas en bloque (Batch Operations).
 *  - Integridad transaccional con LockService (evita duplicados por concurrencia).
 *  - Búsqueda nativa TextFinder y caché de usuarios con CacheService.
 *  - 100% compatible con la base de datos de 16 columnas y con index.html.
 */

const SHEET_TITULOS = "INVENTARIO_TITULOS";
const SHEET_USUARIOS = "USUARIOS";
const SHEET_HISTORIAL = "HISTORIAL";
const TIMEZONE_PERU = "America/Lima";

/**
 * Sirve la interfaz web HTML
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistema de Inventario y Control de Títulos - SUNARP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Helper para obtener la hoja de cálculo activa
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Normaliza el texto del estado respetando las reglas de negocio
 */
function normalizeEstadoGAS(val) {
  if (!val) return 'Recepcionado';
  var str = String(val).trim().toUpperCase();
  if (str === 'ENTREGADO') return 'Entregado';
  if (str === 'RECEPCIONADO') return 'Recepcionado';
  if (str === 'ELIMINADO') return 'Eliminado';
  if (str.indexOf('PENDIENTE') !== -1) return 'Pendiente Eliminación';
  return String(val).trim();
}

/**
 * Inicialización y verificación de la estructura de 16 columnas
 */
function setupDatabaseSheets() {
  const ss = getSpreadsheet();
  
  // A. Hoja INVENTARIO_TITULOS (16 Columnas oficiales)
  let sheetInv = ss.getSheetByName(SHEET_TITULOS);
  if (!sheetInv) {
    sheetInv = ss.insertSheet(SHEET_TITULOS);
    sheetInv.appendRow([
      'ID Registro', 'Año', 'Número de título', 'Tipo de título', 
      'Fecha de presentación', 'Fecha de inscripción o tachado', 'Fecha y hora de recepción', 
      'Estado', 'Área remitente', 'Servidor que recepciona', 'Observaciones', 
      'Fecha registro del sistema', 'Usuario registro', 'Última modificación', 'Usuario modificación',
      'Fecha de Entrega'
    ]);
    sheetInv.getRange("1:1").setFontWeight("bold").setBackground("#1b365d").setFontColor("#ffffff");
  } else {
    const lastCol = Math.max(sheetInv.getLastColumn(), 16);
    const headers = sheetInv.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.length < 16 || String(headers[15] || '').trim() === '') {
      sheetInv.getRange(1, 16).setValue('Fecha de Entrega').setFontWeight("bold").setBackground("#1b365d").setFontColor("#ffffff");
    }
  }
  
  // B. Hoja USUARIOS
  let sheetUsr = ss.getSheetByName(SHEET_USUARIOS);
  if (!sheetUsr) {
    sheetUsr = ss.insertSheet(SHEET_USUARIOS);
    sheetUsr.appendRow(['ID', 'Username', 'Password', 'Nombre Completo', 'Rol', 'Área', 'Estado']);
    sheetUsr.appendRow(['USR-001', 'admin', 'admin123', 'Administrador Principal', 'Administrador', 'Mesa de Partes', 'Activo']);
    sheetUsr.getRange("1:1").setFontWeight("bold").setBackground("#1b365d").setFontColor("#ffffff");
  }
  
  // C. Hoja HISTORIAL (Auditoría)
  let sheetHist = ss.getSheetByName(SHEET_HISTORIAL);
  if (!sheetHist) {
    sheetHist = ss.insertSheet(SHEET_HISTORIAL);
    sheetHist.appendRow(['FechaHora', 'Usuario', 'Acción', 'ID Registro', 'Campo Modificado', 'Valor Anterior', 'Valor Nuevo']);
    sheetHist.getRange("1:1").setFontWeight("bold").setBackground("#1b365d").setFontColor("#ffffff");
  }

  return "Estructura de Base de Datos (16 columnas) verificada con éxito.";
}

/**
 * Carga unificada en 1 solo viaje HTTP cliente-servidor
 */
function getDatosCompletosGAS() {
  return {
    titulos: getTitulosGAS(),
    usuarios: getUsuariosGAS()
  };
}

/**
 * Lectura ultrarrápida de títulos en bloque
 */
function getTitulosGAS() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TITULOS);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return [];

    // 1 sola llamada API de lectura para toda la hoja
    const data = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 16)).getValues();
    const titulos = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1] && !row[2]) continue;

      const fechaEntregaVal = row[15] ? formatDateFull(row[15]) : '';
      let estadoCalc = normalizeEstadoGAS(row[7]);
      if (fechaEntregaVal !== '' && estadoCalc !== 'Eliminado') {
        estadoCalc = 'Entregado';
      }

      titulos.push({
        idRegistro: String(row[0] || '').trim(),
        ano: row[1] ? Number(row[1]) : 2026,
        numeroTitulo: String(row[2] || '').trim(),
        tipoTitulo: String(row[3] || 'Inscrito').trim(),
        fechaPresentacion: formatDateISO(row[4]),
        fechaInscripcionTachado: formatDateISO(row[5]),
        fechaRecepcion: formatDateFull(row[6]),
        estado: estadoCalc,
        areaRemitente: String(row[8] || '').trim(),
        servidorRecepciona: String(row[9] || '').trim(),
        observaciones: String(row[10] || '').trim(),
        fechaRegistro: formatDateFull(row[11]),
        usuarioRegistro: String(row[12] || 'Sistema').trim(),
        ultimaModificacion: formatDateFull(row[13]),
        usuarioModificacion: String(row[14] || 'Sistema').trim(),
        fechaEntrega: fechaEntregaVal
      });
    }

    // Invertir para mostrar primero los más recientes en la tabla
    return titulos.reverse();
  } catch (e) {
    Logger.log("Error en getTitulosGAS: " + e.message);
    return [];
  }
}

/**
 * Obtiene usuarios con aceleración CacheService
 */
function getUsuariosGAS() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get("CACHE_USERS_SUNARP");
    if (cached) {
      return JSON.parse(cached);
    }

    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_USUARIOS);
    if (!sheet) {
      setupDatabaseSheets();
      sheet = ss.getSheetByName(SHEET_USUARIOS);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const usuarios = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue;

      usuarios.push({
        id: String(row[0] || '').trim(),
        username: String(row[1] || '').trim(),
        password: String(row[2] || '').trim(),
        nombre: String(row[3] || 'Usuario').trim(),
        rol: String(row[4] || 'Usuario').trim(),
        area: String(row[5] || 'Mesa de Partes').trim(),
        estado: String(row[6] || 'Activo').trim()
      });
    }

    // Cachear usuarios por 30 minutos (1800 s)
    cache.put("CACHE_USERS_SUNARP", JSON.stringify(usuarios), 1800);
    return usuarios;
  } catch (e) {
    Logger.log("Error en getUsuariosGAS: " + e.message);
    return [];
  }
}

/**
 * Valida credenciales de acceso de forma instantánea
 */
function validarLoginGAS(usernameInput, passwordInput) {
  try {
    const usuarios = getUsuariosGAS();
    const usrClean = String(usernameInput || '').trim().toLowerCase();
    const pwdClean = String(passwordInput || '').trim();

    for (let i = 0; i < usuarios.length; i++) {
      const u = usuarios[i];
      const matchesUser = (u.username.toLowerCase() === usrClean || u.id.toLowerCase() === usrClean);
      const matchesPass = (u.password === pwdClean);

      if (matchesUser && matchesPass) {
        if (u.estado && u.estado.toLowerCase() !== 'activo') {
          return { success: false, message: "Su cuenta se encuentra INACTIVA o en VACACIONES. Contacte con la Jefatura." };
        }

        logAuditGAS('Inicio de Sesión', u.id, 'Autenticación', '-', 'Ingreso correcto al sistema', u.nombre);
        return {
          success: true,
          user: {
            id: u.id,
            username: u.username,
            nombre: u.nombre,
            rol: u.rol,
            area: u.area
          }
        };
      }
    }

    return { success: false, message: "Usuario o contraseña incorrectos." };
  } catch (err) {
    return { success: false, message: "Error al validar acceso: " + err.message };
  }
}

/**
 * Registra un nuevo título con ScriptLock para evitar duplicidad de ID o colisiones
 */
function registrarTituloGAS(record) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 15 segundos de tolerancia
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_TITULOS);
    if (!sheet) {
      setupDatabaseSheets();
      sheet = ss.getSheetByName(SHEET_TITULOS);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // Verificación de duplicado en bloque
      const existingData = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
      const targetAno = Number(record.ano);
      const targetNum = String(record.numeroTitulo).trim().toLowerCase();

      for (let i = 0; i < existingData.length; i++) {
        if (Number(existingData[i][0]) === targetAno && String(existingData[i][1]).trim().toLowerCase() === targetNum) {
          return { success: false, message: 'DUPLICADO: Ya existe un registro para el Año ' + record.ano + ' y Número de título ' + record.numeroTitulo };
        }
      }
    }

    const count = lastRow;
    const newId = 'TIT-' + record.ano + '-' + String(count).padStart(4, '0');
    const now = getPeruTimestamp();
    const fechaHoraRecepcion = record.fechaRecepcion || now;

    const rowData = [
      newId,
      record.ano,
      record.numeroTitulo,
      record.tipoTitulo,
      record.fechaPresentacion,
      record.fechaInscripcionTachado,
      fechaHoraRecepcion,
      'Recepcionado',
      record.areaRemitente,
      record.servidorRecepciona,
      record.observaciones || '',
      now,
      record.usuarioRegistro || 'Sistema',
      now,
      record.usuarioRegistro || 'Sistema',
      record.fechaEntrega || ''
    ];

    sheet.appendRow(rowData);
    logAuditGAS('Registro creado', newId, 'Todos', '-', 'Nuevo Título ' + record.numeroTitulo, record.usuarioRegistro);

    return { success: true, idRegistro: newId };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Edición ultrarrápida: escribe todas las columnas en 1 solo setValues() continuo
 */
function editarTituloGAS(record, usuario) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TITULOS);
    if (!sheet) return { success: false, message: "No existe la hoja INVENTARIO_TITULOS" };

    const targetId = String(record.idRegistro || '').trim();
    let rowNumber = -1;

    // 1. Búsqueda por ID con TextFinder
    if (targetId) {
      const cell = sheet.getRange("A:A").createTextFinder(targetId).matchEntireCell(true).findNext();
      if (cell) rowNumber = cell.getRow();
    }

    // 2. Fallback por coincidencia de Año y Número de Título
    if (rowNumber === -1) {
      const data = sheet.getDataRange().getValues();
      const targetAno = Number(record.ano || 0);
      const targetNum = String(record.numeroTitulo || '').trim().toLowerCase();

      for (let i = 1; i < data.length; i++) {
        if (Number(data[i][1]) === targetAno && String(data[i][2]).trim().toLowerCase() === targetNum) {
          rowNumber = i + 1;
          break;
        }
      }
    }

    if (rowNumber === -1) {
      return { success: false, message: "No se encontró el registro a editar: " + record.idRegistro };
    }

    // Obtener la fila actual completa para preservar fechas originales de registro
    const currentRow = sheet.getRange(rowNumber, 1, 1, 16).getValues()[0];
    const prevEstado = currentRow[7];
    const estadoNorm = normalizeEstadoGAS(record.estado);
    const now = getPeruTimestamp();

    // Actualizar columnas en memoria
    currentRow[3] = record.tipoTitulo;
    currentRow[4] = record.fechaPresentacion;
    currentRow[5] = record.fechaInscripcionTachado;
    currentRow[6] = record.fechaRecepcion;
    currentRow[7] = estadoNorm;
    currentRow[8] = record.areaRemitente;
    currentRow[9] = record.servidorRecepciona;
    currentRow[10] = record.observaciones;
    currentRow[13] = now;
    currentRow[14] = usuario || "Sistema";

    if (record.fechaEntrega !== undefined && record.fechaEntrega !== '') {
      currentRow[15] = record.fechaEntrega;
    }

    // 1 sola escritura en bloque para toda la fila
    sheet.getRange(rowNumber, 1, 1, 16).setValues([currentRow]);
    logAuditGAS('Registro modificado', currentRow[0], 'Estado/Datos', prevEstado, estadoNorm, usuario);

    return { success: true, estadoActualizado: estadoNorm };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Eliminación de títulos en lote: 1 sola escritura a Google Sheets
 */
function ejecutarEliminacionGAS(selectedIds, oficio, fechaElim, aprobador, usuario) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TITULOS);
    if (!sheet) return { success: false, eliminados: 0 };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, eliminados: 0 };

    const cleanIdsSet = new Set(selectedIds.map(id => String(id).trim().toLowerCase().replace(/[^a-z0-9]/g, '')));
    const now = getPeruTimestamp();
    let count = 0;

    // Modificación 100% en memoria
    for (let i = 1; i < data.length; i++) {
      const currentId = String(data[i][0] || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      if (cleanIdsSet.has(currentId)) {
        const prevObs = data[i][10] || '';
        data[i][7] = 'Eliminado';
        data[i][10] = prevObs + ' | ELIMINADO según ' + oficio + ' por ' + aprobador + ' el ' + fechaElim;
        data[i][13] = now;
        data[i][14] = usuario || 'Sistema';
        count++;
      }
    }

    // 1 sola llamada API para volcar los cambios completos
    if (count > 0) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      logAuditGAS('Eliminación administrativa', 'LOTE (' + count + ')', 'Estado', 'Recepcionado', 'ELIMINADO (' + oficio + ')', usuario);
    }

    return { success: true, eliminados: count };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Crea nuevo usuario e invalida la caché
 */
function crearUsuarioGAS(userRecord, usuarioOperador) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_USUARIOS);
    if (!sheet) {
      setupDatabaseSheets();
      sheet = ss.getSheetByName(SHEET_USUARIOS);
    }

    const data = sheet.getDataRange().getValues();
    const usrClean = String(userRecord.username).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === usrClean) {
        return { success: false, message: "El nombre de usuario '" + userRecord.username + "' ya existe en el sistema." };
      }
    }

    const newId = 'USR-' + String(data.length).padStart(3, '0');
    sheet.appendRow([
      newId, userRecord.username, userRecord.password, userRecord.nombre,
      userRecord.rol, userRecord.area, userRecord.estado || 'Activo'
    ]);

    // Limpiar caché para reflejar el nuevo usuario
    CacheService.getScriptCache().remove("CACHE_USERS_SUNARP");
    logAuditGAS('Usuario Creado', newId, 'USUARIOS', '-', userRecord.nombre + ' (' + userRecord.rol + ')', usuarioOperador);

    return { success: true, id: newId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Modifica rol, área y estado de usuario e invalida la caché
 */
function modificarUsuarioGAS(userRecord, usuarioOperador) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USUARIOS);
    if (!sheet) return { success: false, message: "La hoja USUARIOS no está configurada." };

    const data = sheet.getDataRange().getValues();
    const targetId = String(userRecord.id).trim().toLowerCase();
    let rowIdx = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetId) {
        rowIdx = i + 1;
        const prevRol = data[i][4];
        const prevEstado = data[i][6];

        // Escribir fila en 1 sola llamada
        sheet.getRange(rowIdx, 5, 1, 3).setValues([[userRecord.rol, userRecord.area, userRecord.estado]]);
        CacheService.getScriptCache().remove("CACHE_USERS_SUNARP");

        logAuditGAS('Modificación de Usuario', userRecord.id, 'Rol/Área/Estado', 
          'Rol: ' + prevRol + ', Estado: ' + prevEstado, 
          'Rol: ' + userRecord.rol + ', Estado: ' + userRecord.estado, 
          usuarioOperador);

        return { success: true };
      }
    }

    return { success: false, message: "No se encontró el usuario a modificar: " + userRecord.id };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Registra entradas de trazabilidad en la hoja HISTORIAL
 */
function logAuditGAS(accion, idRegistro, campo, anterior, nuevo, usuario) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_HISTORIAL);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_HISTORIAL);
      sheet.appendRow(['FechaHora', 'Usuario', 'Acción', 'ID Registro', 'Campo Modificado', 'Valor Anterior', 'Valor Nuevo']);
      sheet.getRange("1:1").setFontWeight("bold").setBackground("#1b365d").setFontColor("#ffffff");
    }
    const now = getPeruTimestamp();
    sheet.appendRow([now, usuario || 'Sistema', accion, idRegistro, campo, anterior, nuevo]);
  } catch (e) {
    Logger.log("Error logAuditGAS: " + e.message);
  }
}

function getPeruTimestamp() {
  return Utilities.formatDate(new Date(), TIMEZONE_PERU, "yyyy-MM-dd HH:mm:ss");
}

function formatDateISO(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE_PERU, "yyyy-MM-dd");
  }
  let str = String(val).trim();
  if (str.indexOf(' ') !== -1) str = str.split(' ')[0];
  if (str.indexOf('T') !== -1) str = str.split('T')[0];
  return str;
}

function formatDateFull(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE_PERU, "yyyy-MM-dd HH:mm:ss");
  }
  return String(val).trim();
}