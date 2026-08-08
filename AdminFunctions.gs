// ========================================================
// AdminFunctions.gs - Panel de Administración CRUD
// Campus Virtual Pro
// ========================================================

// --------------- HELPERS ADMIN ---------------

/**
 * Abre o crea una pestaña en la hoja de cálculo.
 * Si no existe, la crea con los headers dados.
 */
function openOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

/**
 * Verifica que el email corresponde a un administrador.
 * Lanza error si no es administrador.
 */
function ensureAdmin_(email) {
  const sheet = openSheet_('Usuarios');
  const data = sheet.getDataRange().getValues();
  const emailN = normalizeEmail_(email);
  for (let i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][1]) === emailN) {
      const role = safeStr_(data[i][5]).toLowerCase();
      if (role === 'administrador') return true;
    }
  }
  throw new Error('Acceso denegado: se requiere rol de administrador.');
}

/**
 * Inicializa las hojas de Preguntas y ConfigQuiz si no existen.
 */
function initAdminSheets() {
  openOrCreateSheet_('Preguntas', [
    'PreguntaID', 'CursoID', 'Texto', 'OpcionA', 'OpcionB',
    'OpcionC', 'OpcionD', 'RespuestaCorrecta', 'Puntos'
  ]);
  openOrCreateSheet_('ConfigQuiz', [
    'CursoID', 'TotalPreguntas', 'PuntajeMinimo',
    'IntentosMax', 'TiempoLimiteMin', 'Aleatorio'
  ]);
  // Asegurar que la hoja de Cursos tenga la columna 'Activo'
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const cursosSheet = ss.getSheetByName('Cursos');
    if (cursosSheet) {
      const headers = cursosSheet.getRange(1, 1, 1, Math.max(cursosSheet.getLastColumn(), 1)).getValues()[0];
      const hasActivo = headers.some(h => String(h || '').trim().toLowerCase() === 'activo');
      if (!hasActivo) {
        const nextCol = Math.max(headers.length, 6) + 1;
        cursosSheet.getRange(1, nextCol).setValue('Activo').setFontWeight('bold');
        const lastRow = cursosSheet.getLastRow();
        if (lastRow > 1) {
          cursosSheet.getRange(2, nextCol, lastRow - 1, 1).setValue('Sí');
        }
      }
    }
  } catch (e) { /* ignora */ }
  return { status: 'success' };
}

// ========================================================
// CURSOS CRUD
// ========================================================

function getCursosAdmin(adminEmail) {
  ensureAdmin_(adminEmail);
  const cursos = getDataAsJson(openSheet_('Cursos'));

  // Conteo de preguntas por curso
  let pregCounts = {};
  try {
    const pregSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Preguntas');
    if (pregSheet) {
      const preguntas = getDataAsJson(pregSheet);
      preguntas.forEach(p => {
        const cid = String(p.CursoID || p.ID || p['#'] || Object.values(p)[0]);
        pregCounts[cid] = (pregCounts[cid] || 0) + 1;
      });
    }
  } catch (e) { /* ignora si no existe */ }

  return cursos.map(c => {
    const cursoId = c.CursoID || c.ID || c['#'] || Object.values(c)[0];
    return {
      ...c,
      CursoID: cursoId,
      numPreguntas: pregCounts[String(cursoId)] || 0
    };
  });
}

function saveCursoAdmin(adminEmail, cursoData) {
  ensureAdmin_(adminEmail);
  const sheet = openSheet_('Cursos');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] ? data[0].map(h => String(h).trim().toLowerCase()) : [];

  // Encontrar índices de columnas dinámicamente (0-based)
  // Fallbacks para los nombres de columnas que tiene el usuario actualmente
  let idxTitulo = headers.indexOf('titulo'); if(idxTitulo===-1) idxTitulo = 1;
  let idxDesc = headers.indexOf('descripcion'); if(idxDesc===-1) idxDesc = 2;
  
  let idxTipo = headers.indexOf('tipo');
  if(idxTipo === -1) { idxTipo = headers.indexOf('categoria'); if(idxTipo!==-1) sheet.getRange(1, idxTipo+1).setValue('Tipo'); else idxTipo = 3; }
  
  let idxEnlace = headers.indexOf('enlacecontenido');
  if(idxEnlace === -1) { idxEnlace = headers.indexOf('imagen'); if(idxEnlace!==-1) sheet.getRange(1, idxEnlace+1).setValue('EnlaceContenido'); else idxEnlace = 4; }
  
  let idxUnidad = headers.indexOf('unidad');
  if(idxUnidad === -1) { idxUnidad = headers.indexOf('creadopor'); if(idxUnidad!==-1) sheet.getRange(1, idxUnidad+1).setValue('Unidad'); else idxUnidad = 5; }
  
  let idxActivo = headers.indexOf('activo');
  if(idxActivo === -1) { idxActivo = 6; sheet.getRange(1, idxActivo+1).setValue('Activo'); }

  if (cursoData.CursoID) {
    // ── UPDATE ──
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(cursoData.CursoID)) {
        sheet.getRange(i + 1, idxTitulo + 1).setValue(cursoData.Titulo || '');
        sheet.getRange(i + 1, idxDesc + 1).setValue(cursoData.Descripcion || '');
        sheet.getRange(i + 1, idxTipo + 1).setValue(cursoData.Tipo || 'video');
        sheet.getRange(i + 1, idxEnlace + 1).setValue(cursoData.EnlaceContenido || '');
        sheet.getRange(i + 1, idxUnidad + 1).setValue(cursoData.Unidad || '');
        sheet.getRange(i + 1, idxActivo + 1).setValue(cursoData.Activo || 'Sí');
        logAudit(adminEmail, 'Admin: Curso actualizado - ' + cursoData.Titulo, 0);
        return { status: 'success', cursoId: String(cursoData.CursoID) };
      }
    }
    return { status: 'error', message: 'Curso no encontrado' };
  } else {
    // ── CREATE ──
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const id = Number(data[i][0]);
      if (id > maxId) maxId = id;
    }
    const newId = maxId + 1;
    
    // Crear un array vacío del tamaño de las cabeceras
    let newRow = new Array(headers.length).fill('');
    newRow[0] = newId;
    newRow[idxTitulo] = cursoData.Titulo || '';
    newRow[idxDesc] = cursoData.Descripcion || '';
    newRow[idxTipo] = cursoData.Tipo || 'video';
    newRow[idxEnlace] = cursoData.EnlaceContenido || '';
    newRow[idxUnidad] = cursoData.Unidad || '';
    newRow[idxActivo] = cursoData.Activo || 'Sí';

    sheet.appendRow(newRow);
    logAudit(adminEmail, 'Admin: Curso creado - ' + cursoData.Titulo, 0);
    return { status: 'success', cursoId: String(newId) };
  }
}

function deleteCursoAdmin(adminEmail, cursoId) {
  ensureAdmin_(adminEmail);
  const sheet = openSheet_('Cursos');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(cursoId)) {
      const titulo = data[i][1];
      sheet.deleteRow(i + 1);
      try { deleteAllPreguntasByCurso_(cursoId); } catch (e) { }
      try { deleteConfigQuiz_(cursoId); } catch (e) { }
      logAudit(adminEmail, 'Admin: Curso eliminado - ' + titulo, 0);
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Curso no encontrado' };
}

// ========================================================
// COMPOUND: Guardar curso completo (curso + config + preguntas)
// ========================================================

function saveCursoCompleto(adminEmail, cursoData, configData, preguntasData) {
  ensureAdmin_(adminEmail);

  // 1. Guardar/crear curso
  const result = saveCursoAdmin(adminEmail, cursoData);
  if (result.status !== 'success') return result;
  const cursoId = result.cursoId;

  // 2. Guardar config quiz
  if (configData) {
    saveConfigQuiz(adminEmail, cursoId, configData);
  }

  // 3. Guardar preguntas (reemplaza todas las existentes)
  if (preguntasData && preguntasData.length > 0) {
    savePreguntasForCurso(adminEmail, cursoId, preguntasData);
  } else {
    try { deleteAllPreguntasByCurso_(cursoId); } catch (e) { }
  }

  return { status: 'success', cursoId: cursoId, message: 'Curso guardado exitosamente' };
}

/**
 * Carga datos completos de un curso para el editor
 * (curso + config quiz + preguntas).
 */
function getCursoCompleto(adminEmail, cursoId) {
  ensureAdmin_(adminEmail);

  const cursos = getDataAsJson(openSheet_('Cursos'));
  const curso = cursos.find(c => {
    const id = c.CursoID || c.ID || c['#'] || Object.values(c)[0];
    return String(id) === String(cursoId);
  });
  if (!curso) return { status: 'error', message: 'Curso no encontrado' };

  let config = null;
  try {
    const configSheet = openOrCreateSheet_('ConfigQuiz', [
      'CursoID', 'TotalPreguntas', 'PuntajeMinimo', 'IntentosMax', 'TiempoLimiteMin', 'Aleatorio'
    ]);
    const configs = getDataAsJson(configSheet);
    config = configs.find(c => String(c.CursoID) === String(cursoId)) || null;
  } catch (e) { }

  let preguntas = [];
  try {
    const pregSheet = openOrCreateSheet_('Preguntas', [
      'PreguntaID', 'CursoID', 'Texto', 'OpcionA', 'OpcionB',
      'OpcionC', 'OpcionD', 'RespuestaCorrecta', 'Puntos'
    ]);
    preguntas = getDataAsJson(pregSheet).filter(p => String(p.CursoID) === String(cursoId));
  } catch (e) { }

  return { status: 'success', curso: curso, config: config, preguntas: preguntas };
}

// ========================================================
// PREGUNTAS CRUD
// ========================================================

function getPreguntasByCurso(adminEmail, cursoId) {
  ensureAdmin_(adminEmail);
  const sheet = openOrCreateSheet_('Preguntas', [
    'PreguntaID', 'CursoID', 'Texto', 'OpcionA', 'OpcionB',
    'OpcionC', 'OpcionD', 'RespuestaCorrecta', 'Puntos'
  ]);
  return getDataAsJson(sheet).filter(p => String(p.CursoID) === String(cursoId));
}

function savePreguntasForCurso(adminEmail, cursoId, preguntas) {
  ensureAdmin_(adminEmail);

  // Borra todas las preguntas existentes del curso
  deleteAllPreguntasByCurso_(cursoId);

  // Re-crea desde el array proporcionado
  const sheet = openOrCreateSheet_('Preguntas', [
    'PreguntaID', 'CursoID', 'Texto', 'OpcionA', 'OpcionB',
    'OpcionC', 'OpcionD', 'RespuestaCorrecta', 'Puntos'
  ]);
  const data = sheet.getDataRange().getValues();

  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = Number(data[i][0]);
    if (id > maxId) maxId = id;
  }

  preguntas.forEach(p => {
    maxId++;
    sheet.appendRow([
      maxId,
      cursoId,
      p.Texto || '',
      p.OpcionA || '',
      p.OpcionB || '',
      p.OpcionC || '',
      p.OpcionD || '',
      p.RespuestaCorrecta || 'A',
      Number(p.Puntos) || 10
    ]);
  });

  return { status: 'success', count: preguntas.length };
}

function deleteAllPreguntasByCurso_(cursoId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Preguntas');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(cursoId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ========================================================
// CONFIG QUIZ
// ========================================================

function getConfigQuiz(adminEmail, cursoId) {
  ensureAdmin_(adminEmail);
  const sheet = openOrCreateSheet_('ConfigQuiz', [
    'CursoID', 'TotalPreguntas', 'PuntajeMinimo', 'IntentosMax', 'TiempoLimiteMin', 'Aleatorio'
  ]);
  const all = getDataAsJson(sheet);
  return all.find(c => String(c.CursoID) === String(cursoId)) || null;
}

function saveConfigQuiz(adminEmail, cursoId, config) {
  ensureAdmin_(adminEmail);
  const sheet = openOrCreateSheet_('ConfigQuiz', [
    'CursoID', 'TotalPreguntas', 'PuntajeMinimo', 'IntentosMax', 'TiempoLimiteMin', 'Aleatorio'
  ]);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(cursoId)) {
      sheet.getRange(i + 1, 2).setValue(Number(config.TotalPreguntas) || 0);
      sheet.getRange(i + 1, 3).setValue(Number(config.PuntajeMinimo) || 70);
      sheet.getRange(i + 1, 4).setValue(Number(config.IntentosMax) || 3);
      sheet.getRange(i + 1, 5).setValue(Number(config.TiempoLimiteMin) || 30);
      sheet.getRange(i + 1, 6).setValue(config.Aleatorio ? 'Sí' : 'No');
      return { status: 'success' };
    }
  }

  // No existía, crear nueva fila
  sheet.appendRow([
    cursoId,
    Number(config.TotalPreguntas) || 0,
    Number(config.PuntajeMinimo) || 70,
    Number(config.IntentosMax) || 3,
    Number(config.TiempoLimiteMin) || 30,
    config.Aleatorio ? 'Sí' : 'No'
  ]);
  return { status: 'success' };
}

function deleteConfigQuiz_(cursoId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('ConfigQuiz');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(cursoId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ========================================================
// USUARIOS CRUD
// ========================================================

function getUsuariosAdmin(adminEmail) {
  ensureAdmin_(adminEmail);
  const data = getDataAsJson(openSheet_('Usuarios'));
  // No retorna contraseñas por seguridad
  return data.map(u => ({
    ID: u.ID,
    Email: u.Email,
    Nombre: u.Nombre,
    Apellido: u.Apellido,
    Rol: u.Rol,
    Foto: u.Foto
  }));
}

function saveUsuarioAdmin(adminEmail, userData) {
  ensureAdmin_(adminEmail);
  const sheet = openSheet_('Usuarios');
  const data = sheet.getDataRange().getValues();

  if (userData.ID) {
    // ── UPDATE ──
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userData.ID)) {
        if (userData.Email) sheet.getRange(i + 1, 2).setValue(normalizeEmail_(userData.Email));
        if (userData.Nombre !== undefined) sheet.getRange(i + 1, 4).setValue(userData.Nombre);
        if (userData.Apellido !== undefined) sheet.getRange(i + 1, 5).setValue(userData.Apellido);
        if (userData.Rol) sheet.getRange(i + 1, 6).setValue(userData.Rol);
        if (userData.Foto !== undefined) sheet.getRange(i + 1, 8).setValue(userData.Foto);
        logAudit(adminEmail, 'Admin: Usuario editado - ' + userData.Email, 0);
        return { status: 'success' };
      }
    }
    return { status: 'error', message: 'Usuario no encontrado' };
  } else {
    // ── CREATE ──
    const emailN = normalizeEmail_(userData.Email);
    if (!emailN) return { status: 'error', message: 'El email es obligatorio' };

    for (let i = 1; i < data.length; i++) {
      if (normalizeEmail_(data[i][1]) === emailN) {
        return { status: 'error', message: 'El email ya existe en el sistema' };
      }
    }

    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const id = Number(data[i][0]);
      if (id > maxId) maxId = id;
    }

    const hashedPass = getSha256Hash_(userData.Password || '123456');
    sheet.appendRow([
      maxId + 1,
      emailN,
      hashedPass,
      userData.Nombre || '',
      userData.Apellido || '',
      userData.Rol || 'colaborador',
      '',
      userData.Foto || ''
    ]);
    logAudit(adminEmail, 'Admin: Usuario creado - ' + emailN, 0);
    return { status: 'success', message: 'Usuario creado. Contraseña temporal: 123456' };
  }
}

function deleteUsuarioAdmin(adminEmail, userId) {
  ensureAdmin_(adminEmail);
  const sheet = openSheet_('Usuarios');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      if (normalizeEmail_(data[i][1]) === normalizeEmail_(adminEmail)) {
        return { status: 'error', message: 'No puede eliminarse a sí mismo' };
      }
      const email = data[i][1];
      sheet.deleteRow(i + 1);
      logAudit(adminEmail, 'Admin: Usuario eliminado - ' + email, 0);
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Usuario no encontrado' };
}

function resetPasswordAdmin(adminEmail, userId) {
  ensureAdmin_(adminEmail);
  const sheet = openSheet_('Usuarios');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      const newPass = getSha256Hash_('123456');
      sheet.getRange(i + 1, 3).setValue(newPass);
      logAudit(adminEmail, 'Admin: Password reseteado - ID ' + userId, 0);
      return { status: 'success', message: 'Contraseña restablecida a: 123456' };
    }
  }
  return { status: 'error', message: 'Usuario no encontrado' };
}

// ========================================================
// QUIZ DINÁMICO (para estudiantes)
// ========================================================

/**
 * Obtiene las preguntas de un curso para el quiz del estudiante.
 * No requiere permisos de admin.
 */
function getQuizQuestions(cursoId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    const pregSheet = ss.getSheetByName('Preguntas');
    if (!pregSheet) return { questions: [], config: null };

    let preguntas = getDataAsJson(pregSheet)
      .filter(p => String(p.CursoID) === String(cursoId));

    let config = null;
    const configSheet = ss.getSheetByName('ConfigQuiz');
    if (configSheet) {
      const configs = getDataAsJson(configSheet);
      config = configs.find(c => String(c.CursoID) === String(cursoId)) || null;
    }

    // Mezclar si está configurado como aleatorio
    if (config && (config.Aleatorio === 'Sí' || config.Aleatorio === true)) {
      preguntas = [...preguntas].sort(() => Math.random() - 0.5);
    }

    // Limitar número de preguntas
    if (config && Number(config.TotalPreguntas) > 0) {
      preguntas = preguntas.slice(0, Number(config.TotalPreguntas));
    }

    return { questions: preguntas, config: config };
  } catch (e) {
    return { questions: [], config: null };
  }
}

// ========================================================
// ADMIN DASHBOARD
// ========================================================

function getAdminDashboardData(adminEmail) {
  ensureAdmin_(adminEmail);

  const usuarios = getDataAsJson(openSheet_('Usuarios'));
  const cursos = getDataAsJson(openSheet_('Cursos'));
  const progreso = getDataAsJson(openSheet_('Progreso'));

  const totalUsuarios = usuarios.length;
  const totalCursos = cursos.length;
  const totalCompletados = progreso.filter(p => {
    const e = safeStr_(p.Estado ?? p['Estado (completado/pendiente)']).toLowerCase();
    return e === 'completado';
  }).length;

  const avgScore = progreso.length > 0
    ? (progreso.reduce((acc, p) => acc + safeNumber_(p.Nota, 0), 0) / progreso.length).toFixed(1)
    : '0.0';

  // Estadísticas por curso
  const cursosStats = cursos.map(c => {
    const prog = progreso.filter(p => String(p.CursoID) === String(c.CursoID));
    const comp = prog.filter(p =>
      safeStr_(p.Estado ?? p['Estado (completado/pendiente)']).toLowerCase() === 'completado'
    ).length;
    return {
      cursoId: c.CursoID,
      titulo: c.Titulo,
      totalInscritos: prog.length,
      completados: comp,
      promedioNota: prog.length > 0
        ? (prog.reduce((a, p) => a + safeNumber_(p.Nota, 0), 0) / prog.length).toFixed(1)
        : '0.0'
    };
  });

  // Actividad reciente
  const auditData = getDataAsJson(openSheet_('AuditoriaLog'));
  const recentActivity = auditData.slice()
    .sort((a, b) => new Date(b.FechaHora) - new Date(a.FechaHora))
    .slice(0, 15);

  return {
    totalUsuarios,
    totalCursos,
    totalCompletados,
    avgScore,
    cursosStats,
    recentActivity
  };
}

// ========================================================
// REPORTES
// ========================================================

function getReporteProgreso(adminEmail) {
  ensureAdmin_(adminEmail);

  const usuarios = getDataAsJson(openSheet_('Usuarios'));
  const cursos = getDataAsJson(openSheet_('Cursos'));
  const progreso = getDataAsJson(openSheet_('Progreso'));

  return progreso.map(p => {
    const user = usuarios.find(u => normalizeEmail_(u.Email) === normalizeEmail_(p.Email));
    const curso = cursos.find(c => String(c.CursoID) === String(p.CursoID));
    return {
      email: p.Email,
      nombreUsuario: user
        ? (safeStr_(user.Nombre) + ' ' + safeStr_(user.Apellido)).trim()
        : p.Email,
      cursoTitulo: curso ? curso.Titulo : ('Curso ID: ' + p.CursoID),
      estado: safeStr_(p.Estado ?? p['Estado (completado/pendiente)']),
      nota: safeNumber_(p.Nota, 0),
      intentos: safeNumber_(p.Intentos, 0)
    };
  });
}
