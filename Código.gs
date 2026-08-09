const SHEET_ID = '1mM90aAVbSocbeUqfJvLc01bHEIx3xOwBamy1oMGjnpw'; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Campus Virtual Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ----------------- HELPERS -----------------
function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function safeNumber_(v, fallback = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function safeStr_(v) {
  return String(v ?? '').trim();
}

function openSheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`La pestaña "${name}" no existe en la hoja de cálculo.`);
  return sheet;
}

function getDataAsJson(sheet) {
  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];

  const headersRaw = data[0].map(h => String(h || '').trim());

  return data
    .slice(1)
    .filter(row => row.some(v => v !== '' && v !== null))
    .map(row => {
      const obj = {};
      headersRaw.forEach((h, i) => {
        if (!h) return;
        obj[h] = row[i];
        const alias = h.split('(')[0].trim();
        if (alias && !(alias in obj)) obj[alias] = row[i];
      });
      return obj;
    });
}

function getSha256Hash_(value) {
  if (!value) return '';
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

// ----------------- LOGIN & USUARIOS -----------------
function loginUser(email, password) {
  try {
    const emailN = normalizeEmail_(email);
    const pass = String(password ?? '');
    const hashedPass = getSha256Hash_(pass).trim().toLowerCase(); 

    const sheet = openSheet_('Usuarios');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const rowEmail = normalizeEmail_(data[i][1]); // Columna B (Email)
      const rowPass = String(data[i][2] ?? '').trim().toLowerCase(); // Columna C (Password)    

      if (rowEmail === emailN && rowPass === hashedPass) {
        const nombre = safeStr_(data[i][3]) + ' ' + safeStr_(data[i][4]); 
        const rol = safeStr_(data[i][5]).toLowerCase();
        const foto = safeStr_(data[i][7]);

        // Guardamos la auditoria intentando que si falla, no bloquee el login
        try { logAudit(emailN, 'Login', 0); } catch (e) { /* ignorar error de auditoria */ }

        return {
          status: 'success',
          user: {
            id: data[i][0] || '', 
            email: rowEmail,
            name: nombre.trim(),
            role: rol,
            avatar: foto || ('https://ui-avatars.com/api/?name=' + encodeURIComponent(nombre.trim() || 'Usuario'))
          }
        };
      }
    }
    return { status: 'error', message: 'Credenciales inválidas' };
    
  } catch (error) {
    // ESTO DEVOLVERÁ EL ERROR EXACTO AL FRONTEND
    return { status: 'error', message: 'Error del servidor: ' + error.message };
  }
}

// ----------------- CURSOS -----------------
function getCursos(userEmail) {
  try {
    const emailN = normalizeEmail_(userEmail);
    let cursos = getDataAsJson(openOrCreateSheet_('Cursos', ['CursoID', 'Titulo', 'Modulo', 'Enlace', 'Activo']));
    
    // Normalizar ID y filtrar inactivos
    cursos = cursos.map(c => ({
      ...c,
      CursoID: c.CursoID || c.ID || c['#'] || Object.values(c)[0]
    })).filter(c => {
      const activo = String(c.Activo || 'Sí').trim().toLowerCase();
      return activo !== 'no' && activo !== 'false' && activo !== '0' && activo !== 'inactivo';
    });

    const progreso = getDataAsJson(openOrCreateSheet_('Progreso', ['Email', 'CursoID', 'Estado', 'Nota', 'Intentos'])).filter(p => normalizeEmail_(p.Email || p.ID || p.UserID || Object.values(p)[0]) === emailN);

    return cursos.map(c => {
      const userProg = progreso.find(p => {
        const pCid = String(p.CursoID || p.ID || p['Curso ID'] || p['#'] || Object.values(p)[1]);
        return pCid === String(c.CursoID);
      });
      const estadoRaw = userProg ? (userProg.Estado || userProg['Estado (completado/pendiente)'] || Object.values(userProg)[2]) : 'pendiente';
      const estado = safeStr_(estadoRaw).toLowerCase();
      const nota = safeNumber_(userProg ? (userProg.Nota || userProg.Score || Object.values(userProg)[3]) : 0, 0);
      const intentos = safeNumber_(userProg ? (userProg.Intentos || userProg.Attempts || Object.values(userProg)[4]) : 0, 0);

      return {
        ...c,
        status: estado || 'pendiente',
        score: nota,
        attempts: intentos
      };
    });
  } catch (error) {
    return []; // Si hay error devuelve vacio para no romper
  }
}

// ----------------- GUARDAR PROGRESO -----------------
function saveQuizResult(email, cursoId, score, passed) {
  try {
    const e = normalizeEmail_(email);
    const cId = String(cursoId);
    const sc = safeNumber_(score, 0);

    const sheet = openOrCreateSheet_('Progreso', ['Email', 'CursoID', 'Estado', 'Nota', 'Intentos']);
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 1) return { status: 'error', message: 'No headers' };
    
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = headers.indexOf('email') !== -1 ? headers.indexOf('email') : (headers.indexOf('id') !== -1 ? headers.indexOf('id') : 0);
    const cursoIdIdx = headers.indexOf('cursoid') !== -1 ? headers.indexOf('cursoid') : 1;
    const estadoIdx = headers.indexOf('estado') !== -1 ? headers.indexOf('estado') : 2;
    const notaIdx = headers.indexOf('nota') !== -1 ? headers.indexOf('nota') : (headers.indexOf('porcentaje') !== -1 ? headers.indexOf('porcentaje') : 3);
    const intentosIdx = headers.indexOf('intentos') !== -1 ? headers.indexOf('intentos') : 4;

    let found = false;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = normalizeEmail_(data[i][emailIdx]);
      const rowCid = String(data[i][cursoIdIdx]);
      if (rowEmail === e && rowCid === cId) {
        const prevAttempts = safeNumber_(data[i][intentosIdx], 0);
        sheet.getRange(i + 1, estadoIdx + 1).setValue(passed ? 'completado' : 'pendiente'); 
        sheet.getRange(i + 1, notaIdx + 1).setValue(sc); 
        sheet.getRange(i + 1, intentosIdx + 1).setValue(prevAttempts + 1); 
        found = true;
        break;
      }
    }
    if (!found) {
      const newRow = new Array(headers.length).fill('');
      newRow[emailIdx] = e;
      newRow[cursoIdIdx] = cId;
      newRow[estadoIdx] = passed ? 'completado' : 'pendiente';
      newRow[notaIdx] = sc;
      newRow[intentosIdx] = 1;
      sheet.appendRow(newRow);
    }
    try { logAudit(e, 'Quiz Realizado: ' + cId + ' (Nota: ' + sc + ')', 5); } catch(e){}
    SpreadsheetApp.flush(); // IMPORTANT: Forzar escritura para que las siguientes lecturas estén actualizadas
    return { status: 'success', score: sc, passed: passed };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// ----------------- AUDITORÍA -----------------
function logAudit(email, action, duration) {
  const sheet = openOrCreateSheet_('AuditoriaLog', ['Email', 'FechaHora', 'Accion', 'DuracionMinutos']);
  const headers = sheet.getDataRange().getValues()[0].map(h => String(h).trim().toLowerCase());
  
  const emailIdx = headers.indexOf('email') !== -1 ? headers.indexOf('email') : (headers.indexOf('id') !== -1 ? headers.indexOf('id') : 0);
  const fechaIdx = headers.indexOf('fechahora') !== -1 ? headers.indexOf('fechahora') : (headers.indexOf('timestamp') !== -1 ? headers.indexOf('timestamp') : 1);
  const accionIdx = headers.indexOf('accion') !== -1 ? headers.indexOf('accion') : 2;
  const duracionIdx = headers.indexOf('duracionminutos') !== -1 ? headers.indexOf('duracionminutos') : (headers.indexOf('duracion') !== -1 ? headers.indexOf('duracion') : 3);
  
  const newRow = new Array(headers.length).fill('');
  newRow[emailIdx] = normalizeEmail_(email);
  newRow[fechaIdx] = new Date();
  newRow[accionIdx] = safeStr_(action);
  newRow[duracionIdx] = safeNumber_(duration, 0);
  
  sheet.appendRow(newRow);
}

// ----------------- DASHBOARD -----------------
function getDashboardMetrics(email, role) {
  try {
    const e = normalizeEmail_(email);
    const r = safeStr_(role).toLowerCase();

    const progData = getDataAsJson(openOrCreateSheet_('Progreso', ['Email', 'CursoID', 'Estado', 'Nota', 'Intentos']));
    const auditData = getDataAsJson(openOrCreateSheet_('AuditoriaLog', ['Email', 'FechaHora', 'Accion', 'DuracionMinutos'])); 
    const cursosDataRaw = getDataAsJson(openOrCreateSheet_('Cursos', ['CursoID', 'Titulo', 'Modulo', 'Enlace', 'Activo']));
    
    // Normalizar ID y filtrar inactivos como en getCursos
    const cursosData = cursosDataRaw.map(c => ({
      ...c,
      CursoID: c.CursoID || c.ID || c['#'] || Object.values(c)[0]
    })).filter(c => {
      const activo = String(c.Activo || 'Sí').trim().toLowerCase();
      return activo !== 'no' && activo !== 'false' && activo !== '0' && activo !== 'inactivo';
    });

    const privileged = (r === 'administrador' || r === 'gerente' || r === 'supervisor');

    const myProgress = privileged ? progData : progData.filter(p => normalizeEmail_(p.Email || p.ID || p.UserID || Object.values(p)[0]) === e);
    const myAudit = privileged ? auditData : auditData.filter(a => normalizeEmail_(a.Email || a.ID || a.UserID || Object.values(a)[0]) === e);

    // Agrupar cursos por Título
    const grouped = {};
    cursosData.forEach(c => {
      const t = String(c.Titulo || c.Título || 'Sin Título').trim();
      if (!grouped[t]) grouped[t] = { modulos: [] };
      grouped[t].modulos.push(c);
    });

    let completed = 0;
    let pending = 0;

    Object.keys(grouped).forEach(t => {
      const modulos = grouped[t].modulos;
      let allCompleted = true;

      modulos.forEach(m => {
        const userProg = myProgress.find(p => {
          const pCid = String(p.CursoID || p.ID || p['Curso ID'] || p['#'] || Object.values(p)[1]);
          return pCid === String(m.CursoID);
        });
        const estadoRaw = userProg ? (userProg.Estado || userProg['Estado (completado/pendiente)'] || Object.values(userProg)[2]) : 'pendiente';
        if (safeStr_(estadoRaw).toLowerCase() !== 'completado') {
          allCompleted = false;
        }
      });

      if (allCompleted && modulos.length > 0) {
        completed++;
      } else {
        pending++; 
      }
    });

    const avgScore = myProgress.length > 0 ? myProgress.reduce((acc, curr) => {
      const nota = curr.Nota || curr.Score || Object.values(curr)[3];
      return acc + safeNumber_(nota, 0);
    }, 0) / myProgress.length : 0;

    const totalTime = myAudit.reduce((acc, curr) => acc + safeNumber_(curr.DuracionMinutos || curr.Duracion, 0), 0);
    const auditLog = myAudit.slice().sort((a, b) => new Date(b.FechaHora || b.Fecha) - new Date(a.FechaHora || a.Fecha)).slice(0, 5);

    return { completedCourses: completed, pendingCourses: pending, totalCourses: Object.keys(grouped).length, averageScore: avgScore.toFixed(1), timeSpent: totalTime, auditLog: auditLog };
  } catch (error) {
    return { completedCourses: 0, pendingCourses: 0, totalCourses: 0, averageScore: "0.0", timeSpent: 0, auditLog: [] };
  }
}

function getAuditData(email, role) {
  try {
    const e = normalizeEmail_(email);
    const r = safeStr_(role).toLowerCase();

    const sheet = openOrCreateSheet_('AuditoriaLog', ['Email', 'FechaHora', 'Accion', 'DuracionMinutos']);
    const data = sheet.getDataRange().getValues();

    const logs = data.slice(1).filter(row => row.some(v => v !== '' && v !== null)).map(row => ({
        Usuario: normalizeEmail_(row[0]), Fecha: row[1], Accion: safeStr_(row[2]), Duracion: safeNumber_(row[3], 0)
    }));

    const privileged = (r === 'administrador' || r === 'gerente' || r === 'supervisor');
    if (privileged) return logs.reverse();
    return logs.filter(l => l.Usuario === e).reverse();
  } catch (error) {
    return [];
  }
}

// ----------------- EVALUACIONES PRESENCIALES -----------------
function guardarEvaluacion(datos) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Evaluaciones_Presenciales');
    
    if (!sheet) {
      sheet = ss.insertSheet('Evaluaciones_Presenciales');
      const headers = ['Timestamp', 'Email Estudiante', 'Curso ID', 'Curso Titulo', 'Instructor', 'Calificación General', 'Dominio del Tema', 'Materiales', 'Aplicabilidad', 'Comentarios'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    }
    
    const rowData = [
      new Date(),
      normalizeEmail_(datos.email),
      safeStr_(datos.cursoId),
      safeStr_(datos.cursoTitulo),
      safeStr_(datos.instructor),
      safeNumber_(datos.rating, 0),
      safeStr_(datos.dominio),
      safeStr_(datos.materiales),
      safeStr_(datos.aplicabilidad),
      safeStr_(datos.comentarios)
    ];
    
    sheet.appendRow(rowData);
    
    try { logAudit(datos.email, 'Evaluación enviada para curso: ' + datos.cursoId, 0); } catch(e){}
    
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
