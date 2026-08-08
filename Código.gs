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
    let cursos = getDataAsJson(openSheet_('Cursos'));
    
    // Normalizar ID y filtrar inactivos
    cursos = cursos.map(c => ({
      ...c,
      CursoID: c.CursoID || c.ID || c['#'] || Object.values(c)[0]
    })).filter(c => {
      const activo = String(c.Activo || 'Sí').trim().toLowerCase();
      return activo !== 'no' && activo !== 'false' && activo !== '0' && activo !== 'inactivo';
    });

    const progreso = getDataAsJson(openSheet_('Progreso')).filter(p => normalizeEmail_(p.Email) === emailN);

    return cursos.map(c => {
      const userProg = progreso.find(p => String(p.CursoID) === String(c.CursoID));
      const estado = safeStr_(userProg ? (userProg.Estado ?? userProg['Estado (completado/pendiente)']) : 'pendiente').toLowerCase();
      return {
        ...c,
        status: estado || 'pendiente',
        score: safeNumber_(userProg ? userProg.Nota : 0, 0),
        attempts: safeNumber_(userProg ? userProg.Intentos : 0, 0)
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

    const sheet = openSheet_('Progreso');
    const data = sheet.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (normalizeEmail_(data[i][0]) === e && String(data[i][1]) === cId) {
        const prevAttempts = safeNumber_(data[i][4], 0);
        const attempts = prevAttempts + 1;
        sheet.getRange(i + 1, 3).setValue(passed ? 'completado' : 'pendiente'); 
        sheet.getRange(i + 1, 4).setValue(sc); 
        sheet.getRange(i + 1, 5).setValue(attempts); 
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([e, cId, passed ? 'completado' : 'pendiente', sc, 1]);
    }
    try { logAudit(e, 'Quiz Realizado: ' + cId, 5); } catch(e){}
    return { status: 'saved' };
  } catch (error) {
    return { status: 'error' };
  }
}

// ----------------- AUDITORÍA -----------------
function logAudit(email, action, duration) {
  const sheet = openSheet_('AuditoriaLog');
  sheet.appendRow([normalizeEmail_(email), new Date(), safeStr_(action), safeNumber_(duration, 0)]);
}

// ----------------- DASHBOARD -----------------
function getDashboardMetrics(email, role) {
  try {
    const e = normalizeEmail_(email);
    const r = safeStr_(role).toLowerCase();

    const progData = getDataAsJson(openSheet_('Progreso'));
    const auditData = getDataAsJson(openSheet_('AuditoriaLog')); 
    const cursosData = getDataAsJson(openSheet_('Cursos'));

    const privileged = (r === 'administrador' || r === 'gerente' || r === 'supervisor');

    const myProgress = privileged ? progData : progData.filter(p => normalizeEmail_(p.Email) === e);
    const myAudit = privileged ? auditData : auditData.filter(a => normalizeEmail_(a.Email) === e);

    const completed = myProgress.filter(p => {
      const estado = safeStr_(p.Estado ?? p['Estado (completado/pendiente)']).toLowerCase();
      return estado === 'completado';
    }).length;

    const pending = Math.max(cursosData.length - completed, 0);
    const avgScore = myProgress.length > 0 ? myProgress.reduce((acc, curr) => acc + safeNumber_(curr.Nota, 0), 0) / myProgress.length : 0;
    const totalTime = myAudit.reduce((acc, curr) => acc + safeNumber_(curr.DuracionMinutos, 0), 0);
    const auditLog = myAudit.slice().sort((a, b) => new Date(b.FechaHora) - new Date(a.FechaHora)).slice(0, 5);

    return { completedCourses: completed, pendingCourses: pending, totalCourses: cursosData.length, averageScore: avgScore.toFixed(1), timeSpent: totalTime, auditLog: auditLog };
  } catch (error) {
    return { completedCourses: 0, pendingCourses: 0, totalCourses: 0, averageScore: "0.0", timeSpent: 0, auditLog: [] };
  }
}

function getAuditData(email, role) {
  try {
    const e = normalizeEmail_(email);
    const r = safeStr_(role).toLowerCase();

    const sheet = openSheet_('AuditoriaLog'); 
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
