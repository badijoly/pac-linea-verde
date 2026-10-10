/**
 * Lógica principal del Dashboard Interactivo de Evaluación Línea Verde - PAC 2026
 * Con evaluación cualitativa rigurosa y normalización canónica de datos
 */

let RAW_DATA = null;
let CURRENT_HITO = 'Hito 4'; // Default a Hito 4
let ACTIVE_TAB = 'resumen';
let CHARTS = {};

// Inicialización al cargar el DOM
document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) {
    lucide.createIcons();
  }
  loadDashboardData();
  setupEventListeners();
});

function setupEventListeners() {
  const reloadBtn = document.getElementById('btn-reload');
  if (reloadBtn) {
    reloadBtn.addEventListener('click', () => {
      loadDashboardData(true);
    });
  }

  const fileInput = document.getElementById('file-input');
  if (fileInput) {
    fileInput.addEventListener('change', handleFileUpload);
  }
}

// Cargar datos desde la API local o data.json de respaldo
async function loadDashboardData(forceReload = false) {
  const syncStatus = document.getElementById('sync-status');
  const syncText = document.getElementById('sync-text');
  syncText.innerText = 'Cargando datos...';
  syncStatus.className = 'flex items-center gap-1.5 text-xs text-slate-600 bg-slate-100 px-3 py-1.5 rounded-lg border border-slate-200';

  try {
    const response = await fetch(`/api/data?t=${Date.now()}`);
    if (!response.ok) throw new Error('API local no respondió');
    RAW_DATA = await response.json();
    syncText.innerText = 'Conectado a Excel (En vivo)';
    syncStatus.className = 'flex items-center gap-1.5 text-xs text-emerald-800 bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200';
  } catch (err) {
    console.warn('Fallo API local, cargando data.json estático...', err);
    try {
      const respStatic = await fetch(`./data.json?t=${Date.now()}`);
      if (!respStatic.ok) throw new Error('data.json no disponible');
      RAW_DATA = await respStatic.json();
      syncText.innerText = 'Datos locales cargados (data.json)';
      syncStatus.className = 'flex items-center gap-1.5 text-xs text-emerald-800 bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200';
    } catch (e2) {
      console.error('Error crítico al cargar datos:', e2);
      syncText.innerText = 'Error al cargar datos.';
      syncStatus.className = 'flex items-center gap-1.5 text-xs text-rose-800 bg-rose-50 px-3 py-1.5 rounded-lg border border-rose-200';
      return;
    }
  }

  ensureGlobalHito(RAW_DATA);
  populateFilterDropdowns();
  updateUI();
}

// Obtiene la lista consolidada de participantes únicos (83 estudiantes)
function getUniqueParticipants(data) {
  if (!data || !data.hitos) return [];
  const h3Part = (data.hitos['Hito 3'] && data.hitos['Hito 3'].participants) || [];
  const h4Part = (data.hitos['Hito 4'] && data.hitos['Hito 4'].participants) || [];

  const h3Map = new Map();
  h3Part.forEach(p => h3Map.set(stripAccents(p.name), p));

  const h4Map = new Map();
  h4Part.forEach(p => h4Map.set(stripAccents(p.name), p));

  const allNames = new Set([...h3Map.keys(), ...h4Map.keys()]);
  const uniqueList = [];

  allNames.forEach(cleanName => {
    const p3 = h3Map.get(cleanName);
    const p4 = h4Map.get(cleanName);
    const baseP = p4 || p3;

    if (p3 && p4) {
      const preTot = p3.pre_score + p4.pre_score;
      const postTot = p3.post_score + p4.post_score;
      const deltaTot = postTot - preTot;
      const statusTot = deltaTot > 0 ? 'Mejoró' : (deltaTot === 0 ? 'Mantuvo' : 'Retrocedió');
      
      uniqueList.push({
        name: baseP.name,
        school: baseP.school,
        gender: baseP.gender,
        role: baseP.role || 'Estudiante',
        hito: 'Hito 3 y Hito 4',
        hitos_completados: ['Hito 3', 'Hito 4'],
        pre_score: parseFloat(((p3.pre_score + p4.pre_score) / 2).toFixed(2)),
        post_score: parseFloat(((p3.post_score + p4.post_score) / 2).toFixed(2)),
        pre_score_raw: preTot,
        post_score_raw: postTot,
        max_score: 10,
        delta: parseFloat(((postTot - preTot) / 2).toFixed(2)),
        delta_raw: deltaTot,
        status: statusTot,
        real_status: (p3.status === 'Mejoró' || p4.status === 'Mejoró') ? 'Avance en Aprendizajes' : (statusTot === 'Mantuvo' ? 'Rendimiento Estable' : 'Requiere Refuerzo'),
        h3: p3,
        h4: p4
      });
    } else if (p3) {
      uniqueList.push({
        name: p3.name,
        school: p3.school,
        gender: p3.gender,
        role: p3.role || 'Estudiante',
        hito: 'Hito 3',
        hitos_completados: ['Hito 3'],
        pre_score: p3.pre_score,
        post_score: p3.post_score,
        pre_score_raw: p3.pre_score,
        post_score_raw: p3.post_score,
        max_score: 5,
        delta: p3.delta,
        delta_raw: p3.delta,
        status: p3.status,
        real_status: p3.real_status,
        h3: p3,
        h4: null
      });
    } else {
      uniqueList.push({
        name: p4.name,
        school: p4.school,
        gender: p4.gender,
        role: p4.role || 'Estudiante',
        hito: 'Hito 4',
        hitos_completados: ['Hito 4'],
        pre_score: p4.pre_score,
        post_score: p4.post_score,
        pre_score_raw: p4.pre_score,
        post_score_raw: p4.post_score,
        max_score: 5,
        delta: p4.delta,
        delta_raw: p4.delta,
        status: p4.status,
        real_status: p4.real_status,
        h3: null,
        h4: p4
      });
    }
  });

  uniqueList.sort((a, b) => a.name.localeCompare(b.name));
  return uniqueList;
}

// Búsqueda exhaustiva de las evaluaciones de un participante en ambos hitos
function getParticipantAllData(name) {
  if (!name) return { h3: null, h4: null };
  const target = stripAccents(name.trim());
  let h3 = null;
  let h4 = null;

  if (RAW_DATA && RAW_DATA.hitos) {
    if (RAW_DATA.hitos['Hito 3'] && RAW_DATA.hitos['Hito 3'].participants) {
      h3 = RAW_DATA.hitos['Hito 3'].participants.find(p => stripAccents(p.name) === target) || null;
    }
    if (RAW_DATA.hitos['Hito 4'] && RAW_DATA.hitos['Hito 4'].participants) {
      h4 = RAW_DATA.hitos['Hito 4'].participants.find(p => stripAccents(p.name) === target) || null;
    }
  }

  if ((!h3 || !h4) && RAW_DATA && RAW_DATA.all_participants) {
    const ap = RAW_DATA.all_participants.find(p => stripAccents(p.name) === target);
    if (ap && ap.hitos) {
      if (!h3 && ap.hitos['Hito 3']) h3 = ap.hitos['Hito 3'];
      if (!h4 && ap.hitos['Hito 4']) h4 = ap.hitos['Hito 4'];
    }
  }

  return { h3, h4 };
}

// Garantiza la existencia y consistencia del Resumen Global consolidado
function ensureGlobalHito(data) {
  if (!data || !data.hitos) return;

  const uniqueList = getUniqueParticipants(data);

  if (data.hitos['Global']) {
    data.hitos['Global'].unique_participants = uniqueList;
    data.hitos['Global'].participants = uniqueList;
    return;
  }

  const allPaired = [];
  ['Hito 3', 'Hito 4'].forEach(hname => {
    if (data.hitos[hname] && data.hitos[hname].participants) {
      data.hitos[hname].participants.forEach(p => {
        const pCopy = Object.assign({}, p, { hito: p.hito || hname });
        allPaired.push(pCopy);
      });
    }
  });

  const N = allPaired.length;
  if (N === 0) return;

  const avgPre = allPaired.reduce((acc, p) => acc + p.pre_score, 0) / N;
  const avgPost = allPaired.reduce((acc, p) => acc + p.post_score, 0) / N;
  const avgDelta = allPaired.reduce((acc, p) => acc + p.delta, 0) / N;

  const statusCounts = { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0 };
  allPaired.forEach(p => { statusCounts[p.status] = (statusCounts[p.status] || 0) + 1; });

  const gList = allPaired.filter(p => p.pre_score < 5).map(p => (p.post_score - p.pre_score) / (5.0 - p.pre_score));
  const hakeGain = gList.length ? gList.reduce((a, b) => a + b, 0) / gList.length : 0;

  // Colegios consolidados
  const bySchool = {};
  allPaired.forEach(p => {
    if (!bySchool[p.school]) bySchool[p.school] = [];
    bySchool[p.school].push(p);
  });

  const schoolsRankings = [];
  for (const [sname, plist] of Object.entries(bySchool)) {
    const sn = plist.length;
    const spre = plist.reduce((a, p) => a + p.pre_score, 0) / sn;
    const spost = plist.reduce((a, p) => a + p.post_score, 0) / sn;
    const sdel = plist.reduce((a, p) => a + p.delta, 0) / sn;
    const sst = { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0 };
    plist.forEach(p => { sst[p.status] = (sst[p.status] || 0) + 1; });
    const meta = (data.schools_directory && data.schools_directory[sname]) || {};
    schoolsRankings.push({
      school: sname,
      n: sn,
      pre_avg: parseFloat(spre.toFixed(2)),
      post_avg: parseFloat(spost.toFixed(2)),
      delta_avg: parseFloat(sdel.toFixed(2)),
      status_counts: sst,
      pct_mejora: parseFloat(((sst['Mejoró'] / sn) * 100).toFixed(1)),
      lider1: meta.lider1 || '',
      email1: meta.email1 || '',
      tel1: meta.tel1 || '',
      lider2: meta.lider2 || '',
      email2: meta.email2 || '',
      tel2: meta.tel2 || '',
      dependencia: meta.dependencia || 'Municipal',
      certificacion: meta.certificacion || 'No',
      has_data: true
    });
  }
  schoolsRankings.sort((a, b) => b.delta_avg - a.delta_avg);

  // Género consolidado
  const byGender = {};
  allPaired.forEach(p => {
    if (!byGender[p.gender]) byGender[p.gender] = [];
    byGender[p.gender].push(p);
  });
  const genderMetrics = {};
  for (const [gname, plist] of Object.entries(byGender)) {
    const gn = plist.length;
    const gpre = plist.reduce((a, p) => a + p.pre_score, 0) / gn;
    const gpost = plist.reduce((a, p) => a + p.post_score, 0) / gn;
    const gdel = plist.reduce((a, p) => a + p.delta, 0) / gn;
    const gst = { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0 };
    plist.forEach(p => { gst[p.status] = (gst[p.status] || 0) + 1; });
    genderMetrics[gname] = {
      n: gn,
      pre_avg: parseFloat(gpre.toFixed(2)),
      post_avg: parseFloat(gpost.toFixed(2)),
      delta_avg: parseFloat(gdel.toFixed(2)),
      status_counts: gst,
      pct_mejora: parseFloat(((gst['Mejoró'] / gn) * 100).toFixed(1))
    };
  }

  const globQuestions = [];
  if (data.hitos['Hito 3']) {
    data.hitos['Hito 3'].questions.forEach(q => {
      globQuestions.push(Object.assign({}, q, { hito: 'Hito 3', title: `[Hito 3] ${q.title}` }));
    });
  }
  if (data.hitos['Hito 4']) {
    data.hitos['Hito 4'].questions.forEach(q => {
      globQuestions.push(Object.assign({}, q, { hito: 'Hito 4', title: `[Hito 4] ${q.title}` }));
    });
  }

  const globOpen = [];
  if (data.hitos['Hito 3']) {
    data.hitos['Hito 3'].open_questions.forEach(t => globOpen.push(`[Hito 3] ${t}`));
  }
  if (data.hitos['Hito 4']) {
    data.hitos['Hito 4'].open_questions.forEach(t => globOpen.push(`[Hito 4] ${t}`));
  }

  data.hitos['Global'] = {
    desafio: 'Resumen Global Línea Verde (Hito 3 y Hito 4)',
    n_participants: uniqueList.length,
    all_evaluations_count: N,
    avg_pre: parseFloat(avgPre.toFixed(2)),
    avg_post: parseFloat(avgPost.toFixed(2)),
    avg_delta: parseFloat(avgDelta.toFixed(2)),
    hake_gain: parseFloat(hakeGain.toFixed(3)),
    status_counts: statusCounts,
    questions: globQuestions,
    open_questions: globOpen,
    schools: schoolsRankings,
    gender: genderMetrics,
    unique_participants: uniqueList,
    participants: uniqueList,
    all_evaluations: allPaired
  };
}

// Manejo de carga de archivo Excel directamente en el navegador con SheetJS
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      processWorkbookInBrowser(workbook, file.name);
    } catch (err) {
      alert('Error al leer el archivo Excel: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Helpers de normalización canónica
function stripAccents(s) {
  if (!s) return '';
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function canonicalSchoolName(name) {
  if (!name) return '';
  const n = String(name).trim().replace(/"/g, '').replace(/'/g, '');
  const nc = stripAccents(n);
  if (nc.includes('heroes')) return 'Escuela Héroes de la Concepción';
  if (nc.includes('sabella')) return 'Liceo Bicentenario Andrés Sabella';
  if (nc.includes('bet-el') || nc.includes('bet el') || nc.includes('betel')) return 'Colegio Bet-el';
  if (nc.includes('bandera')) return 'Escuela La Bandera';
  if (nc.includes('presbiteriana')) return 'Escuela Presbiteriana';
  if (nc.includes('romulo')) return 'Escuela Rómulo Peña';
  if (nc.includes('greenhill')) return 'Greenhill School';
  if (nc.includes('claudio matte')) return 'Escuela Claudio Matte Pérez';
  if (nc.includes('humberto gonzalez')) return 'Escuela Ecológica Humberto González Echegoyen';
  if (nc.includes('alberto hurtado')) return 'Escuela Ecológica Padre Alberto Hurtado';
  if (nc.includes('republica de italia')) return 'Escuela República de Italia';
  if (nc.includes('domingo herrera')) return 'Liceo Domingo Herrera Rivera';
  if (nc.includes('estados unidos')) return 'Escuela República de Estados Unidos';
  if (nc.includes('juan pablo')) return 'Escuela Juan Pablo II';
  if (nc.includes('javiera carrera')) return 'Escuela Javiera Carrera';
  return n;
}

const NON_INFORMATIVE = [
  'no se', 'nada', 'ninguna', 'no me acuerdo', 'nose', 'no responde', 
  'no se como decirlo', 'no se como puedo decirlo', 'no se explicarlo', 
  'no se que poner', 'no sabria decir', 'no responder', 'ninguno', 'no lo se'
];

function isNonInformative(text) {
  if (!text) return True;
  const t = stripAccents(text).replace(/[.,;:_\-]/g, '').trim();
  if (t.length < 2) return true;
  for (const b of NON_INFORMATIVE) {
    if (t === b || t.startsWith(b + ' ') || t.endsWith(' ' + b)) {
      return true;
    }
  }
  return false;
}

function evaluateConceptualLevel(text, hito) {
  if (isNonInformative(text)) return [0, 'No sabe / Sin respuesta'];
  const tc = stripAccents(text);
  const kwH3 = ['ecosistema', 'biodiversidad', 'gaviotin', '50 gr', '50 gramos', 'microorganismo', 'flora', 'fauna', 'salina', 'desierto costero', 'cadena trofica', 'humedal', 'reserva', 'portada', 'endemica', 'migratoria', 'piqueros', 'piquero', 'pelicano'];
  const kwH4 = ['atmosfera', 'escudo', 'capa de ozono', 'gases', 'efecto invernadero', 'radiacion', 'uv', 'contaminacion luminica', 'astronomia', 'telescopio', 'paranal', 'luz azul', 'cielos limpios', 'patrimonio', 'postes de luz mirando hacia arriba', 'postes'];
  const kw = hito === 'Hito 3' ? kwH3 : kwH4;
  const matches = kw.filter(k => tc.includes(k)).length;
  const words = tc.split(/\s+/).length;

  if (matches >= 1 && (words >= 4 || tc.includes('gaviotin') || tc.includes('50') || tc.includes('ecosistema') || tc.includes('escudo') || tc.includes('postes'))) {
    return [3, 'Apropiación Científica'];
  } else if (matches >= 1 || words >= 3) {
    return [2, 'Comprensión Intermedia'];
  } else {
    return [1, 'Opinión Básica'];
  }
}

function compareOpenAnswers(preText, postText, hito) {
  const preNon = isNonInformative(preText);
  const postNon = isNonInformative(postText);

  // 1. Ambos sin información
  if (preNon && postNon) {
    return {
      status: 'Sin Evidencia',
      detail: 'Sin respuesta en PRE ni POST (Requiere revisión con profesor/a líder)',
      requires_review: true,
      pre_level: 0,
      post_level: 0
    };
  }

  // 2. Pre no sabía y post sí tiene respuesta real
  if (preNon && !postNon) {
    const [lvl, lbl] = evaluateConceptualLevel(postText, hito);
    return {
      status: 'Mejoró',
      detail: `Avance conceptual real: Pasa de no responder a expresar ${lbl}`,
      requires_review: false,
      pre_level: 0,
      post_level: lvl
    };
  }

  // 3. Pre tenía respuesta y en post dejó de responder
  if (!preNon && postNon) {
    const [lvlPre, lblPre] = evaluateConceptualLevel(preText, hito);
    return {
      status: 'Retrocedió',
      detail: `Dejó de contestar en el POST (Basal era ${lblPre})`,
      requires_review: true,
      pre_level: lvlPre,
      post_level: 0
    };
  }

  // 4. Ambos informados
  const [lvlPre, lblPre] = evaluateConceptualLevel(preText, hito);
  const [lvlPost, lblPost] = evaluateConceptualLevel(postText, hito);

  if (lvlPost > lvlPre) {
    return {
      status: 'Mejoró',
      detail: `Mayor precisión y vocabulario (${lblPre} ➔ ${lblPost})`,
      requires_review: false,
      pre_level: lvlPre,
      post_level: lvlPost
    };
  } else if (lvlPost === lvlPre) {
    return {
      status: 'Mantuvo',
      detail: `Mantuvo respuesta informada (${lblPost})`,
      requires_review: false,
      pre_level: lvlPre,
      post_level: lvlPost
    };
  } else {
    return {
      status: 'Retrocedió',
      detail: `Menor nivel de detalle (${lblPre} ➔ ${lblPost})`,
      requires_review: false,
      pre_level: lvlPre,
      post_level: lvlPost
    };
  }
}

// Procesar libro Excel en navegador con SheetJS (100% autónomo sin backend)
function processWorkbookInBrowser(wb, fileName) {
  const syncText = document.getElementById('sync-text');
  const syncStatus = document.getElementById('sync-status');
  
  try {
    const keys = {
      'Hito 3': ['d', 'a', 'b', 'c', 'c'],
      'Hito 4': ['b', 'c', 'a', 'b', 'b']
    };

    const cleanStr = (s) => (s === null || s === undefined) ? '' : String(s).trim();

    const femaleNames = new Set([
      'amanda', 'danna', 'isabella', 'paulina', 'sophia', 'xochtiel', 'abigail', 'ayleen', 
      'estefania', 'eysa', 'fabiola', 'gladys', 'mariana', 'matilde', 'matilda', 'monserrat',
      'pamela', 'valentina', 'dafne', 'lia', 'leah', 'julieta', 'kamila', 'janely', 'isidora',
      'salome', 'sara', 'daniela', 'fabrianny', 'fernanda', 'susy', 'monica', 'carol', 'pia',
      'guisselle', 'jocelyn', 'rosmery', 'javiera', 'mariel', 'kathleen', 'martina', 'ignacia',
      'constanza', 'catalina', 'camila', 'antonia', 'francisca', 'florencia', 'lucia', 'valeria',
      'sofia', 'emilia', 'trinidad', 'maite', 'mia', 'zoe', 'antonella', 'dominic', 'breymarw',
      'gioberlys', 'nahuel'
    ]);

    const inferGender = (name) => {
      const first = stripAccents(name).split(' ')[0] || '';
      if (femaleNames.has(first)) return 'Femenino';
      if (first.endsWith('a')) return 'Femenino';
      return 'Masculino';
    };

    // 1. Extraer Directorio de Establecimientos
    const schools = {};
    if (wb.SheetNames.includes('Participantes 2026')) {
      const sheet = wb.Sheets['Participantes 2026'];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[1]) continue;
        const sch = canonicalSchoolName(row[1]);
        if (!schools[sch]) {
          schools[sch] = {
            nombre: sch,
            equipo: row[0],
            direccion: cleanStr(row[2]),
            dependencia: cleanStr(row[3]) || 'Municipal',
            certificacion: cleanStr(row[4]) || 'No',
            participacion_previa: cleanStr(row[5]) || 'No',
            lider1: cleanStr(row[7]),
            email1: cleanStr(row[8]),
            tel1: cleanStr(row[9]),
            lider2: cleanStr(row[10]),
            email2: cleanStr(row[11]),
            tel2: cleanStr(row[12])
          };
        }
      }
    }

    if (wb.SheetNames.includes('Resumen Líderes')) {
      const sheet = wb.Sheets['Resumen Líderes'];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[1]) continue;
        const sch = canonicalSchoolName(row[1]);
        if (!schools[sch]) {
          schools[sch] = {
            nombre: sch,
            dependencia: 'Municipal',
            certificacion: 'No',
            lider1: '', email1: '', tel1: '',
            lider2: '', email2: '', tel2: ''
          };
        }
        if (row[2]) schools[sch].lider1 = cleanStr(row[2]);
        if (row[3]) schools[sch].tel1 = cleanStr(row[3]);
        if (row[4]) schools[sch].email1 = cleanStr(row[4]);
        if (row[5]) schools[sch].lider2 = cleanStr(row[5]);
        if (row[6]) schools[sch].tel2 = cleanStr(row[6]);
        if (row[7]) schools[sch].email2 = cleanStr(row[7]);
      }
    }

    // 2. Extraer Hitos
    const hitosData = {};
    const allParticipants = {};

    const configs = [
      ['Hito 3', 'Encuestas Hito 3_Verde', 'Los otros vecinos de Antofagasta'],
      ['Hito 4', 'Encuestas Hito 4_Verde', 'Proteger nuestro escudo invisible']
    ];

    for (const [hitoName, sheetName, desafioName] of configs) {
      if (!wb.SheetNames.includes(sheetName)) continue;
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (rows.length < 2) continue;

      const header = rows[0];
      const qTitles = [cleanStr(header[4]), cleanStr(header[5]), cleanStr(header[6]), cleanStr(header[7]), cleanStr(header[8])];
      const openTitles = [cleanStr(header[9]), cleanStr(header[10]), cleanStr(header[11]), cleanStr(header[12])];
      const key = keys[hitoName];

      const pDict = {};
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[0]) continue;
        const name = cleanStr(row[0]);
        const role = cleanStr(row[1]) || 'Estudiante';
        const school = canonicalSchoolName(cleanStr(row[2]));
        const pp = cleanStr(row[3]).toUpperCase();
        if (pp !== 'PRE' && pp !== 'POST') continue;

        const mcAns = [cleanStr(row[4]).toLowerCase(), cleanStr(row[5]).toLowerCase(), cleanStr(row[6]).toLowerCase(), cleanStr(row[7]).toLowerCase(), cleanStr(row[8]).toLowerCase()];
        const openAns = [cleanStr(row[9]), cleanStr(row[10]), cleanStr(row[11]), cleanStr(row[12])];
        const correct = mcAns.map((ans, idx) => ans === key[idx] ? 1 : 0);
        const score = correct.reduce((a, b) => a + b, 0);

        if (!pDict[name]) pDict[name] = { name, role, school };
        pDict[name][pp] = { mc: mcAns, correct, score, open: openAns };
      }

      const pairedList = [];
      for (const [name, d] of Object.entries(pDict)) {
        const gender = inferGender(name);
        const sch = d.school || 'Sin Asignar';

        if (!allParticipants[name]) {
          allParticipants[name] = { name, role: d.role, school: sch, gender, hitos: {} };
        }

        if (d.PRE && d.POST) {
          const preScore = d.PRE.score;
          const postScore = d.POST.score;
          const delta = postScore - preScore;
          const mcStatus = delta > 0 ? 'Mejoró' : (delta === 0 ? 'Mantuvo' : 'Retrocedió');

          const qEvo = [];
          for (let i = 0; i < 5; i++) {
            const cp = d.PRE.correct[i];
            const cpo = d.POST.correct[i];
            if (cp === 0 && cpo === 1) qEvo.push('Mejoró');
            else if (cp === 1 && cpo === 1) qEvo.push('Mantuvo Correcto');
            else if (cp === 0 && cpo === 0) qEvo.push('Mantuvo Incorrecto');
            else qEvo.push('Retrocedió');
          }

          const openAnalysis = [];
          const qualStats = { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0, 'Sin Evidencia': 0 };

          for (let i = 0; i < 4; i++) {
            const tPre = d.PRE.open[i];
            const tPost = d.POST.open[i];
            const res = compareOpenAnswers(tPre, tPost, hitoName);
            qualStats[res.status] += 1;

            openAnalysis.push({
              question: openTitles[i],
              pre_text: tPre,
              post_text: tPost,
              status: res.status,
              detail: res.detail,
              requires_review: res.requires_review
            });
          }

          // Estado Real Integrado
          let realStatus = 'Rendimiento Estable';
          if (qualStats['Sin Evidencia'] === 4) {
            realStatus = 'Sin evidencia suficiente (Requiere revisión con líder)';
          } else if (qualStats['Mejoró'] > 0 || qualStats['Mantuvo'] >= 2) {
            if (mcStatus === 'Mejoró') realStatus = 'Avance Integral Demostrado';
            else if (mcStatus === 'Mantuvo') realStatus = 'Aprendizaje Consolidado';
            else realStatus = 'Avance Conceptual Cualitativo (con ajuste en alternativas)';
          } else if (mcStatus === 'Retrocedió' && qualStats['Retrocedió'] > 0) {
            realStatus = 'Dificultad Conceptual (Requiere acompañamiento)';
          } else if (mcStatus === 'Mejoró') {
            realStatus = 'Mejora en Alternativas';
          }

          const pObj = {
            name,
            role: d.role,
            school: sch,
            gender,
            pre_score: preScore,
            post_score: postScore,
            delta,
            status: mcStatus,
            real_status: realStatus,
            qual_stats: qualStats,
            q_evo: qEvo,
            pre_mc: d.PRE.mc,
            post_mc: d.POST.mc,
            open_analysis: openAnalysis
          };
          pairedList.push(pObj);
          allParticipants[name].hitos[hitoName] = pObj;
        }
      }

      const N = pairedList.length;
      const avgPre = N ? pairedList.reduce((a, p) => a + p.pre_score, 0) / N : 0;
      const avgPost = N ? pairedList.reduce((a, p) => a + p.post_score, 0) / N : 0;
      const avgDelta = N ? pairedList.reduce((a, p) => a + p.delta, 0) / N : 0;

      const gList = pairedList.filter(p => p.pre_score < 5).map(p => (p.post_score - p.pre_score) / (5 - p.pre_score));
      const hakeGain = gList.length ? gList.reduce((a, b) => a + b, 0) / gList.length : 0;

      const statusCounts = { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0 };
      pairedList.forEach(p => { statusCounts[p.status] = (statusCounts[p.status] || 0) + 1; });

      const questionsMetrics = [];
      for (let i = 0; i < 5; i++) {
        const preC = pairedList.filter(p => p.pre_mc[i] === key[i]).length;
        const postC = pairedList.filter(p => p.post_mc[i] === key[i]).length;
        const trans = {};
        const postDist = {};
        pairedList.forEach(p => {
          trans[p.q_evo[i]] = (trans[p.q_evo[i]] || 0) + 1;
          if (p.post_mc[i] !== key[i]) {
            const opt = p.post_mc[i].toUpperCase() || 'N/A';
            postDist[opt] = (postDist[opt] || 0) + 1;
          }
        });
        questionsMetrics.push({
          num: i + 1,
          title: qTitles[i],
          correct_key: key[i].toUpperCase(),
          pre_correct: preC,
          pre_pct: N ? parseFloat(((preC / N) * 100).toFixed(1)) : 0,
          post_correct: postC,
          post_pct: N ? parseFloat(((postC / N) * 100).toFixed(1)) : 0,
          diff: postC - preC,
          transitions: trans,
          post_distractors: postDist
        });
      }

      // Escuelas agrupadas
      const bySchool = {};
      pairedList.forEach(p => {
        if (!bySchool[p.school]) bySchool[p.school] = [];
        bySchool[p.school].push(p);
      });

      const schoolRankings = Object.entries(bySchool).map(([sname, plist]) => {
        const sn = plist.length;
        const spre = plist.reduce((a, p) => a + p.pre_score, 0) / sn;
        const spost = plist.reduce((a, p) => a + p.post_score, 0) / sn;
        const sdel = plist.reduce((a, p) => a + p.delta, 0) / sn;
        const sst = { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0 };
        plist.forEach(p => { sst[p.status] = (sst[p.status] || 0) + 1; });
        const meta = schools[sname] || {};
        return {
          school: sname,
          n: sn,
          pre_avg: parseFloat(spre.toFixed(2)),
          post_avg: parseFloat(spost.toFixed(2)),
          delta_avg: parseFloat(sdel.toFixed(2)),
          status_counts: sst,
          pct_mejora: parseFloat(((sst['Mejoró'] / sn) * 100).toFixed(1)),
          lider1: meta.lider1 || '',
          email1: meta.email1 || '',
          tel1: meta.tel1 || '',
          lider2: meta.lider2 || '',
          email2: meta.email2 || '',
          tel2: meta.tel2 || '',
          dependencia: meta.dependencia || 'Municipal',
          certificacion: meta.certificacion || 'No',
          has_data: true
        };
      });

      // Incluir Colegio Bet-El explícitamente si no está en este hito
      if (!bySchool['Colegio Bet-el']) {
        const bMeta = schools['Colegio Bet-el'] || {};
        schoolRankings.push({
          school: 'Colegio Bet-el',
          n: 0,
          pre_avg: null,
          post_avg: null,
          delta_avg: 0.0,
          status_counts: { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0 },
          pct_mejora: 0.0,
          lider1: bMeta.lider1 || 'Pamela Pizarro Juica',
          email1: bMeta.email1 || 'pizarropamela2015@gmail.com',
          tel1: bMeta.tel1 || '56995779316',
          lider2: bMeta.lider2 || '',
          email2: bMeta.email2 || '',
          tel2: bMeta.tel2 || '',
          dependencia: 'Particular Subvencionado',
          certificacion: 'No',
          has_data: false,
          note: 'Sin datos registrados en este hito'
        });
      }

      schoolRankings.sort((a, b) => {
        if (a.has_data === false) return 1;
        if (b.has_data === false) return -1;
        return b.delta_avg - a.delta_avg;
      });

      // Género agrupado
      const byGender = {};
      pairedList.forEach(p => {
        if (!byGender[p.gender]) byGender[p.gender] = [];
        byGender[p.gender].push(p);
      });
      const genderMetrics = {};
      for (const [gname, plist] of Object.entries(byGender)) {
        const gn = plist.length;
        const gpre = plist.reduce((a, p) => a + p.pre_score, 0) / gn;
        const gpost = plist.reduce((a, p) => a + p.post_score, 0) / gn;
        const gdel = plist.reduce((a, p) => a + p.delta, 0) / gn;
        const gst = { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0 };
        plist.forEach(p => { gst[p.status] = (gst[p.status] || 0) + 1; });
        genderMetrics[gname] = {
          n: gn,
          pre_avg: parseFloat(gpre.toFixed(2)),
          post_avg: parseFloat(gpost.toFixed(2)),
          delta_avg: parseFloat(gdel.toFixed(2)),
          status_counts: gst,
          pct_mejora: parseFloat(((gst['Mejoró'] / gn) * 100).toFixed(1))
        };
      }

      hitosData[hitoName] = {
        desafio: desafioName,
        n_participants: N,
        avg_pre: parseFloat(avgPre.toFixed(2)),
        avg_post: parseFloat(avgPost.toFixed(2)),
        avg_delta: parseFloat(avgDelta.toFixed(2)),
        hake_gain: parseFloat(hakeGain.toFixed(3)),
        status_counts: statusCounts,
        questions: questionsMetrics,
        open_questions: openTitles,
        schools: schoolRankings,
        gender: genderMetrics,
        participants: pairedList
      };
    }

    RAW_DATA = {
      metadata: {
        programa: 'Programa de Acción Climática (PAC) 2026',
        linea: 'Línea Verde',
        archivo_fuente: fileName,
        total_participantes: Object.keys(allParticipants).length
      },
      schools_directory: schools,
      hitos: hitosData,
      all_participants: Object.values(allParticipants)
    };

    ensureGlobalHito(RAW_DATA);
    syncText.innerText = `Cálculo autónomo: ${fileName}`;
    syncStatus.className = 'flex items-center gap-1.5 text-xs text-emerald-800 bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200';

    populateFilterDropdowns();
    updateUI();
  } catch (err) {
    console.error('Error parseando Excel en navegador:', err);
    alert('Error al procesar el archivo Excel en el navegador: ' + err.message);
  }
}

// Llenar selector de colegios
function populateFilterDropdowns() {
  const schoolSelect = document.getElementById('filter-school');
  const currentVal = schoolSelect.value;
  schoolSelect.innerHTML = '<option value="ALL">Todos los Establecimientos</option>';

  const schoolsSet = new Set();
  if (RAW_DATA && RAW_DATA.hitos) {
    for (const h of Object.values(RAW_DATA.hitos)) {
      h.schools.forEach(s => {
        if (s.has_data !== false) schoolsSet.add(s.school);
      });
    }
  }

  Array.from(schoolsSet).sort().forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    schoolSelect.appendChild(opt);
  });

  if (schoolsSet.has(currentVal)) {
    schoolSelect.value = currentVal;
  }
}

// Cambiar pestaña activa
function switchTab(tabId) {
  ACTIVE_TAB = tabId;
  document.querySelectorAll('.tab-view').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.nav-tab').forEach(el => {
    el.className = 'nav-tab flex items-center gap-2 px-4 py-2 rounded-full text-xs font-medium text-slate-300 hover:text-white hover:bg-white/10 whitespace-nowrap transition';
  });

  const targetView = document.getElementById(`view-${tabId}`);
  const targetBtn = document.getElementById(`tab-btn-${tabId}`);
  if (targetView) targetView.classList.remove('hidden');
  if (targetBtn) {
    targetBtn.className = 'nav-tab active-tab flex items-center gap-2 px-4 py-2 rounded-full text-xs font-bold bg-[#00d09c] text-slate-950 shadow-md shadow-teal-500/20 whitespace-nowrap transition';
  }

  updateUI();
  if (window.lucide) lucide.createIcons();
}

// Cambiar Hito seleccionado
function setHito(hito) {
  CURRENT_HITO = hito;

  const btn4 = document.getElementById('btn-hito-4');
  const btn3 = document.getElementById('btn-hito-3');
  const btnGlobal = document.getElementById('btn-hito-global') || document.getElementById('btn-hito-ambos');

  [btn4, btn3, btnGlobal].forEach(btn => {
    if (btn) btn.className = 'px-4 py-1.5 rounded-full text-xs font-medium text-slate-300 hover:text-white transition';
  });

  const activeBtn = hito === 'Hito 4' ? btn4 : (hito === 'Hito 3' ? btn3 : btnGlobal);
  if (activeBtn) {
    activeBtn.className = 'px-4 py-1.5 rounded-full text-xs font-bold bg-[#00d09c] text-slate-950 shadow-md shadow-teal-500/20 transition';
  }

  updateUI();
}

// Aplicar filtros de búsqueda
function applyFilters() {
  updateUI();
}

// Obtener participantes filtrados según controles actuales
function getFilteredParticipants() {
  if (!RAW_DATA || !RAW_DATA.hitos) return [];

  const hitoKey = (CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global') ? 'Global' : CURRENT_HITO;
  const hitoData = RAW_DATA.hitos[hitoKey] || RAW_DATA.hitos['Hito 4'];
  if (!hitoData) return [];

  let list = hitoData.participants;

  const schoolFilter = document.getElementById('filter-school').value;
  const genderFilter = document.getElementById('filter-gender').value;
  const statusFilter = document.getElementById('filter-status').value;
  const searchFilter = document.getElementById('filter-search').value.toLowerCase().trim();

  if (schoolFilter !== 'ALL') {
    list = list.filter(p => p.school === schoolFilter);
  }

  if (genderFilter !== 'ALL') {
    list = list.filter(p => p.gender === genderFilter);
  }

  if (statusFilter !== 'ALL') {
    list = list.filter(p => p.status === statusFilter);
  }

  if (searchFilter) {
    list = list.filter(p => p.name.toLowerCase().includes(searchFilter));
  }

  return list;
}

// Actualizar toda la interfaz
function updateUI() {
  if (!RAW_DATA) return;

  updateBanner();
  updateKPIs();
  renderCharts();
  renderSchoolsTab();
  renderQuestionsTab();
  renderQualitativeTab();
  renderPedagogicalTab();
  renderParticipantsTab();

  if (window.lucide) lucide.createIcons();
}

function updateBanner() {
  const titleEl = document.getElementById('banner-title');
  const descEl = document.getElementById('banner-desc');
  const hakeEl = document.getElementById('kpi-hake-val');

  if (CURRENT_HITO === 'Hito 4') {
    titleEl.innerText = 'Proteger nuestro escudo invisible (Hito 4)';
    descEl.innerText = 'Evaluación del aprendizaje sobre la atmósfera como regulador climático, la contaminación lumínica en Antofagasta, la protección del cielo oscuro para la astronomía mundial y la fauna costera.';
    const h4 = RAW_DATA.hitos['Hito 4'];
    if (h4) hakeEl.innerText = `g = +${h4.hake_gain}`;
  } else if (CURRENT_HITO === 'Hito 3') {
    titleEl.innerText = 'Los otros vecinos de Antofagasta (Hito 3)';
    descEl.innerText = 'Evaluación del aprendizaje sobre biodiversidad del desierto costero, tramas tróficas, fauna y flora aledaña a la ciudad de Antofagasta (Gaviotín Chico, descomponedores) y el humedal La Chimba.';
    const h3 = RAW_DATA.hitos['Hito 3'];
    if (h3) hakeEl.innerText = `g = +${h3.hake_gain}`;
  } else {
    titleEl.innerText = 'Resumen Global del Programa (Hito 3 y Hito 4)';
    descEl.innerText = 'Consolidado general de aprendizajes en la Línea Verde del Programa de Acción Climática (PAC 2026), integrando tramas tróficas, biodiversidad aledaña a Antofagasta, atmósfera y preservación de cielos limpios.';
    const glob = RAW_DATA.hitos['Global'];
    if (glob) hakeEl.innerText = `g = +${glob.hake_gain}`;
    else hakeEl.innerText = 'Global PAC';
  }
}

function updateKPIs() {
  const filtered = getFilteredParticipants();
  const N = filtered.length;

  const nEl = document.getElementById('kpi-n');
  const preEl = document.getElementById('kpi-pre');
  const postEl = document.getElementById('kpi-post');
  const deltaEl = document.getElementById('kpi-delta');
  const pctMejoraEl = document.getElementById('kpi-pct-mejora');
  const countMejoraEl = document.getElementById('kpi-count-mejora');
  const pctEstancadosEl = document.getElementById('kpi-pct-estancados');
  const countEstancadosEl = document.getElementById('kpi-count-estancados');

  if (N === 0) {
    nEl.innerText = '0';
    preEl.innerHTML = '0.0 <span class="text-xs font-normal text-slate-400">/ 5</span>';
    postEl.innerHTML = '0.0 <span class="text-xs font-normal text-slate-400">/ 5</span>';
    deltaEl.innerText = '0.00';
    pctMejoraEl.innerText = '0%';
    countMejoraEl.innerText = '0 participantes';
    pctEstancadosEl.innerText = '0%';
    countEstancadosEl.innerText = '0 participantes';
    return;
  }

  const avgPre = filtered.reduce((acc, p) => acc + p.pre_score, 0) / N;
  const avgPost = filtered.reduce((acc, p) => acc + p.post_score, 0) / N;
  const avgDelta = filtered.reduce((acc, p) => acc + p.delta, 0) / N;

  const nMejora = filtered.filter(p => p.status === 'Mejoró').length;
  const nMantuvo = filtered.filter(p => p.status === 'Mantuvo').length;
  const nRetrocedio = filtered.filter(p => p.status === 'Retrocedió').length;

  const pctMejora = ((nMejora / N) * 100).toFixed(1);
  const pctEstancados = (((nMantuvo + nRetrocedio) / N) * 100).toFixed(1);

  nEl.innerText = N;
  preEl.innerHTML = `${avgPre.toFixed(2)} <span class="text-xs font-normal text-slate-400">/ 5</span>`;
  postEl.innerHTML = `${avgPost.toFixed(2)} <span class="text-xs font-normal text-slate-400">/ 5</span>`;
  deltaEl.innerText = `${avgDelta >= 0 ? '+' : ''}${avgDelta.toFixed(2)}`;
  deltaEl.className = avgDelta >= 0 ? 'text-3xl font-extrabold text-[#00a87e] tracking-tight mt-1.5' : 'text-3xl font-extrabold text-rose-600 tracking-tight mt-1.5';

  pctMejoraEl.innerText = `${pctMejora}%`;
  countMejoraEl.innerText = `${nMejora} participantes subieron`;

  pctEstancadosEl.innerText = `${pctEstancados}%`;
  countEstancadosEl.innerText = `${nMantuvo} mantuvieron, ${nRetrocedio} bajaron`;
}

// RENDER CHARTS CON CHART.JS (Estilo moderno y gradientes inspirados en diseño de referencia)
function renderCharts() {
  const filtered = getFilteredParticipants();
  const N = filtered.length;

  // 1. Status Donut Chart
  const nMejora = filtered.filter(p => p.status === 'Mejoró').length;
  const nMantuvo = filtered.filter(p => p.status === 'Mantuvo').length;
  const nRetrocedio = filtered.filter(p => p.status === 'Retrocedió').length;

  const ctxStatus = document.getElementById('chart-status-donut').getContext('2d');
  if (CHARTS['status']) CHARTS['status'].destroy();

  CHARTS['status'] = new Chart(ctxStatus, {
    type: 'doughnut',
    data: {
      labels: ['Mejoró (+pts)', 'Mantuvo (0 pts)', 'Retrocedió (-pts)'],
      datasets: [{
        data: [nMejora, nMantuvo, nRetrocedio],
        backgroundColor: ['#00d09c', '#6366f1', '#ff5b5b'],
        borderWidth: 3,
        borderColor: '#ffffff',
        hoverOffset: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      cutout: '72%'
    }
  });

  const legendEl = document.getElementById('legend-status');
  legendEl.innerHTML = `
    <div class="bg-slate-50/80 p-2.5 rounded-2xl border border-slate-100">
      <div class="flex items-center justify-center gap-1.5">
        <span class="inline-block w-2.5 h-2.5 rounded-full bg-[#00d09c]"></span>
        <span class="font-bold text-slate-800">${nMejora}</span>
      </div>
      <div class="text-[10px] text-slate-400 font-medium mt-0.5">Mejoró (${N ? ((nMejora/N)*100).toFixed(0) : 0}%)</div>
    </div>
    <div class="bg-slate-50/80 p-2.5 rounded-2xl border border-slate-100">
      <div class="flex items-center justify-center gap-1.5">
        <span class="inline-block w-2.5 h-2.5 rounded-full bg-[#6366f1]"></span>
        <span class="font-bold text-slate-800">${nMantuvo}</span>
      </div>
      <div class="text-[10px] text-slate-400 font-medium mt-0.5">Mantuvo (${N ? ((nMantuvo/N)*100).toFixed(0) : 0}%)</div>
    </div>
    <div class="bg-slate-50/80 p-2.5 rounded-2xl border border-slate-100">
      <div class="flex items-center justify-center gap-1.5">
        <span class="inline-block w-2.5 h-2.5 rounded-full bg-[#ff5b5b]"></span>
        <span class="font-bold text-slate-800">${nRetrocedio}</span>
      </div>
      <div class="text-[10px] text-slate-400 font-medium mt-0.5">Retrocedió (${N ? ((nRetrocedio/N)*100).toFixed(0) : 0}%)</div>
    </div>
  `;

  // 2. Question Accuracy Bar Chart / Milestone Comparison
  const isGlobal = CURRENT_HITO === 'Global' || CURRENT_HITO === 'Ambos';
  let labels = [];
  let prePcts = [];
  let postPcts = [];

  if (isGlobal) {
    const h3 = RAW_DATA.hitos['Hito 3'];
    const h4 = RAW_DATA.hitos['Hito 4'];
    const glob = RAW_DATA.hitos['Global'];

    labels = ['Hito 3: Otros Vecinos', 'Hito 4: Escudo Invisible', 'Promedio General PAC'];
    const h3Pre = h3 ? ((h3.avg_pre / 5) * 100).toFixed(1) : 0;
    const h3Post = h3 ? ((h3.avg_post / 5) * 100).toFixed(1) : 0;
    const h4Pre = h4 ? ((h4.avg_pre / 5) * 100).toFixed(1) : 0;
    const h4Post = h4 ? ((h4.avg_post / 5) * 100).toFixed(1) : 0;
    const globPre = glob ? ((glob.avg_pre / 5) * 100).toFixed(1) : 0;
    const globPost = glob ? ((glob.avg_post / 5) * 100).toFixed(1) : 0;

    prePcts = [h3Pre, h4Pre, globPre];
    postPcts = [h3Post, h4Post, globPost];
  } else {
    const hitoKey = CURRENT_HITO;
    const questionsMeta = RAW_DATA.hitos[hitoKey].questions;
    const key = questionsMeta.map(q => q.correct_key.toLowerCase());
    labels = ['Pregunta 1', 'Pregunta 2', 'Pregunta 3', 'Pregunta 4', 'Pregunta 5'];

    for (let i = 0; i < 5; i++) {
      const preC = filtered.filter(p => p.pre_mc && p.pre_mc[i] === key[i]).length;
      const postC = filtered.filter(p => p.post_mc && p.post_mc[i] === key[i]).length;
      prePcts.push(N ? ((preC / N) * 100).toFixed(1) : 0);
      postPcts.push(N ? ((postC / N) * 100).toFixed(1) : 0);
    }
  }

  const ctxQuestions = document.getElementById('chart-questions-bar').getContext('2d');
  if (CHARTS['questions']) CHARTS['questions'].destroy();

  // Create smooth gradients matching the reference image (Purple to Cyan to Mint)
  const gradPost = ctxQuestions.createLinearGradient(0, 220, 0, 10);
  gradPost.addColorStop(0, '#6366f1');   // Royal Purple/Indigo
  gradPost.addColorStop(0.5, '#00c5cb'); // Cyan
  gradPost.addColorStop(1, '#00d09c');   // Mint Green

  const gradPre = ctxQuestions.createLinearGradient(0, 220, 0, 10);
  gradPre.addColorStop(0, '#94a3b8');    // Slate
  gradPre.addColorStop(1, '#cbd5e1');    // Light Slate

  CHARTS['questions'] = new Chart(ctxQuestions, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: '% Aciertos PRE (Basal)',
          data: prePcts,
          backgroundColor: gradPre,
          borderRadius: 8,
          borderSkipped: false
        },
        {
          label: '% Aciertos POST (Terreno)',
          data: postPcts,
          backgroundColor: gradPost,
          borderRadius: 8,
          borderSkipped: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          grid: { color: '#f8fafc' },
          ticks: { callback: v => v + '%', font: { size: 10 } }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 11, weight: '600' } }
        }
      },
      plugins: {
        legend: {
          position: 'top',
          labels: {
            boxWidth: 12,
            boxHeight: 12,
            useBorderRadius: true,
            borderRadius: 4,
            font: { size: 11, weight: '600' }
          }
        }
      }
    }
  });

  // 3. Schools Delta Bar Chart
  const hitoKey = isGlobal ? 'Global' : CURRENT_HITO;
  const schoolsMeta = RAW_DATA.hitos[hitoKey].schools.filter(s => s.has_data !== false);
  const schoolLabels = schoolsMeta.map(s => s.school.replace('Escuela ', 'Esc. ').replace('Colegio ', 'Col. '));
  const schoolDeltas = schoolsMeta.map(s => s.delta_avg);

  const ctxSchools = document.getElementById('chart-schools-delta').getContext('2d');
  if (CHARTS['schools']) CHARTS['schools'].destroy();

  const gradSchoolPos = ctxSchools.createLinearGradient(0, 220, 0, 10);
  gradSchoolPos.addColorStop(0, '#00c5cb'); // Cyan
  gradSchoolPos.addColorStop(1, '#00d09c'); // Mint

  const gradSchoolNeg = ctxSchools.createLinearGradient(0, 220, 0, 10);
  gradSchoolNeg.addColorStop(0, '#f43f5e'); // Rose
  gradSchoolNeg.addColorStop(1, '#ff5b5b'); // Coral

  const schoolColors = schoolDeltas.map(d => d >= 0 ? gradSchoolPos : gradSchoolNeg);

  CHARTS['schools'] = new Chart(ctxSchools, {
    type: 'bar',
    data: {
      labels: schoolLabels,
      datasets: [{
        label: 'Delta Promedio (Puntos)',
        data: schoolDeltas,
        backgroundColor: schoolColors,
        borderRadius: 8,
        borderSkipped: false
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: {
          grid: { color: '#f8fafc' },
          ticks: { callback: v => (v >= 0 ? '+' : '') + v, font: { size: 10 } }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 11, weight: '600' } }
        }
      }
    }
  });

  // 4. Gender Comparison Chart
  const genderData = RAW_DATA.hitos[hitoKey].gender;
  const ctxGender = document.getElementById('chart-gender-comparison').getContext('2d');
  if (CHARTS['gender']) CHARTS['gender'].destroy();

  const gLabels = Object.keys(genderData);
  const gPre = gLabels.map(g => genderData[g].pre_avg);
  const gPost = gLabels.map(g => genderData[g].post_avg);

  const gradGenderPost = ctxGender.createLinearGradient(0, 220, 0, 10);
  gradGenderPost.addColorStop(0, '#6366f1');
  gradGenderPost.addColorStop(0.5, '#00c5cb');
  gradGenderPost.addColorStop(1, '#00d09c');

  const gradGenderPre = ctxGender.createLinearGradient(0, 220, 0, 10);
  gradGenderPre.addColorStop(0, '#94a3b8');
  gradGenderPre.addColorStop(1, '#cbd5e1');

  CHARTS['gender'] = new Chart(ctxGender, {
    type: 'bar',
    data: {
      labels: gLabels.map(g => `${g} (N=${genderData[g].n})`),
      datasets: [
        {
          label: 'Puntaje PRE',
          data: gPre,
          backgroundColor: gradGenderPre,
          borderRadius: 8,
          borderSkipped: false
        },
        {
          label: 'Puntaje POST',
          data: gPost,
          backgroundColor: gradGenderPost,
          borderRadius: 8,
          borderSkipped: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          max: 5,
          grid: { color: '#f8fafc' },
          ticks: { font: { size: 10 } }
        },
        x: {
          grid: { display: false },
          ticks: { font: { size: 11, weight: '600' } }
        }
      },
      plugins: {
        legend: {
          position: 'top',
          labels: {
            boxWidth: 12,
            boxHeight: 12,
            useBorderRadius: true,
            borderRadius: 4,
            font: { size: 11, weight: '600' }
          }
        }
      }
    }
  });
}

// 2. TAB: ESTABLECIMIENTOS TABLE
function renderSchoolsTab() {
  const hitoKey = (CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global') ? 'Global' : CURRENT_HITO;
  const schools = RAW_DATA.hitos[hitoKey].schools;
  const tbody = document.getElementById('schools-table-body');
  tbody.innerHTML = '';

  schools.forEach(s => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/70 transition';

    if (s.has_data === false || s.n === 0) {
      tr.innerHTML = `
        <td class="py-3.5 px-4 font-semibold text-slate-800">${s.school}</td>
        <td class="py-3.5 px-3">
          <span class="px-2.5 py-1 rounded-full text-[11px] bg-slate-100 text-slate-700 font-medium">${s.dependencia}</span>
        </td>
        <td class="py-3.5 px-3 text-center text-slate-400 italic font-semibold">0 (Sin datos)</td>
        <td class="py-3.5 px-3 text-center text-slate-400">-</td>
        <td class="py-3.5 px-3 text-center text-slate-400">-</td>
        <td class="py-3.5 px-3 text-center">
          <span class="px-2.5 py-1 rounded-full text-[10px] bg-amber-50 text-amber-800 border border-amber-200 font-medium">Sin datos en este hito</span>
        </td>
        <td class="py-3.5 px-3 text-center text-slate-400">-</td>
        <td class="py-3.5 px-3 text-center text-slate-400 text-[10px]">No evaluado</td>
        <td class="py-3.5 px-4">
          <div class="text-slate-800 font-medium">${s.lider1 || 'Coordinador del equipo'}</div>
          <div class="text-[10px] text-slate-400">${s.email1 || ''} ${s.tel1 ? '· Tel: ' + s.tel1 : ''}</div>
        </td>
      `;
    } else {
      const deltaSign = s.delta_avg >= 0 ? '+' : '';
      const deltaBadge = s.delta_avg > 0 ? 
        `<span class="px-3 py-1 rounded-full text-xs font-bold bg-[#00d09c]/15 text-[#00a87e] border border-[#00d09c]/25">+${s.delta_avg.toFixed(2)}</span>` : 
        (s.delta_avg === 0 ? 
          `<span class="px-3 py-1 rounded-full text-xs font-semibold bg-slate-100 text-slate-600">0.00</span>` : 
          `<span class="px-3 py-1 rounded-full text-xs font-bold bg-rose-50 text-rose-600 border border-rose-200">${s.delta_avg.toFixed(2)}</span>`);

      tr.innerHTML = `
        <td class="py-3.5 px-4 font-bold text-slate-800">${s.school}</td>
        <td class="py-3.5 px-3">
          <span class="px-2.5 py-1 rounded-full text-[11px] bg-slate-100 text-slate-700 font-medium">${s.dependencia}</span>
        </td>
        <td class="py-3.5 px-3 text-center font-bold text-slate-800">${s.n}</td>
        <td class="py-3.5 px-3 text-center text-slate-600 font-medium">${s.pre_avg !== null ? s.pre_avg.toFixed(2) : '-'}</td>
        <td class="py-3.5 px-3 text-center text-[#00a87e] font-bold">${s.post_avg !== null ? s.post_avg.toFixed(2) : '-'}</td>
        <td class="py-3.5 px-3 text-center">${deltaBadge}</td>
        <td class="py-3.5 px-3 text-center">
          <span class="px-3 py-1 rounded-full text-xs font-bold ${s.pct_mejora >= 50 ? 'bg-[#00d09c]/20 text-[#00a87e]' : 'bg-slate-100 text-slate-700'}">
            ${s.pct_mejora}%
          </span>
        </td>
        <td class="py-3.5 px-3">
          <div class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-slate-50 border border-slate-100 text-[11px]">
            <span class="text-[#00a87e] font-bold" title="Mejoró">${s.status_counts['Mejoró'] || 0}↗</span>
            <span class="text-slate-300">·</span>
            <span class="text-slate-500 font-semibold" title="Mantuvo">${s.status_counts['Mantuvo'] || 0}=</span>
            <span class="text-slate-300">·</span>
            <span class="text-rose-600 font-bold" title="Retrocedió">${s.status_counts['Retrocedió'] || 0}↘</span>
          </div>
        </td>
        <td class="py-3.5 px-4">
          <div class="text-slate-800 font-medium">${s.lider1 || 'Coordinador del equipo'}</div>
          <div class="text-[10px] text-slate-400">${s.email1 || ''} ${s.tel1 ? '· Tel: ' + s.tel1 : ''}</div>
        </td>
      `;
    }
    tbody.appendChild(tr);
  });
}

// 3. TAB: PREGUNTAS Y DISTRACTORES
function renderQuestionsTab() {
  const isGlobal = CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global';
  const hitoKey = isGlobal ? 'Global' : CURRENT_HITO;
  const questions = RAW_DATA.hitos[hitoKey].questions;
  const container = document.getElementById('questions-detail-container');
  container.innerHTML = '';

  let currentSectionHito = null;

  questions.forEach((q, idx) => {
    if (isGlobal) {
      const qHito = q.hito || (idx < 5 ? 'Hito 3' : 'Hito 4');
      if (qHito !== currentSectionHito) {
        currentSectionHito = qHito;
        const sectionHeader = document.createElement('div');
        sectionHeader.className = qHito === 'Hito 3' ? 
          'p-4 bg-emerald-50 rounded-2xl border border-emerald-200 text-emerald-950 font-bold text-xs flex items-center gap-2 mt-2 mb-1 shadow-sm' :
          'p-4 bg-teal-50 rounded-2xl border border-teal-200 text-teal-950 font-bold text-xs flex items-center gap-2 mt-6 mb-1 shadow-sm';
        sectionHeader.innerHTML = `
          <i data-lucide="${qHito === 'Hito 3' ? 'leaf' : 'shield'}" class="w-4 h-4"></i>
          <span>${qHito === 'Hito 3' ? 'Hito 3: Los otros vecinos de Antofagasta (Biodiversidad y tramas tróficas)' : 'Hito 4: Proteger nuestro escudo invisible (Atmósfera y cielos oscuros)'}</span>
        `;
        container.appendChild(sectionHeader);
      }
    }

    const card = document.createElement('div');
    card.className = 'border border-slate-100 rounded-2xl p-5 bg-white card-shadow hover:border-slate-200 transition space-y-3.5';

    const distractorPills = Object.entries(q.post_distractors)
      .map(([opt, count]) => `<span class="px-2.5 py-1 rounded-full bg-rose-50 text-rose-700 border border-rose-200 font-semibold text-[11px]">Opción ${opt}: ${count} veces</span>`)
      .join(' ') || '<span class="text-slate-400 text-xs">Sin errores registrados</span>';

    const diffBadge = q.diff >= 0 ? 
      `<span class="badge-improved text-xs px-3 py-1 rounded-full font-bold">+${q.diff} aciertos (+${(q.post_pct - q.pre_pct).toFixed(1)}%)</span>` :
      `<span class="badge-regressed text-xs px-3 py-1 rounded-full font-bold">${q.diff} aciertos (${(q.post_pct - q.pre_pct).toFixed(1)}%)</span>`;

    card.innerHTML = `
      <div class="flex flex-wrap items-start justify-between gap-2">
        <div class="space-y-1">
          <div class="flex items-center gap-2.5">
            <span class="w-7 h-7 rounded-full bg-[#00d09c] text-slate-950 font-extrabold text-xs flex items-center justify-center shadow-sm">${q.num}</span>
            <span class="text-xs font-bold uppercase tracking-wider text-slate-400">Alternativa Correcta: [ ${q.correct_key} ]</span>
          </div>
          <h4 class="text-sm font-bold text-slate-800 pl-9.5">${q.title}</h4>
        </div>
        <div>${diffBadge}</div>
      </div>

      <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-50/70 p-3.5 rounded-xl border border-slate-100 text-xs pl-9">
        <div>
          <span class="text-slate-400 text-[10px] font-bold uppercase tracking-wider block">Aciertos PRE</span>
          <span class="font-extrabold text-slate-700 text-sm mt-0.5 block">${q.pre_correct} (${q.pre_pct}%)</span>
        </div>
        <div>
          <span class="text-slate-400 text-[10px] font-bold uppercase tracking-wider block">Aciertos POST</span>
          <span class="font-extrabold text-[#00a87e] text-sm mt-0.5 block">${q.post_correct} (${q.post_pct}%)</span>
        </div>
        <div>
          <span class="text-slate-400 text-[10px] font-bold uppercase tracking-wider block">Mantuvieron Correcto</span>
          <span class="font-extrabold text-slate-700 text-sm mt-0.5 block">${q.transitions['Mantuvo Correcto'] || 0}</span>
        </div>
        <div>
          <span class="text-slate-400 text-[10px] font-bold uppercase tracking-wider block">Retrocedieron</span>
          <span class="font-extrabold text-rose-600 text-sm mt-0.5 block">${q.transitions['Retrocedió'] || 0}</span>
        </div>
      </div>

      <div class="pl-9 pt-1 text-xs">
        <span class="font-semibold text-slate-500 mr-2">Distractores más seleccionados en el POST:</span>
        ${distractorPills}
      </div>
    `;
    container.appendChild(card);
  });
}

// 4. TAB: CUALITATIVO RIGUROSO
function renderQualitativeTab() {
  const isGlobal = CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global';
  const hitoKey = isGlobal ? 'Global' : CURRENT_HITO;
  const openQuestions = RAW_DATA.hitos[hitoKey].open_questions;
  const selectEl = document.getElementById('select-open-q');

  if (selectEl.children.length === 0 || selectEl.dataset.hito !== hitoKey) {
    selectEl.innerHTML = '';
    openQuestions.forEach((oq, idx) => {
      const opt = document.createElement('option');
      opt.value = idx;
      opt.textContent = isGlobal ? `${oq.substring(0, 75)}...` : `P${idx + 1}: ${oq.substring(0, 65)}...`;
      selectEl.appendChild(opt);
    });
    selectEl.dataset.hito = hitoKey;
  }

  const selectedIdx = parseInt(selectEl.value || 0);
  const filtered = getFilteredParticipants();
  const listEl = document.getElementById('qualitative-responses-list');
  listEl.innerHTML = '';

  filtered.forEach(p => {
    let qData = null;
    let originHito = p.hito;
    if (isGlobal) {
      if (selectedIdx < 4) {
        const p3 = p.h3 || (p.hito === 'Hito 3' ? p : null);
        if (!p3) return;
        qData = p3.open_analysis && p3.open_analysis[selectedIdx];
        originHito = 'Hito 3';
      } else {
        const p4 = p.h4 || (p.hito === 'Hito 4' ? p : null);
        if (!p4) return;
        qData = p4.open_analysis && p4.open_analysis[selectedIdx - 4];
        originHito = 'Hito 4';
      }
    } else {
      qData = p.open_analysis && p.open_analysis[selectedIdx];
    }

    if (!qData) return;

    const item = document.createElement('div');
    item.className = 'bg-white p-5 rounded-2xl border border-slate-100 card-shadow hover:border-slate-200 transition text-xs space-y-3';

    let badgeClass = 'bg-slate-100 text-slate-700 border border-slate-300';
    let badgeText = qData.status;

    if (qData.status === 'Sin Evidencia') {
      badgeClass = 'bg-slate-100 text-slate-600 border border-slate-300';
      badgeText = 'Sin Evidencia (Revisar con líder)';
    } else if (qData.status === 'Mejoró') {
      badgeClass = 'badge-improved font-bold';
      badgeText = 'Mejora Cualitativa ↗';
    } else if (qData.status === 'Mantuvo') {
      badgeClass = 'bg-blue-50 text-blue-700 border border-blue-200 font-medium';
      badgeText = 'Mantuvo Respuesta Informada =';
    } else if (qData.status === 'Retrocedió') {
      badgeClass = 'badge-regressed font-bold';
      badgeText = 'Retroceso en Texto ↘';
    }

    const reviewAlert = qData.requires_review ? 
      `<div class="mt-2 text-[11px] font-semibold text-amber-800 bg-amber-50 px-3 py-2 rounded-xl border border-amber-200 flex items-center gap-2">
        <i data-lucide="alert-triangle" class="w-4 h-4 text-amber-600 shrink-0"></i>
        <span>Elemento de evaluación no concluyente por falta de respuesta escrita. Se sugiere revisión individual con el profesor/a líder.</span>
      </div>` : '';

    item.innerHTML = `
      <div class="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-2.5">
        <div>
          <span class="font-extrabold text-slate-900 text-sm">${p.name}</span>
          <span class="text-slate-400 ml-1.5 font-medium">· ${p.school}</span>
          ${isGlobal ? `<span class="ml-2 px-2.5 py-0.5 rounded-full text-[10px] font-bold ${originHito === 'Hito 4' ? 'bg-indigo-100 text-indigo-800' : 'bg-emerald-100 text-emerald-800'}">${originHito}</span>` : ''}
        </div>
        <div class="flex items-center gap-2">
          <span class="px-3 py-1 rounded-full text-[11px] font-bold ${badgeClass}">${badgeText}</span>
        </div>
      </div>

      <div class="text-[11px] text-slate-500 font-medium">
        Diagnóstico conceptual: <span class="text-slate-800 font-semibold">${qData.detail}</span>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-3 pt-1">
        <div class="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
          <span class="text-[10px] font-bold text-slate-400 block uppercase tracking-wide mb-1">Respuesta PRE (Inicial):</span>
          <p class="text-slate-700 italic">${qData.pre_text && qData.pre_text.trim() ? '"' + qData.pre_text + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta o no sabe)</span>'}</p>
        </div>
        <div class="bg-emerald-50/40 p-3.5 rounded-xl border border-emerald-100">
          <span class="text-[10px] font-bold text-[#00a87e] block uppercase tracking-wide mb-1">Respuesta POST (Final en terreno):</span>
          <p class="text-slate-800 font-medium">${qData.post_text && qData.post_text.trim() ? '"' + qData.post_text + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta o no sabe)</span>'}</p>
        </div>
      </div>

      ${reviewAlert}
    `;
    listEl.appendChild(item);
  });
}

// 5. TAB: CASOS CRÍTICOS Y ORIENTACIONES PEDAGÓGICAS
function renderPedagogicalTab() {
  const hitoKey = (CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global') ? 'Global' : CURRENT_HITO;
  const schools = RAW_DATA.hitos[hitoKey].schools;
  const container = document.getElementById('pedagogical-cards');
  container.innerHTML = '';

  schools.forEach(s => {
    const card = document.createElement('div');
    card.className = 'border border-slate-100 rounded-2xl p-5 bg-white card-shadow space-y-3.5';

    let recomendacion = '';
    let statusClass = 'bg-emerald-50 border-emerald-200 text-emerald-900';

    if (s.has_data === false || s.n === 0) {
      recomendacion = `<strong>Sin evaluaciones registradas en este hito:</strong> Este establecimiento no registra respuestas PRE ni POST para este desafío. Se sugiere coordinar directamente con el profesor/a líder (<strong>${s.lider1 || 'Coordinador'}</strong>) para verificar la ejecución de la actividad en terreno o aplicar la evaluación de manera diferida.`;
      statusClass = 'bg-amber-50 border-amber-200 text-amber-950';
    } else if (s.delta_avg > 0.7) {
      recomendacion = `<strong>Consolidación e Ideación:</strong> Este equipo demostró una comprensión sobresaliente de los conceptos de la intervención (+${s.delta_avg} pts). Se sugiere a la líder <strong>${s.lider1 || 'Docente Líder'}</strong> canalizar este entusiasmo directamente hacia la etapa de ideación (ej. iniciativas escolares de protección de cielos oscuros o conservación de la biodiversidad local).`;
      statusClass = 'bg-emerald-50 border-emerald-200 text-emerald-950';
    } else if (s.delta_avg >= 0.2) {
      recomendacion = `<strong>Refuerzo de conceptos intermedios:</strong> El equipo tiene una asimilación positiva pero heterogénea (+${s.delta_avg} pts). Se recomienda a la líder <strong>${s.lider1 || 'Docente Líder'}</strong> realizar un plenario corto de 15 minutos repasando las diferencias conceptuales clave, antes de pasar a la etapa de ideación.`;
      statusClass = 'bg-amber-50 border-amber-200 text-amber-950';
    } else {
      recomendacion = `<strong>Atención a 'Efecto Techo' y Distractores:</strong> El puntaje basal fue muy elevado o se registraron confusiones sutiles en las opciones cerradas (${s.delta_avg} pts). No interpretar esto como falta de aprendizaje: cualitativamente los estudiantes manejan la terminología. Se aconseja pedirles que expliquen con sus propias palabras el impacto en la fauna local para afianzar la seguridad en sus conocimientos antes de la etapa de ideación.`;
      statusClass = 'bg-slate-50 border-slate-200 text-slate-800';
    }

    card.innerHTML = `
      <div class="flex items-center justify-between border-b border-slate-100 pb-2.5">
        <div>
          <h4 class="font-extrabold text-slate-900 text-sm">${s.school}</h4>
          <span class="text-[11px] text-slate-500 font-medium">Líder: ${s.lider1 || 'Coordinador del equipo'} ${s.lider2 ? '/ ' + s.lider2 : ''}</span>
        </div>
        <div class="text-right">
          ${s.has_data !== false ? `
            <span class="text-xs font-bold px-2.5 py-0.5 rounded-full ${s.delta_avg >= 0 ? 'bg-[#00d09c]/15 text-[#00a87e]' : 'bg-rose-50 text-rose-600'}">Delta: ${s.delta_avg >= 0 ? '+' : ''}${s.delta_avg}</span>
            <div class="text-[10px] text-slate-400 font-medium mt-0.5">${s.pct_mejora}% mejoró</div>
          ` : `
            <span class="text-xs font-bold text-amber-800 bg-amber-100 px-2.5 py-0.5 rounded-full">Sin datos</span>
          `}
        </div>
      </div>

      <div class="p-3.5 rounded-xl border text-xs leading-relaxed ${statusClass}">
        ${recomendacion}
      </div>

      <div class="flex items-center justify-between text-[11px] text-slate-500 pt-1">
        <span>Estudiantes evaluados: <strong>${s.n}</strong></span>
        <span>${s.email1 ? '✉ ' + s.email1 : ''} ${s.tel1 ? '· Tel: ' + s.tel1 : ''}</span>
      </div>
    `;
    container.appendChild(card);
  });
}

// Helper para sanitizar texto HTML
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// 6. TAB: FICHAS INDIVIDUALES
function renderParticipantsTab() {
  const isGlobal = CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global';
  const filtered = getFilteredParticipants();
  document.getElementById('participants-count-filtered').innerText = filtered.length;

  const tbody = document.getElementById('participants-table-body');
  tbody.innerHTML = '';

  filtered.forEach((p, idx) => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/70 transition cursor-pointer';
    tr.onclick = () => openModal(p);

    const deltaSign = p.delta >= 0 ? '+' : '';
    const badgeClass = p.status === 'Mejoró' ? 'badge-improved' : (p.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained');

    let hitoBadge = '';
    if (isGlobal) {
      if (p.hitos_completados && p.hitos_completados.length > 1) {
        hitoBadge = `<span class="ml-1.5 px-2 py-0.5 rounded-full text-[10px] font-bold bg-[#00d09c]/15 text-[#00a87e] border border-[#00d09c]/30">Ambos Hitos (H3 + H4)</span>`;
      } else if (p.hito === 'Hito 3') {
        hitoBadge = `<span class="ml-1.5 px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800">Solo Hito 3</span>`;
      } else if (p.hito === 'Hito 4') {
        hitoBadge = `<span class="ml-1.5 px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-100 text-indigo-800">Solo Hito 4</span>`;
      }
    }

    const preDisplay = isGlobal && p.h3 && p.h4 ? 
      `<div class="font-bold text-slate-800">${p.pre_score} <span class="text-[10px] text-slate-400 font-normal">/ 5</span></div><div class="text-[10px] text-slate-400 font-normal">(${p.pre_score_raw} / 10 pts)</div>` : 
      `<span class="font-bold text-slate-800">${p.pre_score}</span> <span class="text-[10px] text-slate-400">/ 5</span>`;

    const postDisplay = isGlobal && p.h3 && p.h4 ? 
      `<div class="font-bold text-[#00a87e]">${p.post_score} <span class="text-[10px] text-emerald-600 font-normal">/ 5</span></div><div class="text-[10px] text-emerald-600 font-normal">(${p.post_score_raw} / 10 pts)</div>` : 
      `<span class="font-bold text-[#00a87e]">${p.post_score}</span> <span class="text-[10px] text-slate-400">/ 5</span>`;

    tr.innerHTML = `
      <td class="py-3.5 px-4 font-bold text-slate-900 flex items-center gap-2.5">
        <span class="w-2.5 h-2.5 rounded-full ${p.gender === 'Femenino' ? 'bg-pink-400' : 'bg-blue-400'}"></span>
        ${escapeHtml(p.name)}
      </td>
      <td class="py-3.5 px-3 text-slate-600 font-medium">${escapeHtml(p.school)} ${hitoBadge}</td>
      <td class="py-3.5 px-3 text-slate-500 text-[11px]">${escapeHtml(p.gender)}</td>
      <td class="py-3.5 px-3 text-center">${preDisplay}</td>
      <td class="py-3.5 px-3 text-center">${postDisplay}</td>
      <td class="py-3.5 px-3 text-center font-bold ${p.delta >= 0 ? 'text-[#00a87e]' : 'text-rose-600'}">${deltaSign}${p.delta}</td>
      <td class="py-3.5 px-3 text-center">
        <span class="px-2.5 py-0.5 rounded-full text-[11px] font-bold ${badgeClass}">${p.status}</span>
        <div class="text-[10px] text-slate-400 mt-0.5 font-medium">${escapeHtml(p.real_status || '')}</div>
      </td>
      <td class="py-3.5 px-4 text-center">
        <button onclick="event.stopPropagation(); openModal(window.currentFiltered[${idx}])" class="px-3.5 py-1 text-xs font-bold text-slate-900 bg-[#00d09c] hover:bg-[#00bfa5] rounded-full transition shadow-sm">
          Ver Ficha
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  window.currentFiltered = filtered;
}

// Variables de estado del modal
let currentModalMode = 'consolidado';
let currentModalParticipant = null;

// MODAL INDIVIDUAL PARTICIPANTE (Soporte integral de ambos desafíos y test-retest)
function openModal(p) {
  if (!p) return;

  const allData = getParticipantAllData(p.name);
  const h3Data = p.h3 || allData.h3;
  const h4Data = p.h4 || allData.h4;
  p.h3 = h3Data;
  p.h4 = h4Data;
  currentModalParticipant = p;

  // Actualizar metadatos en header del modal
  document.getElementById('modal-name').innerText = p.name;
  document.getElementById('modal-school').innerText = `${p.school} · Género: ${p.gender}`;
  document.getElementById('modal-role').innerText = p.real_status || 'Estudiante';

  // Badge de cobertura de evaluación
  const badgeContainer = document.getElementById('modal-attendance-badge');
  if (h3Data && h4Data) {
    badgeContainer.innerHTML = `<span class="text-[10px] font-bold px-2.5 py-0.5 rounded-full bg-teal-50 text-teal-700 border border-teal-200">Test-Retest Completo (Hito 3 y Hito 4)</span>`;
  } else if (h3Data) {
    badgeContainer.innerHTML = `<span class="text-[10px] font-bold px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">Evaluado en Hito 3</span>`;
  } else if (h4Data) {
    badgeContainer.innerHTML = `<span class="text-[10px] font-bold px-2.5 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200">Evaluado en Hito 4</span>`;
  } else {
    badgeContainer.innerHTML = '';
  }

  // Visibilidad de los botones del selector según disponibilidad de datos
  const btnConsolidado = document.getElementById('btn-modal-mode-consolidado');
  const btnH3 = document.getElementById('btn-modal-mode-h3');
  const btnH4 = document.getElementById('btn-modal-mode-h4');

  if (btnConsolidado) {
    btnConsolidado.style.display = (h3Data && h4Data) ? 'inline-block' : 'none';
  }
  if (btnH3) {
    btnH3.style.display = h3Data ? 'inline-block' : 'none';
  }
  if (btnH4) {
    btnH4.style.display = h4Data ? 'inline-block' : 'none';
  }

  // Determinar modo inicial
  let initialMode = 'consolidado';
  if (CURRENT_HITO === 'Hito 3' && h3Data) {
    initialMode = 'Hito 3';
  } else if (CURRENT_HITO === 'Hito 4' && h4Data) {
    initialMode = 'Hito 4';
  } else if (h3Data && h4Data) {
    initialMode = 'consolidado';
  } else if (h3Data) {
    initialMode = 'Hito 3';
  } else if (h4Data) {
    initialMode = 'Hito 4';
  }

  setModalViewMode(initialMode);
  document.getElementById('participant-modal').classList.remove('hidden');
  if (window.lucide) lucide.createIcons();
}

function setModalViewMode(mode) {
  currentModalMode = mode;
  const p = currentModalParticipant;
  if (!p) return;

  const btnConsolidado = document.getElementById('btn-modal-mode-consolidado');
  const btnH3 = document.getElementById('btn-modal-mode-h3');
  const btnH4 = document.getElementById('btn-modal-mode-h4');
  const caption = document.getElementById('modal-mode-caption');

  const activeClasses = 'px-3.5 py-1 rounded-full text-xs font-bold bg-[#00d09c] text-slate-950 shadow-sm transition';
  const inactiveClasses = 'px-3.5 py-1 rounded-full text-xs font-medium text-slate-600 hover:text-slate-900 transition';

  if (btnConsolidado) btnConsolidado.className = mode === 'consolidado' ? activeClasses : inactiveClasses;
  if (btnH3) btnH3.className = mode === 'Hito 3' ? activeClasses : inactiveClasses;
  if (btnH4) btnH4.className = mode === 'Hito 4' ? activeClasses : inactiveClasses;

  if (caption) {
    if (mode === 'consolidado') {
      caption.innerText = 'Desglose integral combinado (10 preguntas cerradas y 8 cualitativas)';
    } else if (mode === 'Hito 3') {
      caption.innerText = 'Desglose Hito 3: Los otros vecinos de Antofagasta (Biodiversidad y Humedales)';
    } else if (mode === 'Hito 4') {
      caption.innerText = 'Desglose Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos Limpios)';
    }
  }

  renderModalContent(mode);
}

// Renderiza el contenido dinámico del modal según el modo seleccionado
function renderModalContent(mode) {
  const container = document.getElementById('modal-content-area');
  if (!container || !currentModalParticipant) return;
  const p = currentModalParticipant;
  const h3Data = p.h3;
  const h4Data = p.h4;

  const h3Meta = (RAW_DATA.hitos['Hito 3'] && RAW_DATA.hitos['Hito 3'].questions) || [];
  const h4Meta = (RAW_DATA.hitos['Hito 4'] && RAW_DATA.hitos['Hito 4'].questions) || [];

  container.innerHTML = '';

  if (mode === 'consolidado') {
    // 1. Resumen consolidado KPIs
    const h3Pre = h3Data ? h3Data.pre_score : 0;
    const h3Post = h3Data ? h3Data.post_score : 0;
    const h4Pre = h4Data ? h4Data.pre_score : 0;
    const h4Post = h4Data ? h4Data.post_score : 0;
    const totPre = h3Pre + h4Pre;
    const totPost = h3Post + h4Post;
    const totDelta = totPost - totPre;
    const deltaSign = totDelta >= 0 ? '+' : '';

    const kpiSummary = document.createElement('div');
    kpiSummary.className = 'grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-50/90 p-4 rounded-2xl border border-slate-200 text-center';
    kpiSummary.innerHTML = `
      <div class="p-2.5 bg-white rounded-xl border border-slate-100 shadow-sm">
        <div class="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Puntaje Total Combinado</div>
        <div class="text-xl font-extrabold text-slate-900 mt-1">${totPost} <span class="text-xs text-slate-400 font-medium">/ 10 pts</span></div>
        <div class="text-[11px] text-slate-500 font-medium mt-0.5">PRE: ${totPre} / 10 · ${deltaSign}${totDelta} pts</div>
      </div>
      <div class="p-2.5 bg-emerald-50/40 rounded-xl border border-emerald-100 shadow-sm">
        <div class="text-[10px] text-emerald-800 font-bold uppercase tracking-wider">Hito 3 (Biodiversidad)</div>
        <div class="text-xl font-extrabold text-[#00a87e] mt-1">${h3Data ? h3Data.post_score : '-'} <span class="text-xs text-slate-400 font-medium">/ 5 pts</span></div>
        <div class="text-[11px] text-slate-500 font-medium mt-0.5">${h3Data ? `PRE: ${h3Data.pre_score} · ${h3Data.delta >= 0 ? '+' : ''}${h3Data.delta} pts` : 'No registrado'}</div>
      </div>
      <div class="p-2.5 bg-indigo-50/40 rounded-xl border border-indigo-100 shadow-sm">
        <div class="text-[10px] text-indigo-800 font-bold uppercase tracking-wider">Hito 4 (Atmósfera)</div>
        <div class="text-xl font-extrabold text-indigo-700 mt-1">${h4Data ? h4Data.post_score : '-'} <span class="text-xs text-slate-400 font-medium">/ 5 pts</span></div>
        <div class="text-[11px] text-slate-500 font-medium mt-0.5">${h4Data ? `PRE: ${h4Data.pre_score} · ${h4Data.delta >= 0 ? '+' : ''}${h4Data.delta} pts` : 'No registrado'}</div>
      </div>
      <div class="p-2.5 bg-white rounded-xl border border-slate-100 shadow-sm">
        <div class="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Diagnóstico de Aprendizajes</div>
        <div class="mt-1">
          <span class="inline-block font-bold px-2.5 py-0.5 rounded-full text-[11px] ${p.status === 'Mejoró' ? 'badge-improved' : (p.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained')}">${p.status}</span>
        </div>
        <div class="text-[10px] text-slate-500 font-semibold mt-1">${escapeHtml(p.real_status || 'Rendimiento Registrado')}</div>
      </div>
    `;
    container.appendChild(kpiSummary);

    // 2. Preguntas Cerradas
    const mcSection = document.createElement('div');
    mcSection.className = 'space-y-4';
    mcSection.innerHTML = `
      <div class="flex items-center gap-2 border-b border-slate-200 pb-2">
        <i data-lucide="check-circle-2" class="w-4 h-4 text-[#00d09c]"></i>
        <h4 class="font-bold text-slate-900 text-sm">Desglose de Preguntas Cerradas (Alternativas) - Ambos Desafíos</h4>
      </div>
      <div class="space-y-4">
        ${buildHitoMCBlock('Hito 3: Los otros vecinos de Antofagasta (Biodiversidad y Humedales)', h3Data, h3Meta, 'emerald')}
        ${buildHitoMCBlock('Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos Limpios)', h4Data, h4Meta, 'indigo')}
      </div>
    `;
    container.appendChild(mcSection);

    // 3. Preguntas Abiertas
    const openSection = document.createElement('div');
    openSection.className = 'space-y-4';
    openSection.innerHTML = `
      <div class="flex items-center gap-2 border-b border-slate-200 pb-2">
        <i data-lucide="message-square" class="w-4 h-4 text-[#00d09c]"></i>
        <h4 class="font-bold text-slate-900 text-sm">Desglose de Preguntas Abiertas (Cualitativo) - Ambos Desafíos</h4>
      </div>
      <div class="space-y-4">
        ${buildHitoOpenBlock('Hito 3: Los otros vecinos de Antofagasta (Biodiversidad y Ecosistemas)', h3Data, 'emerald')}
        ${buildHitoOpenBlock('Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos Limpios)', h4Data, 'indigo')}
      </div>
    `;
    container.appendChild(openSection);

  } else if (mode === 'Hito 3') {
    const dataH = h3Data;
    if (!dataH) {
      container.innerHTML = `<div class="p-8 text-center text-slate-400 font-medium">Este participante no registra evaluación en Hito 3.</div>`;
      return;
    }
    const deltaSign = dataH.delta >= 0 ? '+' : '';
    const kpiH3 = document.createElement('div');
    kpiH3.className = 'grid grid-cols-4 gap-3 bg-slate-50/80 p-4 rounded-2xl border border-slate-100 text-center';
    kpiH3.innerHTML = `
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Puntaje PRE</div>
        <div class="text-xl font-extrabold text-slate-800 mt-0.5">${dataH.pre_score} / 5</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Puntaje POST</div>
        <div class="text-xl font-extrabold text-[#00a87e] mt-0.5">${dataH.post_score} / 5</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Diferencia</div>
        <div class="text-xl font-extrabold ${dataH.delta >= 0 ? 'text-[#00a87e]' : 'text-rose-600'} mt-0.5">${deltaSign}${dataH.delta} pts</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Condición</div>
        <span class="mt-1 inline-block font-bold px-2.5 py-0.5 rounded-full text-[11px] ${dataH.status === 'Mejoró' ? 'badge-improved' : (dataH.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained')}">${dataH.status}</span>
      </div>
    `;
    container.appendChild(kpiH3);

    const mcSection = document.createElement('div');
    mcSection.className = 'space-y-3';
    mcSection.innerHTML = `
      <h4 class="font-bold text-slate-900 text-sm">Preguntas Cerradas (Alternativas) - Hito 3</h4>
      ${buildHitoMCBlock('Hito 3: Los otros vecinos de Antofagasta', dataH, h3Meta, 'emerald')}
    `;
    container.appendChild(mcSection);

    const openSection = document.createElement('div');
    openSection.className = 'space-y-3';
    openSection.innerHTML = `
      <h4 class="font-bold text-slate-900 text-sm">Preguntas Abiertas (Cualitativo) - Hito 3</h4>
      ${buildHitoOpenBlock('Hito 3: Los otros vecinos de Antofagasta', dataH, 'emerald')}
    `;
    container.appendChild(openSection);

  } else if (mode === 'Hito 4') {
    const dataH = h4Data;
    if (!dataH) {
      container.innerHTML = `<div class="p-8 text-center text-slate-400 font-medium">Este participante no registra evaluación en Hito 4.</div>`;
      return;
    }
    const deltaSign = dataH.delta >= 0 ? '+' : '';
    const kpiH4 = document.createElement('div');
    kpiH4.className = 'grid grid-cols-4 gap-3 bg-slate-50/80 p-4 rounded-2xl border border-slate-100 text-center';
    kpiH4.innerHTML = `
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Puntaje PRE</div>
        <div class="text-xl font-extrabold text-slate-800 mt-0.5">${dataH.pre_score} / 5</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Puntaje POST</div>
        <div class="text-xl font-extrabold text-[#00a87e] mt-0.5">${dataH.post_score} / 5</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Diferencia</div>
        <div class="text-xl font-extrabold ${dataH.delta >= 0 ? 'text-[#00a87e]' : 'text-rose-600'} mt-0.5">${deltaSign}${dataH.delta} pts</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Condición</div>
        <span class="mt-1 inline-block font-bold px-2.5 py-0.5 rounded-full text-[11px] ${dataH.status === 'Mejoró' ? 'badge-improved' : (dataH.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained')}">${dataH.status}</span>
      </div>
    `;
    container.appendChild(kpiH4);

    const mcSection = document.createElement('div');
    mcSection.className = 'space-y-3';
    mcSection.innerHTML = `
      <h4 class="font-bold text-slate-900 text-sm">Preguntas Cerradas (Alternativas) - Hito 4</h4>
      ${buildHitoMCBlock('Hito 4: Proteger nuestro escudo invisible', dataH, h4Meta, 'indigo')}
    `;
    container.appendChild(mcSection);

    const openSection = document.createElement('div');
    openSection.className = 'space-y-3';
    openSection.innerHTML = `
      <h4 class="font-bold text-slate-900 text-sm">Preguntas Abiertas (Cualitativo) - Hito 4</h4>
      ${buildHitoOpenBlock('Hito 4: Proteger nuestro escudo invisible', dataH, 'indigo')}
    `;
    container.appendChild(openSection);
  }

  if (window.lucide) lucide.createIcons();
}

// Bloque de preguntas cerradas para un hito específico
function buildHitoMCBlock(title, hData, questionsMeta, theme) {
  if (!hData || !hData.pre_mc || !hData.post_mc) {
    return `
      <div class="p-4 rounded-2xl border border-dashed border-slate-200 bg-slate-50/50 text-slate-400 text-center text-xs">
        Sin registro de respuestas cerradas para este desafío.
      </div>
    `;
  }

  const qMeta = questionsMeta || [];
  const key = qMeta.slice(0, 5).map(q => (q.correct_key || '').toLowerCase());
  const themeBorder = theme === 'indigo' ? 'border-indigo-100 bg-indigo-50/20' : 'border-emerald-100 bg-emerald-50/20';
  const themeBadge = theme === 'indigo' ? 'bg-indigo-100 text-indigo-800' : 'bg-emerald-100 text-emerald-800';

  let itemsHtml = '';
  (hData.q_evo || []).forEach((evo, idx) => {
    if (idx >= qMeta.length) return;
    const isCorrectPost = (hData.post_mc[idx] || '').toLowerCase() === key[idx];
    const isCorrectPre = (hData.pre_mc[idx] || '').toLowerCase() === key[idx];

    let evoBadgeClass = 'bg-slate-100 text-slate-700';
    if (evo === 'Mejoró') evoBadgeClass = 'bg-[#00d09c]/20 text-[#00a87e] font-bold';
    else if (evo === 'Retrocedió') evoBadgeClass = 'bg-rose-100 text-rose-800 font-bold';
    else if (evo === 'Mantuvo Correcto') evoBadgeClass = 'bg-teal-50 text-teal-800 border border-teal-200 font-medium';

    itemsHtml += `
      <div class="p-3.5 rounded-2xl border ${isCorrectPost ? 'bg-emerald-50/40 border-emerald-200' : 'bg-rose-50/30 border-rose-200'} text-xs flex flex-wrap items-center justify-between gap-2">
        <div class="space-y-0.5 max-w-[75%]">
          <span class="font-bold text-slate-800">Q${idx + 1}: ${escapeHtml(qMeta[idx].title)}</span>
          <div class="text-[11px] text-slate-500 font-medium">
            PRE: <span class="font-bold ${isCorrectPre ? 'text-[#00a87e]' : 'text-slate-700'}">${(hData.pre_mc[idx] || 'N/A').toUpperCase()}</span> ➔ 
            POST: <span class="font-bold ${isCorrectPost ? 'text-[#00a87e]' : 'text-rose-600'}">${(hData.post_mc[idx] || 'N/A').toUpperCase()}</span>
            <span class="text-slate-400 font-normal ml-1">(Correcta: ${(key[idx] || '').toUpperCase()})</span>
          </div>
        </div>
        <div>
          <span class="text-[10px] px-3 py-1 rounded-full font-bold ${evoBadgeClass}">
            ${evo}
          </span>
        </div>
      </div>
    `;
  });

  return `
    <div class="p-4 rounded-2xl border ${themeBorder} space-y-3">
      <div class="flex items-center justify-between border-b border-slate-100 pb-2">
        <span class="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full ${themeBadge}">${title}</span>
        <div class="text-xs font-bold text-slate-700">
          Puntaje: <span class="text-[#00a87e] font-extrabold text-sm">${hData.post_score} / 5</span> 
          <span class="text-[11px] text-slate-400 font-normal">(PRE: ${hData.pre_score}/5 · ${hData.delta >= 0 ? '+' : ''}${hData.delta} pts)</span>
        </div>
      </div>
      <div class="space-y-2">
        ${itemsHtml}
      </div>
    </div>
  `;
}

// Bloque de preguntas abiertas para un hito específico
function buildHitoOpenBlock(title, hData, theme) {
  if (!hData || !hData.open_analysis || hData.open_analysis.length === 0) {
    return `
      <div class="p-4 rounded-2xl border border-dashed border-slate-200 bg-slate-50/50 text-slate-400 text-center text-xs">
        Sin registro de respuestas abiertas para este desafío.
      </div>
    `;
  }

  const themeBorder = theme === 'indigo' ? 'border-indigo-100 bg-indigo-50/20' : 'border-emerald-100 bg-emerald-50/20';
  const themeBadge = theme === 'indigo' ? 'bg-indigo-100 text-indigo-800' : 'bg-emerald-100 text-emerald-800';

  let itemsHtml = '';
  hData.open_analysis.forEach((oq, idx) => {
    let oBadge = 'badge-maintained';
    let oBadgeText = oq.status;

    if (oq.status === 'Mejoró') {
      oBadge = 'badge-improved font-bold';
      oBadgeText = 'Mejora Cualitativa ↗';
    } else if (oq.status === 'Retrocedió') {
      oBadge = 'badge-regressed font-bold';
      oBadgeText = 'Retroceso en Texto ↘';
    } else if (oq.status === 'Mantuvo') {
      oBadge = 'bg-blue-50 text-blue-700 border border-blue-200 font-semibold';
      oBadgeText = 'Mantuvo Respuesta =';
    } else if (oq.status === 'Sin Evidencia') {
      oBadge = 'bg-slate-100 text-slate-600 border border-slate-300 font-semibold';
      oBadgeText = 'Sin Evidencia (Revisar)';
    }

    const reviewTag = oq.requires_review ? 
      `<div class="text-[11px] text-amber-800 bg-amber-50 p-2.5 rounded-xl border border-amber-200 font-medium flex items-center gap-2">
        <i data-lucide="alert-triangle" class="w-4 h-4 text-amber-600 shrink-0"></i>
        <span>Elemento de evaluación no concluyente por falta de respuesta escrita. Se sugiere revisión individual con el profesor/a líder.</span>
      </div>` : '';

    itemsHtml += `
      <div class="p-4 rounded-2xl border border-slate-100 bg-white space-y-2.5 card-shadow">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <span class="font-extrabold text-slate-900 text-xs">P${idx + 1}: ${escapeHtml(oq.question)}</span>
          <span class="text-[10px] px-2.5 py-0.5 rounded-full ${oBadge}">${oBadgeText}</span>
        </div>
        <div class="text-[11px] text-slate-500 font-medium">
          Diagnóstico conceptual: <span class="text-slate-800 font-semibold">${escapeHtml(oq.detail)}</span>
        </div>
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5 text-xs">
          <div class="bg-slate-50 p-3 rounded-xl border border-slate-200">
            <span class="text-[10px] font-bold text-slate-400 block uppercase tracking-wide mb-0.5">PRE (Inicial):</span>
            <p class="italic text-slate-700">${oq.pre_text && oq.pre_text.trim() ? '"' + escapeHtml(oq.pre_text) + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta o no sabe)</span>'}</p>
          </div>
          <div class="bg-emerald-50/40 p-3 rounded-xl border border-emerald-200">
            <span class="text-[10px] font-bold text-[#00a87e] block uppercase tracking-wide mb-0.5">POST (En terreno):</span>
            <p class="font-medium text-slate-800">${oq.post_text && oq.post_text.trim() ? '"' + escapeHtml(oq.post_text) + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta o no sabe)</span>'}</p>
          </div>
        </div>
        ${reviewTag}
      </div>
    `;
  });

  return `
    <div class="p-4 rounded-2xl border ${themeBorder} space-y-3">
      <div class="flex items-center justify-between border-b border-slate-100 pb-2">
        <span class="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full ${themeBadge}">${title}</span>
        <span class="text-[11px] text-slate-400 font-medium">4 preguntas cualitativas</span>
      </div>
      <div class="space-y-3">
        ${itemsHtml}
      </div>
    </div>
  `;
}

function closeModal() {
  document.getElementById('participant-modal').classList.add('hidden');
}

// Generación y descarga de la Ficha Individual en formato PDF
function exportParticipantPDF() {
  const p = currentModalParticipant;
  if (!p) return;

  const btn = document.getElementById('btn-modal-pdf');
  const originalHtml = btn ? btn.innerHTML : '';
  if (btn) {
    btn.innerHTML = `
      <svg class="animate-spin -ml-1 mr-2 h-3.5 w-3.5 text-slate-900 inline" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
        <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
        <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path>
      </svg>
      <span>Generando PDF...</span>
    `;
    btn.disabled = true;
  }

  const h3Data = p.h3;
  const h4Data = p.h4;
  const h3Meta = (RAW_DATA.hitos['Hito 3'] && RAW_DATA.hitos['Hito 3'].questions) || [];
  const h4Meta = (RAW_DATA.hitos['Hito 4'] && RAW_DATA.hitos['Hito 4'].questions) || [];

  const printableDoc = document.createElement('div');
  printableDoc.id = 'pdf-printable-document';
  printableDoc.innerHTML = buildPDFDocumentHTML(p, h3Data, h4Data, currentModalMode, h3Meta, h4Meta);

  // Ubicar elemento para renderizado fiel
  printableDoc.style.position = 'fixed';
  printableDoc.style.left = '-9999px';
  printableDoc.style.top = '0';
  printableDoc.style.width = '794px'; // Formato A4 a 96dpi aprox
  printableDoc.style.backgroundColor = '#ffffff';
  document.body.appendChild(printableDoc);

  const cleanFileName = `Ficha_PAC_2026_${p.name.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]/g, '_')}.pdf`;

  if (typeof html2pdf !== 'undefined') {
    const opt = {
      margin: [10, 10, 10, 10],
      filename: cleanFileName,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, logging: false },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    html2pdf().set(opt).from(printableDoc).save().then(() => {
      if (btn) {
        btn.innerHTML = originalHtml;
        btn.disabled = false;
      }
      if (document.body.contains(printableDoc)) {
        document.body.removeChild(printableDoc);
      }
    }).catch(err => {
      console.error('Error al generar PDF con html2pdf:', err);
      if (btn) {
        btn.innerHTML = originalHtml;
        btn.disabled = false;
      }
      if (document.body.contains(printableDoc)) {
        document.body.removeChild(printableDoc);
      }
      alert('Se abrirá la opción de impresión del navegador para guardar en PDF.');
      window.print();
    });
  } else {
    if (btn) {
      btn.innerHTML = originalHtml;
      btn.disabled = false;
    }
    if (document.body.contains(printableDoc)) {
      document.body.removeChild(printableDoc);
    }
    window.print();
  }
}

// Construye la estructura HTML optimizada para impresión en PDF A4
function buildPDFDocumentHTML(p, h3Data, h4Data, mode, h3Meta, h4Meta) {
  const h3Pre = h3Data ? h3Data.pre_score : 0;
  const h3Post = h3Data ? h3Data.post_score : 0;
  const h4Pre = h4Data ? h4Data.pre_score : 0;
  const h4Post = h4Data ? h4Data.post_score : 0;
  const totPre = h3Pre + h4Pre;
  const totPost = h3Post + h4Post;
  const totDelta = totPost - totPre;
  const deltaSign = totDelta >= 0 ? '+' : '';

  let coverageText = 'Evaluación Individual';
  if (h3Data && h4Data) coverageText = 'Ambos Desafíos: Hito 3 y Hito 4 (Test-Retest Completo)';
  else if (h3Data) coverageText = 'Hito 3: Los otros vecinos de Antofagasta';
  else if (h4Data) coverageText = 'Hito 4: Proteger nuestro escudo invisible';

  // Filas para preguntas cerradas H3
  let h3MCRows = '';
  if (h3Data && h3Data.pre_mc && h3Data.post_mc) {
    const key3 = h3Meta.slice(0, 5).map(q => (q.correct_key || '').toUpperCase());
    h3Data.q_evo.forEach((evo, idx) => {
      if (idx >= h3Meta.length) return;
      const isPostCorr = (h3Data.post_mc[idx] || '').toUpperCase() === key3[idx];
      h3MCRows += `
        <tr>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; font-weight: bold; text-align: center;">Q${idx + 1}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0;">${escapeHtml(h3Meta[idx].title)}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold;">${(h3Data.pre_mc[idx] || 'N/A').toUpperCase()}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: ${isPostCorr ? '#00a87e' : '#e11d48'};">${(h3Data.post_mc[idx] || 'N/A').toUpperCase()}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #00a87e;">${key3[idx]}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; text-align: center; font-size: 10px;">${evo}</td>
        </tr>
      `;
    });
  }

  // Filas para preguntas cerradas H4
  let h4MCRows = '';
  if (h4Data && h4Data.pre_mc && h4Data.post_mc) {
    const key4 = h4Meta.slice(0, 5).map(q => (q.correct_key || '').toUpperCase());
    h4Data.q_evo.forEach((evo, idx) => {
      if (idx >= h4Meta.length) return;
      const isPostCorr = (h4Data.post_mc[idx] || '').toUpperCase() === key4[idx];
      h4MCRows += `
        <tr>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; font-weight: bold; text-align: center;">Q${idx + 1}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0;">${escapeHtml(h4Meta[idx].title)}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold;">${(h4Data.pre_mc[idx] || 'N/A').toUpperCase()}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: ${isPostCorr ? '#00a87e' : '#e11d48'};">${(h4Data.post_mc[idx] || 'N/A').toUpperCase()}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #00a87e;">${key4[idx]}</td>
          <td style="padding: 5px 6px; border: 1px solid #e2e8f0; text-align: center; font-size: 10px;">${evo}</td>
        </tr>
      `;
    });
  }

  // Tarjetas para preguntas abiertas H3
  let h3OpenCards = '';
  if (h3Data && h3Data.open_analysis) {
    h3Data.open_analysis.forEach((oq, idx) => {
      h3OpenCards += `
        <div style="margin-bottom: 8px; padding: 8px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #fafafa; page-break-inside: avoid;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
            <strong style="font-size: 11px; color: #0f172a;">P${idx + 1}: ${escapeHtml(oq.question)}</strong>
            <span style="font-size: 9px; font-weight: bold; padding: 2px 8px; border-radius: 12px; background-color: ${oq.status === 'Mejoró' ? '#e6faf5' : (oq.status === 'Retrocedió' ? '#fff1f2' : '#f1f5f9')}; color: ${oq.status === 'Mejoró' ? '#00a87e' : (oq.status === 'Retrocedió' ? '#e11d48' : '#475569')};">${oq.status}</span>
          </div>
          <div style="font-size: 10px; color: #64748b; margin-bottom: 6px;">Diagnóstico: <strong>${escapeHtml(oq.detail)}</strong></div>
          <table style="width: 100%; border-collapse: collapse; font-size: 10px;">
            <tr>
              <td style="width: 50%; padding: 6px; border: 1px solid #e2e8f0; background-color: #ffffff; vertical-align: top;">
                <span style="font-size: 9px; font-weight: bold; color: #94a3b8; display: block; text-transform: uppercase;">Respuesta PRE (Inicial):</span>
                <span style="font-style: italic; color: #334155;">${oq.pre_text && oq.pre_text.trim() ? '"' + escapeHtml(oq.pre_text) + '"' : '(Sin respuesta o no sabe)'}</span>
              </td>
              <td style="width: 50%; padding: 6px; border: 1px solid #e2e8f0; background-color: #f0fdf4; vertical-align: top;">
                <span style="font-size: 9px; font-weight: bold; color: #00a87e; display: block; text-transform: uppercase;">Respuesta POST (En terreno):</span>
                <span style="color: #0f172a; font-weight: 500;">${oq.post_text && oq.post_text.trim() ? '"' + escapeHtml(oq.post_text) + '"' : '(Sin respuesta o no sabe)'}</span>
              </td>
            </tr>
          </table>
          ${oq.requires_review ? `<div style="margin-top: 4px; font-size: 9px; color: #92400e; background-color: #fef3c7; padding: 4px 6px; border-radius: 4px;">⚠️ Elemento de evaluación no concluyente por falta de respuesta escrita. Se sugiere revisión individual con el profesor/a líder.</div>` : ''}
        </div>
      `;
    });
  }

  // Tarjetas para preguntas abiertas H4
  let h4OpenCards = '';
  if (h4Data && h4Data.open_analysis) {
    h4Data.open_analysis.forEach((oq, idx) => {
      h4OpenCards += `
        <div style="margin-bottom: 8px; padding: 8px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #fafafa; page-break-inside: avoid;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
            <strong style="font-size: 11px; color: #0f172a;">P${idx + 1}: ${escapeHtml(oq.question)}</strong>
            <span style="font-size: 9px; font-weight: bold; padding: 2px 8px; border-radius: 12px; background-color: ${oq.status === 'Mejoró' ? '#e6faf5' : (oq.status === 'Retrocedió' ? '#fff1f2' : '#f1f5f9')}; color: ${oq.status === 'Mejoró' ? '#00a87e' : (oq.status === 'Retrocedió' ? '#e11d48' : '#475569')};">${oq.status}</span>
          </div>
          <div style="font-size: 10px; color: #64748b; margin-bottom: 6px;">Diagnóstico: <strong>${escapeHtml(oq.detail)}</strong></div>
          <table style="width: 100%; border-collapse: collapse; font-size: 10px;">
            <tr>
              <td style="width: 50%; padding: 6px; border: 1px solid #e2e8f0; background-color: #ffffff; vertical-align: top;">
                <span style="font-size: 9px; font-weight: bold; color: #94a3b8; display: block; text-transform: uppercase;">Respuesta PRE (Inicial):</span>
                <span style="font-style: italic; color: #334155;">${oq.pre_text && oq.pre_text.trim() ? '"' + escapeHtml(oq.pre_text) + '"' : '(Sin respuesta o no sabe)'}</span>
              </td>
              <td style="width: 50%; padding: 6px; border: 1px solid #e2e8f0; background-color: #f0fdf4; vertical-align: top;">
                <span style="font-size: 9px; font-weight: bold; color: #00a87e; display: block; text-transform: uppercase;">Respuesta POST (En terreno):</span>
                <span style="color: #0f172a; font-weight: 500;">${oq.post_text && oq.post_text.trim() ? '"' + escapeHtml(oq.post_text) + '"' : '(Sin respuesta o no sabe)'}</span>
              </td>
            </tr>
          </table>
          ${oq.requires_review ? `<div style="margin-top: 4px; font-size: 9px; color: #92400e; background-color: #fef3c7; padding: 4px 6px; border-radius: 4px;">⚠️ Elemento de evaluación no concluyente por falta de respuesta escrita. Se sugiere revisión individual con el profesor/a líder.</div>` : ''}
        </div>
      `;
    });
  }

  return `
    <div style="font-family: Arial, Helvetica, sans-serif; color: #1e293b; padding: 20px; line-height: 1.4; background-color: #ffffff;">
      
      <!-- Encabezado con Logos Oficiales -->
      <table style="width: 100%; border-bottom: 2px solid #00d09c; padding-bottom: 12px; margin-bottom: 16px;">
        <tr>
          <td style="width: 70px; vertical-align: middle;">
            <img src="logo.png" style="height: 48px; width: auto; object-fit: contain;" alt="PAC 2026">
          </td>
          <td style="vertical-align: middle; padding-left: 10px;">
            <div style="font-size: 13px; font-weight: bold; text-transform: uppercase; color: #0f172a; letter-spacing: 0.5px;">Programa de Acción Climática 2026 · Línea Verde</div>
            <div style="font-size: 17px; font-weight: 900; color: #00a87e; margin-top: 2px;">FICHA INDIVIDUAL DE EVALUACIÓN TEST-RETEST</div>
            <div style="font-size: 10px; color: #64748b;">Antofagasta: Laboratorio Vivo hacia el 2050 · Reporte de Aprendizajes</div>
          </td>
          <td style="width: 120px; text-align: right; vertical-align: middle;">
            <img src="logo_verdical.png" style="height: 38px; width: auto; object-fit: contain;" alt="Grupo Verdical">
          </td>
        </tr>
      </table>

      <!-- Ficha de Identificación del Participante -->
      <table style="width: 100%; border-collapse: collapse; margin-bottom: 16px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;">
        <tr>
          <td style="padding: 8px 12px; width: 50%; vertical-align: top;">
            <div style="font-size: 9px; text-transform: uppercase; font-weight: bold; color: #94a3b8;">Participante:</div>
            <div style="font-size: 15px; font-weight: 900; color: #0f172a;">${escapeHtml(p.name)}</div>
            <div style="font-size: 11px; color: #475569; margin-top: 2px;">Establecimiento: <strong>${escapeHtml(p.school)}</strong></div>
          </td>
          <td style="padding: 8px 12px; width: 50%; vertical-align: top; border-left: 1px solid #e2e8f0;">
            <div style="font-size: 9px; text-transform: uppercase; font-weight: bold; color: #94a3b8;">Detalles de Participación:</div>
            <div style="font-size: 11px; color: #334155; margin-top: 2px;">Género: <strong>${escapeHtml(p.gender)}</strong> · Rol: <strong>${escapeHtml(p.role || 'Estudiante')}</strong></div>
            <div style="font-size: 11px; color: #00a87e; font-weight: bold; margin-top: 2px;">${coverageText}</div>
            <div style="font-size: 11px; color: #64748b; margin-top: 2px;">Diagnóstico: <strong style="color: #0f172a;">${escapeHtml(p.real_status || p.status)}</strong></div>
          </td>
        </tr>
      </table>

      <!-- Tabla de Rendimiento Cuantitativo -->
      <div style="margin-bottom: 16px;">
        <div style="font-size: 12px; font-weight: bold; text-transform: uppercase; color: #0f172a; margin-bottom: 6px; border-left: 3px solid #00d09c; padding-left: 6px;">1. Resumen Cuantitativo de Rendimiento (Puntajes PRE vs. POST)</div>
        <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
          <thead>
            <tr style="background-color: #f1f5f9; text-align: left;">
              <th style="padding: 6px 8px; border: 1px solid #cbd5e1;">Desafío / Hito Evaluado</th>
              <th style="padding: 6px 8px; border: 1px solid #cbd5e1; text-align: center;">Puntaje PRE</th>
              <th style="padding: 6px 8px; border: 1px solid #cbd5e1; text-align: center;">Puntaje POST</th>
              <th style="padding: 6px 8px; border: 1px solid #cbd5e1; text-align: center;">Diferencia (Delta)</th>
              <th style="padding: 6px 8px; border: 1px solid #cbd5e1; text-align: center;">Condición de Aprendizaje</th>
            </tr>
          </thead>
          <tbody>
            ${h3Data ? `
              <tr>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; font-weight: bold;">Hito 3: Los otros vecinos de Antofagasta (Biodiversidad)</td>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; text-align: center;">${h3Data.pre_score} / 5</td>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #00a87e;">${h3Data.post_score} / 5</td>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: ${h3Data.delta >= 0 ? '#00a87e' : '#e11d48'};">${h3Data.delta >= 0 ? '+' : ''}${h3Data.delta} pts</td>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; text-align: center;">${h3Data.status}</td>
              </tr>
            ` : ''}
            ${h4Data ? `
              <tr>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; font-weight: bold;">Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos)</td>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; text-align: center;">${h4Data.pre_score} / 5</td>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #00a87e;">${h4Data.post_score} / 5</td>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: ${h4Data.delta >= 0 ? '#00a87e' : '#e11d48'};">${h4Data.delta >= 0 ? '+' : ''}${h4Data.delta} pts</td>
                <td style="padding: 6px 8px; border: 1px solid #e2e8f0; text-align: center;">${h4Data.status}</td>
              </tr>
            ` : ''}
            ${h3Data && h4Data ? `
              <tr style="background-color: #ecfdf5; font-weight: bold;">
                <td style="padding: 6px 8px; border: 1px solid #cbd5e1; color: #065f46;">Consolidado Combinado (10 preguntas cerradas)</td>
                <td style="padding: 6px 8px; border: 1px solid #cbd5e1; text-align: center;">${totPre} / 10 pts</td>
                <td style="padding: 6px 8px; border: 1px solid #cbd5e1; text-align: center; color: #00a87e;">${totPost} / 10 pts</td>
                <td style="padding: 6px 8px; border: 1px solid #cbd5e1; text-align: center; color: ${totDelta >= 0 ? '#00a87e' : '#e11d48'};">${deltaSign}${totDelta} pts</td>
                <td style="padding: 6px 8px; border: 1px solid #cbd5e1; text-align: center; color: #065f46;">${p.status}</td>
              </tr>
            ` : ''}
          </tbody>
        </table>
      </div>

      <!-- Sección 2: Desglose Preguntas Cerradas -->
      <div style="margin-bottom: 16px;">
        <div style="font-size: 12px; font-weight: bold; text-transform: uppercase; color: #0f172a; margin-bottom: 6px; border-left: 3px solid #00d09c; padding-left: 6px;">2. Desglose de Preguntas Cerradas (Alternativas)</div>
        
        ${h3Data ? `
          <div style="margin-bottom: 10px;">
            <div style="font-size: 11px; font-weight: bold; color: #065f46; margin-bottom: 4px;">Hito 3: Los otros vecinos de Antofagasta (Biodiversidad y Humedales) · Puntaje: ${h3Data.post_score} / 5</div>
            <table style="width: 100%; border-collapse: collapse; font-size: 10px;">
              <thead>
                <tr style="background-color: #f8fafc;">
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 30px; text-align: center;">#</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1;">Pregunta Evaluada</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 45px; text-align: center;">PRE</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 45px; text-align: center;">POST</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 50px; text-align: center;">Correcta</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 85px; text-align: center;">Evolución</th>
                </tr>
              </thead>
              <tbody>
                ${h3MCRows}
              </tbody>
            </table>
          </div>
        ` : ''}

        ${h4Data ? `
          <div style="margin-bottom: 10px;">
            <div style="font-size: 11px; font-weight: bold; color: #3730a3; margin-bottom: 4px;">Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos Limpios) · Puntaje: ${h4Data.post_score} / 5</div>
            <table style="width: 100%; border-collapse: collapse; font-size: 10px;">
              <thead>
                <tr style="background-color: #f8fafc;">
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 30px; text-align: center;">#</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1;">Pregunta Evaluada</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 45px; text-align: center;">PRE</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 45px; text-align: center;">POST</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 50px; text-align: center;">Correcta</th>
                  <th style="padding: 4px 6px; border: 1px solid #cbd5e1; width: 85px; text-align: center;">Evolución</th>
                </tr>
              </thead>
              <tbody>
                ${h4MCRows}
              </tbody>
            </table>
          </div>
        ` : ''}
      </div>

      <!-- Sección 3: Desglose Preguntas Abiertas (Cualitativo) -->
      <div style="margin-bottom: 16px;">
        <div style="font-size: 12px; font-weight: bold; text-transform: uppercase; color: #0f172a; margin-bottom: 6px; border-left: 3px solid #00d09c; padding-left: 6px;">3. Desglose de Preguntas Abiertas (Análisis Cualitativo de Apropiación)</div>
        
        ${h3Data ? `
          <div style="margin-bottom: 12px;">
            <div style="font-size: 11px; font-weight: bold; color: #065f46; margin-bottom: 6px;">Hito 3: Los otros vecinos de Antofagasta (Biodiversidad y Ecosistemas)</div>
            ${h3OpenCards}
          </div>
        ` : ''}

        ${h4Data ? `
          <div style="margin-bottom: 12px;">
            <div style="font-size: 11px; font-weight: bold; color: #3730a3; margin-bottom: 6px;">Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos Limpios)</div>
            ${h4OpenCards}
          </div>
        ` : ''}
      </div>

      <!-- Sección 4: Orientaciones Pedagógicas -->
      <div style="margin-bottom: 20px; padding: 10px 12px; background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; page-break-inside: avoid;">
        <div style="font-size: 11px; font-weight: bold; color: #166534; margin-bottom: 4px;">Orientación para el Profesor/a Líder:</div>
        <p style="font-size: 10px; color: #14532d; margin: 0; line-height: 1.45;">
          ${p.status === 'Mejoró' ? 
            'El participante evidenció un avance significativo en la apropiación de conceptos y terminología científica tanto en preguntas de selección múltiple como en sus explicaciones escritas. Se recomienda canalizar esta motivación hacia la fase de ideación de soluciones climáticas escolares.' : 
            (p.status === 'Mantuvo' ? 
              'El participante sostuvo un nivel de respuesta coherente y estructurado entre el PRE y POST test. Se sugiere afianzar los conceptos clave en un plenario grupal de 15 minutos previo a la formulación del proyecto de acción climática.' : 
              'Se observaron elementos no concluyentes o respuestas breves asociadas a la aplicación inmediata en terreno. Se recomienda una breve conversación formativa con el docente líder para verificar la asimilación del contenido.')}
        </p>
      </div>

      <!-- Pie de Página Institucional -->
      <table style="width: 100%; border-top: 1px solid #e2e8f0; padding-top: 8px; font-size: 9px; color: #94a3b8;">
        <tr>
          <td>Herramienta desarrollada por <strong>rodolfo.tapia@verdical.cl</strong> para <strong>Grupo Verdical SPA</strong></td>
          <td style="text-align: right;">Programa de Acción Climática (PAC 2026) · Metodología VERDICAL · Antofagasta, Chile</td>
        </tr>
      </table>

    </div>
  `;
}

// Exportar tabla a CSV
function exportSchoolsTable() {
  const hitoKey = (CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global') ? 'Global' : CURRENT_HITO;
  const schools = RAW_DATA.hitos[hitoKey].schools;

  let csv = 'Establecimiento,Dependencia,N_Participantes,Puntaje_PRE,Puntaje_POST,Delta,Pct_Mejora,Lider_1,Email_1,Telefono_1\n';
  schools.forEach(s => {
    const pre = s.pre_avg !== null ? s.pre_avg : 'Sin datos';
    const post = s.post_avg !== null ? s.post_avg : 'Sin datos';
    const del = s.has_data !== false ? s.delta_avg : 'Sin datos';
    csv += `"${s.school}","${s.dependencia}",${s.n},"${pre}","${post}","${del}","${s.pct_mejora}%","${s.lider1}","${s.email1}","${s.tel1}"\n`;
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', `evaluacion_establecimientos_${hitoKey}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
