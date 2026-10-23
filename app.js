/**
 * Lógica principal del Dashboard Interactivo de Evaluación Línea Verde - PAC 2026
 * Con evaluación cualitativa rigurosa y normalización canónica de datos
 */

let RAW_DATA = null;
let CURRENT_HITO = 'Global'; // Default a Resumen Global
let ACTIVE_TAB = 'resumen';
let CHARTS = {};

const GOOGLE_SPREADSHEET_ID = '1tI7030zyCF16tAQ_HvlUSzrFkZ-aWSjVABZDInmHejY';

// Lista blanca estricta y canónica de las escuelas de la Línea Verde
const LINEA_VERDE_SCHOOLS = {
  'Hito 3': [
    'Escuela La Bandera',
    'Escuela Presbiteriana',
    'Escuela Héroes de la Concepción',
    'Escuela Rómulo Peña'
  ],
  'Hito 4': [
    'Liceo Bicentenario Andrés Sabella',
    'Greenhill School',
    'Colegio Bet-el'
  ],
  'Global': [
    'Escuela La Bandera',
    'Escuela Presbiteriana',
    'Escuela Héroes de la Concepción',
    'Escuela Rómulo Peña',
    'Liceo Bicentenario Andrés Sabella',
    'Greenhill School',
    'Colegio Bet-el'
  ]
};

const ALL_LINEA_VERDE_SCHOOLS = new Set(LINEA_VERDE_SCHOOLS['Global']);

// Inicialización resiliente al cargar el DOM o script
function initApp() {
  const filtersContainer = document.getElementById('filters-container');
  if (filtersContainer) {
    if (ACTIVE_TAB === 'resumen') {
      filtersContainer.classList.remove('hidden');
      filtersContainer.classList.add('flex');
    } else {
      filtersContainer.classList.add('hidden');
      filtersContainer.classList.remove('flex');
    }
  }

  if (window.lucide) {
    lucide.createIcons();
  }
  loadDashboardData();
  setupEventListeners();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}

function setupEventListeners() {
  window.addEventListener('hashchange', () => {
    if (window.location.hash === '#biblioteca') {
      switchTab('biblioteca');
    }
  });
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

// Sincronización en tiempo real desde Google Drive
async function loadFromGoogleDrive() {
  const sheets = ['Encuestas Hito 3_Verde', 'Encuestas Hito 4_Verde', 'Participantes 2026', 'Resumen Líderes'];
  const wb = XLSX.utils.book_new();

  await Promise.all(sheets.map(async (sheetName) => {
    const url = `https://docs.google.com/spreadsheets/d/${GOOGLE_SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}&t=${Date.now()}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status} al leer ${sheetName}`);
    const csv = await res.text();
    const tempWb = XLSX.read(csv, { type: 'string' });
    const ws = tempWb.Sheets[tempWb.SheetNames[0]];
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }));

  processWorkbookInBrowser(wb, 'Google Drive (En vivo)');
}

// Cargar datos: 1° Google Drive (en vivo), 2° API local, 3° data.json de respaldo
async function loadDashboardData(forceReload = false) {
  const syncStatus = document.getElementById('sync-status');
  const syncText = document.getElementById('sync-text');
  if (syncText) syncText.innerText = 'Cargando datos desde Google Drive...';
  if (syncStatus) syncStatus.className = 'flex items-center gap-2 text-xs text-slate-600 bg-slate-100 px-3.5 py-1.5 rounded-full border border-slate-200';

  // 1. Intentar cargar en tiempo real desde Google Drive
  try {
    await loadFromGoogleDrive();
    return;
  } catch (errDrive) {
    console.warn('No se pudo sincronizar directamente con Google Drive, intentando fuente local...', errDrive);
  }

  // 2. Intentar API local (si está corriendo server.py)
  try {
    const response = await fetch(`/api/data?t=${Date.now()}`);
    if (!response.ok) throw new Error('API local no respondió');
    RAW_DATA = await response.json();
    if (syncText) syncText.innerText = 'Conectado a Excel (Local en vivo)';
    if (syncStatus) syncStatus.className = 'flex items-center gap-2 text-xs text-emerald-800 bg-emerald-50 px-3.5 py-1.5 rounded-full border border-emerald-200';
    ensureGlobalHito(RAW_DATA);
    populateFilterDropdowns();
    updateUI();
    return;
  } catch (errApi) {
    console.warn('Fallo API local, cargando data.json de respaldo...', errApi);
  }

  // 3. Fallback estático a data.json
  try {
    const respStatic = await fetch(`./data.json?t=${Date.now()}`);
    if (!respStatic.ok) throw new Error('data.json no disponible');
    RAW_DATA = await respStatic.json();
    if (syncText) syncText.innerText = 'Datos de respaldo cargados (data.json)';
    if (syncStatus) syncStatus.className = 'flex items-center gap-2 text-xs text-emerald-800 bg-emerald-50 px-3.5 py-1.5 rounded-full border border-emerald-200';
    ensureGlobalHito(RAW_DATA);
    populateFilterDropdowns();
    updateUI();
  } catch (e2) {
    console.error('Error crítico al cargar datos:', e2);
    if (syncText) syncText.innerText = 'Error al cargar datos.';
    if (syncStatus) syncStatus.className = 'flex items-center gap-2 text-xs text-rose-800 bg-rose-50 px-3.5 py-1.5 rounded-full border border-rose-200';
  }
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
    if (!baseP || !ALL_LINEA_VERDE_SCHOOLS.has(baseP.school)) return;

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
        real_status: (p3.status === 'Mejoró' || p4.status === 'Mejoró') ? 'Avance en Aprendizajes' : (statusTot === 'Mantuvo' ? 'Aprendizaje Consolidado' : 'Requiere Refuerzo'),
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

  // Filtrar escuelas en todos los hitos para asegurar solo Línea Verde
  for (const hname of Object.keys(data.hitos)) {
    if (data.hitos[hname].schools) {
      data.hitos[hname].schools = data.hitos[hname].schools.filter(s => ALL_LINEA_VERDE_SCHOOLS.has(s.school));
    }
  }

  const uniqueList = getUniqueParticipants(data);

  if (data.hitos['Global']) {
    data.hitos['Global'].unique_participants = uniqueList;
    data.hitos['Global'].participants = uniqueList;
    if (data.hitos['Global'].schools) {
      data.hitos['Global'].schools = data.hitos['Global'].schools.filter(s => ALL_LINEA_VERDE_SCHOOLS.has(s.school));
    }
    return;
  }

  const allPaired = [];
  ['Hito 3', 'Hito 4'].forEach(hname => {
    if (data.hitos[hname] && data.hitos[hname].participants) {
      data.hitos[hname].participants.forEach(p => {
        if (ALL_LINEA_VERDE_SCHOOLS.has(p.school)) {
          const pCopy = Object.assign({}, p, { hito: p.hito || hname });
          allPaired.push(pCopy);
        }
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
  return n;
}

const NON_INFORMATIVE = [
  'no se', 'nada', 'ninguna', 'no me acuerdo', 'nose', 'no responde', 
  'no se como decirlo', 'no se como puedo decirlo', 'no se explicarlo', 
  'no se que poner', 'no sabria decir', 'no responder', 'ninguno', 'no lo se'
];

function isNonInformative(text) {
  if (!text) return true;
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
      detail: `Consolidó respuesta informada (${lblPost})`,
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
        if (!ALL_LINEA_VERDE_SCHOOLS.has(sch)) continue;
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
        if (!ALL_LINEA_VERDE_SCHOOLS.has(sch)) continue;
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
        if (!ALL_LINEA_VERDE_SCHOOLS.has(school)) continue;
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
          let realStatus = mcStatus === 'Retrocedió' ? 'Requiere Refuerzo Conceptual' : 'Aprendizaje Consolidado';
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
        if (!ALL_LINEA_VERDE_SCHOOLS.has(p.school)) return;
        if (!bySchool[p.school]) bySchool[p.school] = [];
        bySchool[p.school].push(p);
      });

      const schoolRankings = Object.entries(bySchool)
        .filter(([sname]) => ALL_LINEA_VERDE_SCHOOLS.has(sname))
        .map(([sname, plist]) => {
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
            dependencia: meta.dependencia || (sname.includes('Escuela') || sname.includes('Liceo') ? 'Municipal' : 'Particular Subvencionado'),
            certificacion: meta.certificacion || 'No',
            has_data: true
          };
        });

      // Incluir colegios de la Línea Verde de este hito que no registren datos (ej. Bet-el en Hito 4)
      const targetSchoolsHito = LINEA_VERDE_SCHOOLS[hitoName] || [];
      targetSchoolsHito.forEach(targetSch => {
        if (!bySchool[targetSch]) {
          const bMeta = schools[targetSch] || {};
          schoolRankings.push({
            school: targetSch,
            n: 0,
            pre_avg: null,
            post_avg: null,
            delta_avg: 0.0,
            status_counts: { 'Mejoró': 0, 'Mantuvo': 0, 'Retrocedió': 0 },
            pct_mejora: 0.0,
            lider1: bMeta.lider1 || (targetSch.includes('Bet-el') ? 'Pamela Pizarro Juica' : ''),
            email1: bMeta.email1 || (targetSch.includes('Bet-el') ? 'pizarropamela2015@gmail.com' : ''),
            tel1: bMeta.tel1 || (targetSch.includes('Bet-el') ? '56995779316' : ''),
            lider2: bMeta.lider2 || '',
            email2: bMeta.email2 || '',
            tel2: bMeta.tel2 || '',
            dependencia: bMeta.dependencia || (targetSch.includes('Escuela') || targetSch.includes('Liceo') ? 'Municipal' : 'Particular Subvencionado'),
            certificacion: bMeta.certificacion || 'No',
            has_data: false,
            note: 'Sin datos registrados en este hito'
          });
        }
      });

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
    if (syncText) {
      syncText.innerText = fileName.includes('Google Drive') ? 'Conectado a Google Drive (En vivo)' : `Cálculo autónomo: ${fileName}`;
    }
    if (syncStatus) {
      syncStatus.className = 'flex items-center gap-2 text-xs text-emerald-800 bg-emerald-50 px-3.5 py-1.5 rounded-full border border-emerald-200';
    }

    populateFilterDropdowns();
    updateUI();
  } catch (err) {
    console.error('Error parseando Excel en navegador:', err);
    if (!fileName.includes('Google Drive')) {
      alert('Error al procesar el archivo Excel en el navegador: ' + err.message);
    }
    throw err;
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
        if (s.has_data !== false && ALL_LINEA_VERDE_SCHOOLS.has(s.school)) schoolsSet.add(s.school);
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

// Estado de sub-pestaña para Evaluación de Contenidos (cerradas vs abiertas)
let PREGUNTAS_SUBTAB = 'cerradas';

function setPreguntasSubtab(subtab) {
  PREGUNTAS_SUBTAB = subtab;
  const secCerradas = document.getElementById('sec-preguntas-cerradas');
  const secAbiertas = document.getElementById('sec-preguntas-abiertas');
  const btnCerradas = document.getElementById('btn-subtab-cerradas');
  const btnAbiertas = document.getElementById('btn-subtab-abiertas');

  if (subtab === 'cerradas') {
    if (secCerradas) secCerradas.classList.remove('hidden');
    if (secAbiertas) secAbiertas.classList.add('hidden');
    if (btnCerradas) {
      btnCerradas.className = 'px-4 py-1.5 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm transition';
    }
    if (btnAbiertas) {
      btnAbiertas.className = 'px-4 py-1.5 rounded-full text-xs font-medium text-slate-600 hover:text-slate-900 transition';
    }
  } else {
    if (secCerradas) secCerradas.classList.add('hidden');
    if (secAbiertas) secAbiertas.classList.remove('hidden');
    if (btnCerradas) {
      btnCerradas.className = 'px-4 py-1.5 rounded-full text-xs font-medium text-slate-600 hover:text-slate-900 transition';
    }
    if (btnAbiertas) {
      btnAbiertas.className = 'px-4 py-1.5 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm transition';
    }
    renderQualitativeTab();
  }
  if (window.lucide) lucide.createIcons();
}

// Control de seguridad con contraseña para Fichas Individuales (PIN: 1234)
function isFichasUnlocked() {
  return sessionStorage.getItem('fichas_unlocked') === 'true';
}

function verifyFichasPassword() {
  const pinInput = document.getElementById('fichas-pin-input');
  const pinErr = document.getElementById('fichas-pin-error');
  const val = pinInput ? pinInput.value.trim() : '';

  if (val === '1234') {
    sessionStorage.setItem('fichas_unlocked', 'true');
    if (pinErr) pinErr.classList.add('hidden');
    if (pinInput) pinInput.value = '';
    renderParticipantsTab();
  } else {
    if (pinErr) pinErr.classList.remove('hidden');
    if (pinInput) {
      pinInput.value = '';
      pinInput.focus();
    }
  }
  if (window.lucide) lucide.createIcons();
}

function lockFichas() {
  sessionStorage.removeItem('fichas_unlocked');
  renderParticipantsTab();
  if (window.lucide) lucide.createIcons();
}

// Cambiar pestaña activa
function switchTab(tabId) {
  ACTIVE_TAB = tabId;
  document.querySelectorAll('.tab-view').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.nav-tab').forEach(el => {
    el.className = 'nav-tab flex items-center gap-2 px-4 py-2 rounded-full text-xs font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-100 whitespace-nowrap transition';
  });

  const targetView = document.getElementById(`view-${tabId}`);
  const targetBtn = document.getElementById(`tab-btn-${tabId}`);
  const headerBiblioBtn = document.getElementById('btn-header-biblioteca');
  if (headerBiblioBtn) {
    if (tabId === 'biblioteca') {
      headerBiblioBtn.className = 'inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-black text-slate-950 bg-[#88C23F] border border-[#88C23F] rounded-full transition shadow-sm';
    } else {
      headerBiblioBtn.className = 'inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-bold text-slate-800 bg-[#88C23F]/25 hover:bg-[#88C23F]/40 border border-[#88C23F]/60 rounded-full transition shadow-sm';
    }
  }
  if (targetView) targetView.classList.remove('hidden');
  if (targetBtn) {
    targetBtn.className = 'nav-tab active-tab flex items-center gap-2 px-4 py-2 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm whitespace-nowrap transition';
  }

  // Visibilidad del filtro: solo visible en la pestaña de resumen ejecutivo
  const filtersContainer = document.getElementById('filters-container');
  if (filtersContainer) {
    if (tabId === 'resumen') {
      filtersContainer.classList.remove('hidden');
      filtersContainer.classList.add('flex');
    } else {
      filtersContainer.classList.add('hidden');
      filtersContainer.classList.remove('flex');
    }
  }

  // En todas las pestañas la vista por defecto debe ser con la información de Resumen Global
  setHito('Global');

  if (tabId === 'preguntas') {
    setPreguntasSubtab(PREGUNTAS_SUBTAB);
  }

  if (tabId === 'participantes' && !isFichasUnlocked()) {
    setTimeout(() => {
      const pinInput = document.getElementById('fichas-pin-input');
      if (pinInput) pinInput.focus();
    }, 100);
  }

  if (window.lucide) lucide.createIcons();
}

// Cambiar Hito seleccionado
function setHito(hito) {
  CURRENT_HITO = hito;

  const btn4 = document.getElementById('btn-hito-4');
  const btn3 = document.getElementById('btn-hito-3');
  const btnGlobal = document.getElementById('btn-hito-global') || document.getElementById('btn-hito-ambos');

  [btn4, btn3, btnGlobal].forEach(btn => {
    if (btn) btn.className = 'px-4 py-1.5 rounded-full text-xs font-medium text-slate-600 hover:text-slate-900 transition';
  });

  const activeBtn = hito === 'Hito 4' ? btn4 : (hito === 'Hito 3' ? btn3 : btnGlobal);
  if (activeBtn) {
    if (hito === 'Hito 4') {
      activeBtn.className = 'px-4 py-1.5 rounded-full text-xs font-bold bg-[#00ADBB] text-white shadow-sm transition';
    } else if (hito === 'Hito 3') {
      activeBtn.className = 'px-4 py-1.5 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm transition';
    } else {
      activeBtn.className = 'px-4 py-1.5 rounded-full text-xs font-bold bg-[#00ADBB] text-white shadow-sm transition';
    }
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
  const hitoData = RAW_DATA.hitos[hitoKey] || RAW_DATA.hitos['Global'] || RAW_DATA.hitos['Hito 4'];
  if (!hitoData) return [];

  let list = hitoData.participants || [];

  // Los filtros de colegio, género, evolución y búsqueda aplican exclusivamente en la pestaña de resumen
  if (ACTIVE_TAB !== 'resumen') {
    return list;
  }

  const schoolEl = document.getElementById('filter-school');
  const genderEl = document.getElementById('filter-gender');
  const statusEl = document.getElementById('filter-status');
  const searchEl = document.getElementById('filter-search');

  const schoolFilter = schoolEl ? schoolEl.value : 'ALL';
  const genderFilter = genderEl ? genderEl.value : 'ALL';
  const statusFilter = statusEl ? statusEl.value : 'ALL';
  const searchFilter = searchEl ? searchEl.value.toLowerCase().trim() : '';

  if (schoolFilter !== 'ALL') {
    list = list.filter(p => p.school === schoolFilter);
  }

  if (genderFilter !== 'ALL') {
    list = list.filter(p => p.gender === genderFilter);
  }

  if (statusFilter !== 'ALL') {
    list = list.filter(p => p.status === statusFilter || 
      (statusFilter === 'Consolidó' && p.status === 'Mantuvo') || 
      (statusFilter === 'Mantuvo' && p.status === 'Mantuvo') ||
      (statusFilter === 'Requiere refuerzo' && p.status === 'Retrocedió') ||
      (statusFilter === 'Retrocedió' && p.status === 'Retrocedió'));
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
  renderClimatonTab();
  renderQuestionsTab();
  renderQualitativeTab();
  renderPedagogicalTab();
  renderParticipantsTab();

  if (window.lucide) lucide.createIcons();
}

function updateBanner() {
  const bannerEl = document.getElementById('banner-desafio');
  const tagEl = bannerEl ? bannerEl.querySelector('span') : null;
  const titleEl = document.getElementById('banner-title');
  const descEl = document.getElementById('banner-desc');
  const hakeEl = document.getElementById('kpi-hake-val');

  if (CURRENT_HITO === 'Hito 4') {
    titleEl.innerText = 'Proteger nuestro escudo invisible (Hito 4)';
    descEl.innerText = 'Evaluación del aprendizaje sobre la atmósfera como regulador climático, la contaminación lumínica en Antofagasta, la protección del cielo oscuro para la astronomía mundial y la fauna costera.';
    if (tagEl) {
      tagEl.className = 'text-xs uppercase tracking-wider text-[#00ADBB] font-bold';
    }
    if (bannerEl) {
      bannerEl.className = 'p-6 rounded-3xl bg-white border border-slate-200 text-slate-900 card-shadow flex flex-wrap items-center justify-between gap-4 border-l-4 border-l-[#00ADBB]';
    }
    const h4 = RAW_DATA.hitos['Hito 4'];
    if (h4) hakeEl.innerText = `g = +${h4.hake_gain}`;
  } else if (CURRENT_HITO === 'Hito 3') {
    titleEl.innerText = 'Los otros vecinos de Antofagasta (Hito 3)';
    descEl.innerText = 'Evaluación del aprendizaje sobre biodiversidad del desierto costero, tramas tróficas, fauna y flora aledaña a la ciudad de Antofagasta (Gaviotín Chico, descomponedores) y el humedal La Chimba.';
    if (tagEl) {
      tagEl.className = 'text-xs uppercase tracking-wider text-[#52801b] font-bold';
    }
    if (bannerEl) {
      bannerEl.className = 'p-6 rounded-3xl bg-white border border-slate-200 text-slate-900 card-shadow flex flex-wrap items-center justify-between gap-4 border-l-4 border-l-[#88C23F]';
    }
    const h3 = RAW_DATA.hitos['Hito 3'];
    if (h3) hakeEl.innerText = `g = +${h3.hake_gain}`;
  } else {
    titleEl.innerText = 'Resumen Global del Programa (Hito 3 y Hito 4)';
    descEl.innerText = 'Consolidado general de aprendizajes en la Línea Verde del Programa de Acción Climática (PAC 2026), integrando tramas tróficas, biodiversidad aledaña a Antofagasta, atmósfera y preservación de cielos limpios.';
    if (tagEl) {
      tagEl.className = 'text-xs uppercase tracking-wider text-[#00737d] font-bold';
    }
    if (bannerEl) {
      bannerEl.className = 'p-6 rounded-3xl bg-white border border-slate-200 text-slate-900 card-shadow flex flex-wrap items-center justify-between gap-4 border-l-4 border-l-[#00ADBB]';
    }
    const glob = RAW_DATA.hitos['Global'];
    if (glob) hakeEl.innerText = `g = +${glob.hake_gain}`;
    else hakeEl.innerText = 'g = +0.214';
  }

  const sampleEl = document.getElementById('banner-sample-text');
  if (sampleEl) {
    if (CURRENT_HITO === 'Hito 4') {
      sampleEl.innerText = 'Muestra: 61 evaluaciones pareadas (Hito 4)';
    } else if (CURRENT_HITO === 'Hito 3') {
      sampleEl.innerText = 'Muestra: 66 evaluaciones pareadas (Hito 3)';
    } else {
      sampleEl.innerText = 'Muestra: 127 evaluaciones pareadas (83 estudiantes únicos)';
    }
  }
}

function updateKPIs() {
  const filtered = getFilteredParticipants();
  const N = filtered.length;

  const hitoKey = (CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global') ? 'Global' : CURRENT_HITO;
  const hData = (RAW_DATA && RAW_DATA.hitos && RAW_DATA.hitos[hitoKey]) || (RAW_DATA && RAW_DATA.hitos && RAW_DATA.hitos['Global']);

  const pctFavorableEl = document.getElementById('kpi-pct-favorable');
  const countFavorableEl = document.getElementById('kpi-count-favorable');
  const postEl = document.getElementById('kpi-post');
  const postDetailEl = document.getElementById('kpi-post-detail');
  const hakeCardEl = document.getElementById('kpi-hake-card');
  const hakeDetailEl = document.getElementById('kpi-hake-detail');
  const deltaEl = document.getElementById('kpi-delta');
  const deltaDetailEl = document.getElementById('kpi-delta-detail');
  const coberturaEl = document.getElementById('kpi-cobertura');
  const coberturaDetailEl = document.getElementById('kpi-cobertura-detail');
  const ideasEl = document.getElementById('kpi-ideas');
  const ideasDetailEl = document.getElementById('kpi-ideas-detail');

  if (N === 0) {
    if (pctFavorableEl) pctFavorableEl.innerText = '0%';
    if (countFavorableEl) countFavorableEl.innerHTML = '<span>0 participantes</span>';
    if (postEl) postEl.innerHTML = '0.0 <span class="text-xs font-normal text-slate-400">/ 5</span>';
    if (deltaEl) deltaEl.innerText = '0.00';
    return;
  }

  const avgPre = filtered.reduce((acc, p) => acc + p.pre_score, 0) / N;
  const avgPost = filtered.reduce((acc, p) => acc + p.post_score, 0) / N;
  const avgDelta = filtered.reduce((acc, p) => acc + p.delta, 0) / N;

  const nMejora = filtered.filter(p => p.status === 'Mejoró').length;
  const nMantuvo = filtered.filter(p => p.status === 'Mantuvo').length;
  const nFavorable = nMejora + nMantuvo;
  const pctFavorable = ((nFavorable / N) * 100).toFixed(1);
  const pctLogro = ((avgPost / 5) * 100).toFixed(1);
  const pctInc = avgPre > 0 ? ((avgDelta / avgPre) * 100).toFixed(1) : '0.0';

  const hakeVal = hData ? hData.hake_gain : 0.214;
  const schoolsList = (hData && hData.schools) || [];
  const nActiveSchools = schoolsList.filter(s => s.has_data !== false).length;

  if (pctFavorableEl) pctFavorableEl.innerText = `${pctFavorable}%`;
  if (countFavorableEl) {
    countFavorableEl.innerHTML = `<span>${nFavorable} de ${N} consolidaron o avanzaron</span><span class="w-2 h-2 rounded-full bg-[#00ADBB]"></span>`;
  }

  if (postEl) {
    postEl.innerHTML = `${avgPost.toFixed(2)} <span class="text-xs font-normal text-slate-400">/ 5</span>`;
  }
  if (postDetailEl) {
    postDetailEl.innerHTML = `<span>${pctLogro}% de logro conceptual</span><div class="w-12 h-1.5 bg-[#f1f9e9] rounded-full overflow-hidden"><div class="h-full bg-[#88C23F] rounded-full" style="width: ${pctLogro}%"></div></div>`;
  }

  if (hakeCardEl) {
    hakeCardEl.innerText = `${hakeVal >= 0 ? '+' : ''}${hakeVal.toFixed(3)}`;
  }
  if (hakeDetailEl) {
    hakeDetailEl.innerHTML = `<span>Efectividad formativa</span><span class="text-xs">★</span>`;
  }

  if (deltaEl) {
    deltaEl.innerText = `${avgDelta >= 0 ? '+' : ''}${avgDelta.toFixed(2)}`;
    deltaEl.className = avgDelta >= 0 ? 'text-3xl font-extrabold text-[#52801b] tracking-tight mt-1.5' : 'text-3xl font-extrabold text-rose-600 tracking-tight mt-1.5';
  }
  if (deltaDetailEl) {
    deltaDetailEl.innerHTML = `<span>${avgDelta >= 0 ? '+' : ''}${pctInc}% incremento neto</span><span class="text-xs">↗</span>`;
  }

  if (coberturaEl) {
    coberturaEl.innerText = '100%';
  }
  if (coberturaDetailEl) {
    coberturaDetailEl.innerHTML = `<span>${nActiveSchools} de ${schoolsList.length || 7} Escuelas activas</span><span class="w-2 h-2 rounded-full bg-[#88C23F]"></span>`;
  }

  if (ideasEl) {
    ideasEl.innerText = '35';
  }
  if (ideasDetailEl) {
    ideasDetailEl.innerHTML = `<span>Proyectos en Climatón</span><i data-lucide="sparkles" class="w-3.5 h-3.5 text-[#f59e0b]"></i>`;
  }
}

// RENDER CHARTS CON CHART.JS (Estilo moderno y paleta oficial del logo PAC)
function renderCharts() {
  const filtered = getFilteredParticipants();
  const N = filtered.length;

  // 1. Status Donut Chart (Verde Lima #88C23F, Turquesa #00ADBB y Ámbar #F8B04C del logo PAC)
  const nMejora = filtered.filter(p => p.status === 'Mejoró').length;
  const nMantuvo = filtered.filter(p => p.status === 'Mantuvo').length;
  const nRetrocedio = filtered.filter(p => p.status === 'Retrocedió').length;

  const ctxStatus = document.getElementById('chart-status-donut').getContext('2d');
  if (CHARTS['status']) CHARTS['status'].destroy();

  CHARTS['status'] = new Chart(ctxStatus, {
    type: 'doughnut',
    data: {
      labels: ['Mejoró (+pts)', 'Consolidó (0 pts)', 'Requiere refuerzo (-pts)'],
      datasets: [{
        data: [nMejora, nMantuvo, nRetrocedio],
        backgroundColor: ['#88C23F', '#00ADBB', '#F8B04C'],
        borderWidth: 3,
        borderColor: '#ffffff',
        hoverOffset: 6
      }]
    },
    options: {
      animation: false,
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
        <span class="inline-block w-2.5 h-2.5 rounded-full bg-[#88C23F]"></span>
        <span class="font-bold text-slate-800">${nMejora}</span>
      </div>
      <div class="text-[10px] text-slate-400 font-medium mt-0.5">Mejoró (${N ? ((nMejora/N)*100).toFixed(0) : 0}%)</div>
    </div>
    <div class="bg-slate-50/80 p-2.5 rounded-2xl border border-slate-100">
      <div class="flex items-center justify-center gap-1.5">
        <span class="inline-block w-2.5 h-2.5 rounded-full bg-[#00ADBB]"></span>
        <span class="font-bold text-slate-800">${nMantuvo}</span>
      </div>
      <div class="text-[10px] text-slate-400 font-medium mt-0.5">Consolidó (${N ? ((nMantuvo/N)*100).toFixed(0) : 0}%)</div>
    </div>
    <div class="bg-slate-50/80 p-2.5 rounded-2xl border border-slate-100">
      <div class="flex items-center justify-center gap-1.5">
        <span class="inline-block w-2.5 h-2.5 rounded-full bg-[#F8B04C]"></span>
        <span class="font-bold text-slate-800">${nRetrocedio}</span>
      </div>
      <div class="text-[10px] text-slate-400 font-medium mt-0.5">Requiere refuerzo (${N ? ((nRetrocedio/N)*100).toFixed(0) : 0}%)</div>
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

  CHARTS['questions'] = new Chart(ctxQuestions, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: '% Aciertos PRE (Basal)',
          data: prePcts,
          backgroundColor: '#94a3b8',
          borderRadius: 8,
          borderSkipped: false
        },
        {
          label: '% Aciertos POST (Terreno)',
          data: postPcts,
          backgroundColor: '#88C23F',
          borderRadius: 8,
          borderSkipped: false
        }
      ]
    },
    options: {
      animation: false,
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
  const targetSchoolNames = new Set(LINEA_VERDE_SCHOOLS[hitoKey] || LINEA_VERDE_SCHOOLS['Global']);
  const schoolsMeta = (RAW_DATA.hitos[hitoKey].schools || []).filter(s => s.has_data !== false && targetSchoolNames.has(s.school));
  const schoolLabels = schoolsMeta.map(s => s.school.replace('Escuela ', 'Esc. ').replace('Colegio ', 'Col. '));
  const schoolDeltas = schoolsMeta.map(s => s.delta_avg);

  const ctxSchools = document.getElementById('chart-schools-delta').getContext('2d');
  if (CHARTS['schools']) CHARTS['schools'].destroy();

  const schoolColors = schoolDeltas.map(d => d >= 0 ? '#88C23F' : '#F8B04C');

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
      animation: false,
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

  CHARTS['gender'] = new Chart(ctxGender, {
    type: 'bar',
    data: {
      labels: gLabels.map(g => `${g} (N=${genderData[g].n})`),
      datasets: [
        {
          label: 'Puntaje PRE',
          data: gPre,
          backgroundColor: '#94a3b8',
          borderRadius: 8,
          borderSkipped: false
        },
        {
          label: 'Puntaje POST',
          data: gPost,
          backgroundColor: '#88C23F',
          borderRadius: 8,
          borderSkipped: false
        }
      ]
    },
    options: {
      animation: false,
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

// =========================================================================
// MÓDULO CLIMATÓN 2026: IDEAS HACIA EL 2050 (LÍNEA VERDE)
// =========================================================================

const CLIMATON_DATA_FALLBACK = {
  "Escuela Presbiteriana": {
    "desafio_codigo": "V1",
    "desafio_nombre": "Los otros vecinos",
    "eje": "Biodiversidad y Humedales",
    "equipo": 4,
    "ideas": [
      { "numero": 1, "titulo": "Lugar para ver las aves con binoculares." },
      { "numero": 2, "titulo": "Crearemos un QR con un video informativo para culturizar a la comunidad sobre nuestras aves costeras" },
      { "numero": 3, "titulo": "Dispensador de bolsas de basura." },
      { "numero": 4, "titulo": "Crear carteles de madera informativos sobre aves costeras como viven y habitan en nuestra región" },
      { "numero": 5, "titulo": "Drones que vigilen la playa" }
    ]
  },
  "Escuela Rómulo Peña": {
    "desafio_codigo": "V1",
    "desafio_nombre": "Los otros vecinos",
    "eje": "Biodiversidad y Humedales",
    "equipo": 1,
    "ideas": [
      { "numero": 1, "titulo": "Hacer juegos con materiales reciclables para la escuela y paneles para que los niños jueguen" },
      { "numero": 2, "titulo": "Poner basureros de colores con materiales de reciclajes para seleccionar basura de la escuela" },
      { "numero": 3, "titulo": "Formar equipo o grupo de estudiantes para recoger basura cuando cuidemos medio ambiente de la escuela ser: guardianes de la biodiversidad" },
      { "numero": 4, "titulo": "Hacer un jardín utilizando residuos de comida orgánica para mejorar la tierra, así las abejas tienen lugar donde vivir" },
      { "numero": 5, "titulo": "Campaña: Hacer afiches informativos para cuidar el hábitat de los animales y plantas y enseñar en actos" }
    ]
  },
  "Colegio Bet-El": {
    "desafio_codigo": "V2",
    "desafio_nombre": "El escudo invisible",
    "eje": "Atmósfera y Cielos Limpios",
    "equipo": 7,
    "ideas": [
      { "numero": 1, "titulo": "Usar luz día" },
      { "numero": 2, "titulo": "Solo luces amarillas apuntando hacia abajo" },
      { "numero": 3, "titulo": "Usar paneles solares" },
      { "numero": 4, "titulo": "Luces con sensor de movimiento" },
      { "numero": 5, "titulo": "Carteles publicitarios apagados durante el día" }
    ]
  },
  "Colegio Greenhill": {
    "desafio_codigo": "V2",
    "desafio_nombre": "El escudo invisible",
    "eje": "Atmósfera y Cielos Limpios",
    "equipo": 14,
    "ideas": [
      { "numero": 1, "titulo": "Que los mall apaguen sus luces en las noches/tardes. Poner sensores de movimiento en las casas y programar la luz para que se prenda y apaguen a una hora exacta." },
      { "numero": 2, "titulo": "Pactando con las empresas las luces en la noche como ejemplo 8 PM - 1 AM" },
      { "numero": 3, "titulo": "Bajando la intensidad de las luces, así no encandila mucho y no le hace mal a nadie" },
      { "numero": 4, "titulo": "Solo ocupar las luces en horas necesarias y no desperdiciar" },
      { "numero": 5, "titulo": "Cambiar los focos a luces cálidas y cambiar su forma para que miren hacia abajo, y las aves no se desorienten tanto" }
    ]
  },
  "Escuela La Bandera": {
    "desafio_codigo": "V1",
    "desafio_nombre": "Los otros vecinos",
    "eje": "Biodiversidad y Humedales",
    "equipo": 6,
    "ideas": [
      { "numero": 1, "titulo": "Basurero del centro y dividir contenedores en 4 secciones de reciclaje y hacerlo con dispositivo de voz claro" },
      { "numero": 2, "titulo": "Hacer una red de apoyo entre colegios y estudiantes para compartir sus experiencias con la fauna mediante un podcast" },
      { "numero": 3, "titulo": "El camino de las huellas: Poner huellas de patas en el suelo que conduzcan hacia un mirador :D" },
      { "numero": 4, "titulo": "La persona contaminada: Un estudiante se viste con residuos limpios pero con contaminación y llevar un cartel ¿Así queremos vivir?" },
      { "numero": 5, "titulo": "Cubos con fotografías de los animales, información, e imágenes" }
    ]
  },
  "Escuela Héroes de la Concepción": {
    "desafio_codigo": "V1",
    "desafio_nombre": "Los otros vecinos",
    "eje": "Biodiversidad y Humedales",
    "equipo": 2,
    "ideas": [
      { "numero": 1, "titulo": "Reunir residuos orgánicos con la comunidad, alimentar lombrices y finalmente trocar plantas a vecinos a cambio de su compromiso" },
      { "numero": 2, "titulo": "Hacer un tipo de scout que se encargue de generar conciencia sobre el medio ambiente" },
      { "numero": 3, "titulo": "Crear una app para denunciar microbasurales y que a la vez enseñe sobre la ley y cómo ayudar" },
      { "numero": 4, "titulo": "Charlas en colegios sobre los peligros de la flora y fauna costera. No botar basura, talleres en comunidades y juntas de vecinos" },
      { "numero": 5, "titulo": "Crear desafíos para la comunidad con conceptos de humedales y participación" }
    ]
  },
  "Liceo Bicentenario Andrés Sabella": {
    "desafio_codigo": "V2",
    "desafio_nombre": "El escudo invisible",
    "eje": "Atmósfera y Cielos Limpios",
    "equipo": 11,
    "ideas": [
      { "numero": 1, "titulo": "Apoyo a la ordenanza y fiscalizaciones reales hacia empresas, industrias y particulares" },
      { "numero": 2, "titulo": "Radio escolar que funcione para concientizar. Prácticas didácticas dirigidas a estudiantes según sus motivaciones" },
      { "numero": 3, "titulo": "Campaña 'Que tu noche no sea mi día' en los barrios a través de federaciones y sedes sociales" },
      { "numero": 4, "titulo": "Pantalla en ciudad mostrando fotos que se ven en Antofagasta hoy versus las de antes y que en ambos lados se cuide" },
      { "numero": 5, "titulo": "Realizar talleres y actividades integrando a amigos de las escuelas sustentables y que cada uno comparta la iniciativa que les enseñaron" }
    ]
  }
};

function getClimatonData() {
  if (RAW_DATA && RAW_DATA.climaton_ideas && Object.keys(RAW_DATA.climaton_ideas).length > 0) {
    return RAW_DATA.climaton_ideas;
  }
  return CLIMATON_DATA_FALLBACK;
}

function getSchoolClimaton(schoolName) {
  if (!schoolName) return null;
  const cData = getClimatonData();
  const target = stripAccents(schoolName);
  
  for (const [key, val] of Object.entries(cData)) {
    const kClean = stripAccents(key);
    if (kClean === target || target.includes(kClean) || kClean.includes(target)) {
      return val;
    }
    if (target.includes('heroes') && kClean.includes('heroes')) return val;
    if (target.includes('sabella') && kClean.includes('sabella')) return val;
    if (target.includes('bet-el') && (kClean.includes('bet-el') || kClean.includes('bet el'))) return val;
    if (target.includes('bandera') && kClean.includes('bandera')) return val;
    if (target.includes('presbiteriana') && kClean.includes('presbiteriana')) return val;
    if (target.includes('romulo') && kClean.includes('romulo')) return val;
    if (target.includes('greenhill') && kClean.includes('greenhill')) return val;
  }
  return null;
}

function createClimatonCardHTML(idea, schoolName, desafioCodigo, desafioNombre, totalIdeas) {
  const isV1 = desafioCodigo === 'V1';
  const headerClass = isV1 ? 'climaton-header-v1' : 'climaton-header-v2';

  let indicatorsHtml = '';
  for (let i = 1; i <= 5; i++) {
    if (i === idea.numero) {
      indicatorsHtml += `
        <span class="inline-flex items-center justify-center w-5 h-5 rounded-full ${isV1 ? 'bg-[#74ab2e]' : 'bg-[#008f9b]'} text-white font-extrabold text-[10px] shadow-xs">
          ${i}
        </span>
      `;
    } else {
      indicatorsHtml += `
        <span class="inline-flex items-center justify-center w-4 h-4 text-slate-300 font-semibold text-[10px]">
          ${i}
        </span>
      `;
    }
  }

  return `
    <div class="climaton-card flex flex-col justify-between">
      <!-- Header con gradiente oficial PAC fiel a la tarjeta física -->
      <div class="${headerClass} px-4 py-3 flex items-center justify-between text-white">
        <div class="flex items-center gap-2">
          <span class="px-2 py-0.5 rounded bg-white/20 text-white font-black text-xs tracking-wider border border-white/20">${desafioCodigo}</span>
          <span class="text-[11px] font-semibold text-white/95 tracking-wide">${escapeHtml(desafioNombre)}</span>
        </div>
        <div class="text-right">
          <span class="text-xs font-bold text-white tracking-tight">${escapeHtml(schoolName)}</span>
        </div>
      </div>

      <!-- Cuerpo pautado con renglones de libreta -->
      <div class="climaton-body-ruled p-5 flex-1 flex flex-col justify-between space-y-3 min-h-[145px]">
        <div>
          <div class="text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1 flex items-center gap-1.5">
            <span class="w-1.5 h-1.5 rounded-full ${isV1 ? 'bg-[#88C23F]' : 'bg-[#00ADBB]'}"></span>
            <span>NUESTRA IDEA</span>
          </div>
          <p class="text-slate-900 font-bold text-xs sm:text-[13px] leading-relaxed tracking-tight">
            "${escapeHtml(idea.titulo)}"
          </p>
        </div>
      </div>

      <!-- Pie de tarjeta con contador IDEA y marca PAC -->
      <div class="bg-slate-50/90 px-4 py-2 border-t border-slate-200/70 flex items-center justify-between">
        <div class="flex items-center gap-1">
          <span class="text-[9px] font-bold text-slate-500 uppercase tracking-wider mr-1">IDEA:</span>
          <div class="flex items-center gap-0.5">
            ${indicatorsHtml}
          </div>
        </div>
        <div class="flex items-center gap-1">
          <span class="text-[9px] font-black tracking-wider uppercase text-slate-600 bg-white px-2 py-0.5 rounded border border-slate-200 shadow-xs">
            PAC 2026
          </span>
        </div>
      </div>
    </div>
  `;
}

let CLIMATON_FILTER_DESAFIO = 'ALL';

function filterClimatonByDesafio(codigo) {
  CLIMATON_FILTER_DESAFIO = codigo;
  const btnAll = document.getElementById('btn-climaton-filter-all');
  const btnV1 = document.getElementById('btn-climaton-filter-v1');
  const btnV2 = document.getElementById('btn-climaton-filter-v2');

  const inactiveClasses = 'px-3.5 py-1.5 rounded-full text-xs font-medium text-slate-600 hover:text-slate-900 transition';

  if (btnAll) btnAll.className = codigo === 'ALL' ? 'px-3.5 py-1.5 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm transition' : inactiveClasses;
  if (btnV1) btnV1.className = codigo === 'V1' ? 'px-3.5 py-1.5 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm transition' : inactiveClasses;
  if (btnV2) btnV2.className = codigo === 'V2' ? 'px-3.5 py-1.5 rounded-full text-xs font-bold bg-[#00ADBB] text-white shadow-sm transition' : inactiveClasses;

  renderClimatonCards();
}

function openSchoolInClimaton(schoolName) {
  const sel = document.getElementById('select-climaton-school');
  if (sel) {
    let matchedOpt = 'ALL';
    const sClean = stripAccents(schoolName);
    for (let i = 0; i < sel.options.length; i++) {
      const optValClean = stripAccents(sel.options[i].value);
      if (optValClean.includes(sClean) || sClean.includes(optValClean)) {
        matchedOpt = sel.options[i].value;
        break;
      }
    }
    sel.value = matchedOpt;
  }
  switchTab('climaton');
  renderClimatonCards();
}

function renderClimatonCards() {
  const container = document.getElementById('climaton-cards-grid');
  if (!container) return;
  container.innerHTML = '';

  const cData = getClimatonData();
  const schoolFilter = (document.getElementById('select-climaton-school') && document.getElementById('select-climaton-school').value) || 'ALL';
  const searchFilter = (document.getElementById('input-climaton-search') && document.getElementById('input-climaton-search').value.toLowerCase().trim()) || '';

  const allCards = [];

  for (const [schName, schClim] of Object.entries(cData)) {
    // Filtrar por desafío
    if (CLIMATON_FILTER_DESAFIO !== 'ALL' && schClim.desafio_codigo !== CLIMATON_FILTER_DESAFIO) {
      continue;
    }

    // Filtrar por escuela
    if (schoolFilter !== 'ALL') {
      const sClean = stripAccents(schName);
      const fClean = stripAccents(schoolFilter);
      if (!sClean.includes(fClean) && !fClean.includes(sClean)) {
        continue;
      }
    }

    (schClim.ideas || []).forEach(idea => {
      if (searchFilter) {
        const textToSearch = `${idea.titulo} ${schName} ${schClim.desafio_nombre} ${schClim.eje}`.toLowerCase();
        if (!textToSearch.includes(searchFilter)) {
          return;
        }
      }

      allCards.push({
        idea,
        schName,
        desafio_codigo: schClim.desafio_codigo,
        desafio_nombre: schClim.desafio_nombre,
        eje: schClim.eje
      });
    });
  }

  if (allCards.length === 0) {
    container.innerHTML = `
      <div class="col-span-full p-12 text-center bg-white rounded-3xl border border-slate-100 card-shadow space-y-3">
        <i data-lucide="search-x" class="w-10 h-10 text-slate-300 mx-auto"></i>
        <h4 class="font-bold text-slate-700 text-sm">No se encontraron ideas con los filtros seleccionados</h4>
        <p class="text-xs text-slate-400">Prueba cambiando el establecimiento o la palabra de búsqueda.</p>
        <button onclick="document.getElementById('input-climaton-search').value=''; document.getElementById('select-climaton-school').value='ALL'; filterClimatonByDesafio('ALL');" class="px-4 py-1.5 text-xs font-bold text-slate-800 bg-slate-100 hover:bg-slate-200 rounded-full transition">
          Restablecer filtros
        </button>
      </div>
    `;
    if (window.lucide) lucide.createIcons();
    return;
  }

  allCards.forEach(c => {
    const cardWrapper = document.createElement('div');
    cardWrapper.innerHTML = createClimatonCardHTML(c.idea, c.schName, c.desafio_codigo, c.desafio_nombre, 5);
    container.appendChild(cardWrapper.firstElementChild);
  });

  if (window.lucide) lucide.createIcons();
}

function renderClimatonTab() {
  renderClimatonCards();
}

// 2. TAB: ESTABLECIMIENTOS TABLE & FICHAS INSTITUCIONALES
function renderSchoolsTab() {
  const hitoKey = (CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global') ? 'Global' : CURRENT_HITO;
  const targetSchoolNames = new Set(LINEA_VERDE_SCHOOLS[hitoKey] || LINEA_VERDE_SCHOOLS['Global']);
  const schools = (RAW_DATA.hitos[hitoKey].schools || []).filter(s => targetSchoolNames.has(s.school));
  const tbody = document.getElementById('schools-table-body');
  if (tbody) tbody.innerHTML = '';

  const cardsContainer = document.getElementById('schools-cards-container');
  if (cardsContainer) cardsContainer.innerHTML = '';

  schools.forEach(s => {
    const meta = (RAW_DATA.schools_directory && RAW_DATA.schools_directory[s.school]) || {};
    const climData = getSchoolClimaton(s.school);
    const eqNum = meta.equipo || s.equipo || (climData && climData.equipo) || null;
    const cert = meta.certificacion || s.certificacion || 'No';
    const hasData = s.has_data !== false && s.n > 0;

    // 1. Fila en la tabla comparativa
    if (tbody) {
      const tr = document.createElement('tr');
      tr.className = 'hover:bg-slate-50/70 transition';

      const eqBadge = eqNum ? 
        `<span class="px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-[#88C23F]/15 text-[#52801b] border border-[#88C23F]/35 mr-1.5 shrink-0">Eq. ${parseInt(eqNum)}</span>` : '';
      
      const certBadge = cert.toLowerCase() === 'si' ? 
        `<span class="ml-1 px-1.5 py-0.5 rounded text-[9px] font-bold bg-[#f1f9e9] text-[#52801b] border border-[#cbe99c]">SNCAE</span>` : '';

      const climatonBtn = climData ? 
        `<button onclick="openSchoolInClimaton('${escapeHtml(s.school)}')" class="px-3 py-1 rounded-full text-xs font-bold bg-[#88C23F]/15 text-[#52801b] border border-[#88C23F]/35 hover:bg-[#88C23F]/25 transition inline-flex items-center gap-1.5 shadow-xs" title="Ver las 5 ideas formuladas por este colegio en el Climatón">
          <i data-lucide="sparkles" class="w-3 h-3 text-[#88C23F]"></i>
          <span>5 Ideas</span>
        </button>` : 
        `<span class="text-slate-300 text-[10px]">-</span>`;

      if (!hasData) {
        tr.innerHTML = `
          <td class="py-3.5 px-4 font-bold text-slate-800 flex items-center">${eqBadge} <span>${escapeHtml(s.school)}</span></td>
          <td class="py-3.5 px-3">
            <span class="px-2.5 py-1 rounded-full text-[11px] bg-slate-100 text-slate-700 font-medium">${escapeHtml(s.dependencia)} ${certBadge}</span>
          </td>
          <td class="py-3.5 px-3 text-center text-slate-400 italic font-semibold">0 (Sin datos)</td>
          <td class="py-3.5 px-3 text-center text-slate-400">-</td>
          <td class="py-3.5 px-3 text-center text-slate-400">-</td>
          <td class="py-3.5 px-3 text-center">
            <span class="px-2.5 py-1 rounded-full text-[10px] bg-amber-50 text-amber-800 border border-amber-200 font-medium">Sin datos en este hito</span>
          </td>
          <td class="py-3.5 px-3 text-center text-slate-400">-</td>
          <td class="py-3.5 px-3 text-center text-slate-400 text-[10px]">No evaluado</td>
          <td class="py-3.5 px-3">
            <div class="text-slate-800 font-medium">${escapeHtml(s.lider1 || 'Coordinador del equipo')}</div>
            <div class="text-[10px] text-slate-400">${s.email1 || ''} ${s.tel1 ? '· Tel: ' + s.tel1 : ''}</div>
          </td>
          <td class="py-3.5 px-4 text-center">${climatonBtn}</td>
        `;
      } else {
        const deltaSign = s.delta_avg >= 0 ? '+' : '';
        const deltaBadge = s.delta_avg > 0 ? 
          `<span class="px-3 py-1 rounded-full text-xs font-bold bg-[#f1f9e9] text-[#52801b] border border-[#cbe99c]">+${s.delta_avg.toFixed(2)}</span>` : 
          (s.delta_avg === 0 ? 
            `<span class="px-3 py-1 rounded-full text-xs font-semibold bg-slate-100 text-slate-600">0.00</span>` : 
            `<span class="px-3 py-1 rounded-full text-xs font-bold bg-rose-50 text-rose-600 border border-rose-200">${s.delta_avg.toFixed(2)}</span>`);

        tr.innerHTML = `
          <td class="py-3.5 px-4 font-bold text-slate-800 flex items-center">${eqBadge} <span>${escapeHtml(s.school)}</span></td>
          <td class="py-3.5 px-3">
            <span class="px-2.5 py-1 rounded-full text-[11px] bg-slate-100 text-slate-700 font-medium">${escapeHtml(s.dependencia)} ${certBadge}</span>
          </td>
          <td class="py-3.5 px-3 text-center font-bold text-slate-800">${s.n}</td>
          <td class="py-3.5 px-3 text-center text-slate-600 font-medium">${s.pre_avg !== null ? s.pre_avg.toFixed(2) : '-'}</td>
          <td class="py-3.5 px-3 text-center text-[#52801b] font-bold">${s.post_avg !== null ? s.post_avg.toFixed(2) : '-'}</td>
          <td class="py-3.5 px-3 text-center">${deltaBadge}</td>
          <td class="py-3.5 px-3 text-center">
            <span class="px-3 py-1 rounded-full text-xs font-bold ${s.pct_mejora >= 50 ? 'bg-[#f1f9e9] text-[#52801b]' : 'bg-slate-100 text-slate-700'}">
              ${s.pct_mejora}%
            </span>
          </td>
          <td class="py-3.5 px-3">
            <div class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-slate-50 border border-slate-100 text-[11px]">
              <span class="text-[#52801b] font-bold" title="Mejoró">${s.status_counts['Mejoró'] || 0}↗</span>
              <span class="text-slate-300">·</span>
              <span class="text-[#00737d] font-semibold" title="Consolidó">${s.status_counts['Mantuvo'] || 0}=</span>
              <span class="text-slate-300">·</span>
              <span class="text-amber-700 font-bold" title="Requiere refuerzo">${s.status_counts['Retrocedió'] || 0}↘</span>
            </div>
          </td>
          <td class="py-3.5 px-3">
            <div class="text-slate-800 font-medium">${escapeHtml(s.lider1 || 'Coordinador del equipo')}</div>
            <div class="text-[10px] text-slate-400">${s.email1 || ''} ${s.tel1 ? '· Tel: ' + s.tel1 : ''}</div>
          </td>
          <td class="py-3.5 px-4 text-center">${climatonBtn}</td>
        `;
      }
      tbody.appendChild(tr);
    }

    // 2. Ficha institucional expandida
    if (cardsContainer) {
      const card = document.createElement('div');
      card.className = 'bg-white p-6 rounded-3xl border border-slate-100 card-shadow space-y-4';

      let climatonSectionHTML = '';
      if (climData && climData.ideas) {
        const miniCards = climData.ideas.map(id => `
          <div class="p-3.5 bg-slate-50/80 rounded-2xl border border-slate-200/70 hover:border-[#88C23F] transition space-y-1.5 flex flex-col justify-between">
            <div>
              <div class="flex items-center justify-between text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">
                <span>Idea ${id.numero}</span>
                <span class="px-1.5 py-0.2 rounded text-[9px] font-extrabold ${climData.desafio_codigo === 'V1' ? 'bg-[#f1f9e9] text-[#52801b]' : 'bg-[#e6f9fa] text-[#00737d]'}">${climData.desafio_codigo}</span>
              </div>
              <p class="text-xs font-bold text-slate-800 leading-snug">"${escapeHtml(id.titulo)}"</p>
            </div>
          </div>
        `).join('');

        climatonSectionHTML = `
          <div class="pt-4 border-t border-slate-100 space-y-3">
            <div class="flex items-center justify-between">
              <div class="flex items-center gap-2">
                <i data-lucide="sparkles" class="w-4 h-4 text-[#88C23F]"></i>
                <h5 class="text-xs font-bold uppercase tracking-wider text-slate-800">Propuestas Climatón 2026 (${climData.desafio_codigo} · ${climData.desafio_nombre})</h5>
              </div>
              <button onclick="openSchoolInClimaton('${escapeHtml(s.school)}')" class="text-[11px] font-bold text-[#52801b] hover:underline inline-flex items-center gap-1">
                <span>Ver en Muro Digital</span>
                <i data-lucide="arrow-right" class="w-3 h-3"></i>
              </button>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
              ${miniCards}
            </div>
          </div>
        `;
      }

      card.innerHTML = `
        <div class="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3.5">
          <div class="flex items-center gap-2.5">
            ${eqNum ? `<span class="px-3 py-1 rounded-full text-xs font-extrabold bg-[#88C23F]/15 text-[#52801b] border border-[#88C23F]/35">Equipo ${parseInt(eqNum)}</span>` : ''}
            <div>
              <h4 class="text-base font-extrabold text-slate-900">${escapeHtml(s.school)}</h4>
              <p class="text-xs text-slate-400 font-medium">${meta.direccion ? meta.direccion + ' · ' : ''}Dependencia: <strong>${s.dependencia}</strong></p>
            </div>
          </div>
          <div class="flex items-center gap-2 flex-wrap">
            <span class="text-xs px-2.5 py-1 rounded-full font-bold ${cert.toLowerCase() === 'si' ? 'bg-[#f1f9e9] text-[#52801b] border border-[#cbe99c]' : 'bg-slate-100 text-slate-500'}">
              SNCAE: ${cert}
            </span>
            <span class="text-xs px-2.5 py-1 rounded-full font-bold bg-slate-100 text-slate-600">
              PAC Previo: ${meta.participacion_previa || 'No'}
            </span>
            ${hasData ? `
              <span class="text-xs px-3 py-1 rounded-full font-bold bg-[#f1f9e9] text-[#52801b]">
                Delta: ${s.delta_avg >= 0 ? '+' : ''}${s.delta_avg.toFixed(2)} pts (${s.pct_mejora}%)
              </span>
            ` : `<span class="text-xs px-2.5 py-1 rounded-full font-bold bg-amber-100 text-amber-800">Sin datos en ${hitoKey}</span>`}
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
          <div class="p-3.5 rounded-2xl bg-slate-50/80 border border-slate-100 space-y-1">
            <div class="text-[10px] uppercase font-bold text-slate-400">Profesor/a Líder 1</div>
            <div class="text-slate-900 font-bold">${escapeHtml(s.lider1 || 'Coordinador del equipo')}</div>
            <div class="text-slate-500 text-[11px]">${s.email1 ? '✉ ' + s.email1 : ''} ${s.tel1 ? '· 📞 ' + s.tel1 : ''}</div>
          </div>
          <div class="p-3.5 rounded-2xl bg-slate-50/80 border border-slate-100 space-y-1">
            <div class="text-[10px] uppercase font-bold text-slate-400">Profesor/a Líder 2 / Apoyo</div>
            <div class="text-slate-900 font-bold">${escapeHtml(s.lider2 || meta.lider2 || 'No registrado')}</div>
            <div class="text-slate-500 text-[11px]">${s.email2 || meta.email2 ? '✉ ' + (s.email2 || meta.email2) : ''} ${(s.tel2 || meta.tel2) ? '· 📞 ' + (s.tel2 || meta.tel2) : ''}</div>
          </div>
        </div>

        ${climatonSectionHTML}
      `;
      cardsContainer.appendChild(card);
    }
  });

  if (window.lucide) lucide.createIcons();
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
          'p-4 bg-[#f1f9e9] rounded-2xl border border-[#cbe99c] text-[#3f6515] font-bold text-xs flex items-center gap-2 mt-2 mb-1 shadow-sm' :
          'p-4 bg-[#e6f9fa] rounded-2xl border border-[#99e1e7] text-[#005b63] font-bold text-xs flex items-center gap-2 mt-6 mb-1 shadow-sm';
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
            <span class="w-7 h-7 rounded-full bg-[#88C23F] text-slate-950 font-extrabold text-xs flex items-center justify-center shadow-sm">${q.num}</span>
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
          <span class="font-extrabold text-[#52801b] text-sm mt-0.5 block">${q.post_correct} (${q.post_pct}%)</span>
        </div>
        <div>
          <span class="text-slate-400 text-[10px] font-bold uppercase tracking-wider block">Consolidaron Correcto</span>
          <span class="font-extrabold text-[#00737d] text-sm mt-0.5 block">${q.transitions['Mantuvo Correcto'] || 0}</span>
        </div>
        <div>
          <span class="text-slate-400 text-[10px] font-bold uppercase tracking-wider block">Requieren Refuerzo</span>
          <span class="font-extrabold text-amber-700 text-sm mt-0.5 block">${q.transitions['Retrocedió'] || 0}</span>
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
      badgeClass = 'bg-[#e6f9fa] text-[#00737d] border border-[#99e1e7] font-semibold';
      badgeText = 'Consolidó Respuesta Informada =';
    } else if (qData.status === 'Retrocedió') {
      badgeClass = 'badge-regressed font-bold';
      badgeText = 'Requiere refuerzo en texto';
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
          ${isGlobal ? `<span class="ml-2 px-2.5 py-0.5 rounded-full text-[10px] font-bold ${originHito === 'Hito 4' ? 'bg-[#e6f9fa] text-[#00737d] border border-[#99e1e7]' : 'bg-[#f1f9e9] text-[#52801b] border border-[#cbe99c]'}">${originHito}</span>` : ''}
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
        <div class="bg-[#f1f9e9]/50 p-3.5 rounded-xl border border-[#cbe99c]">
          <span class="text-[10px] font-bold text-[#52801b] block uppercase tracking-wide mb-1">Respuesta POST (Final en terreno):</span>
          <p class="text-slate-800 font-medium">${qData.post_text && qData.post_text.trim() ? '"' + qData.post_text + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta o no sabe)</span>'}</p>
        </div>
      </div>

      ${reviewAlert}
    `;
    listEl.appendChild(item);
  });
}

// Generador pedagógico de actividades de aula (15-20 min) para eliminar brechas conceptuales
function getSchoolPedagogicalRecommendation(s, hitoKey) {
  const sch = s.school || '';
  const lider = s.lider1 || 'Profesor/a Líder';
  
  if (s.has_data === false || s.n === 0) {
    return {
      type: 'pending',
      badge: 'Evaluación Pendiente',
      badgeClass: 'bg-amber-100 text-amber-800 border-amber-300',
      boxClass: 'bg-[#fef9ed] border-[#fcd34d] text-amber-950',
      title: 'Actividad de Activación y Nivelación Inicial (20 min)',
      content: `<strong>Diagnóstico de Situación:</strong> El establecimiento no registra respuestas pareadas en este desafío.<br>
      <strong>Actividad sugerida en aula (20 min):</strong> <em>"Nivelación Rápida de Entrada"</em>. Proyectar la infografía oficial del desafío y aplicar la encuesta PRE impresa o digital en 10 minutos, seguida de una lluvia de ideas guiada para nivelar conocimientos basales.<br>
      <strong>Coordinación:</strong> Contactar directamente a la líder <strong>${lider}</strong> para coordinar la nivelación antes de la etapa de ideación del Climatón.`
    };
  }

  // Recomendaciones específicas con actividades de aula de 15-20 min por escuela
  if (sch.includes('La Bandera')) {
    return {
      type: 'action',
      badge: 'Brecha de Hábitat Costero',
      badgeClass: 'bg-amber-100 text-amber-800 border-amber-300',
      boxClass: 'bg-[#fef9ed] border-[#fcd34d] text-amber-950',
      title: 'Actividad de Aula (15-20 min): "Mapeo Territorial de Huellas"',
      content: `<strong>Diagnóstico de Brecha:</strong> Los estudiantes identifican la fauna pero persiste confusión sobre el impacto directo del tránsito 4x4 y microbasurales sobre los nidos en planicies costeras.<br>
      <strong>Paso a paso para el docente:</strong>
      <ol class="list-decimal list-inside mt-1.5 space-y-1 text-slate-700">
        <li><strong>Fase 1 (5 min):</strong> El docente proyecta un croquis simple del borde costero norte y divide al curso en grupos de 3 a 4 estudiantes.</li>
        <li><strong>Fase 2 (8 min):</strong> Cada equipo marca con puntos rojos las amenazas antrópicas (vehículos no regulados, perros, basura) y con verde las zonas seguras de nidificación.</li>
        <li><strong>Fase 3 (5 min):</strong> Pregunta clave para el Climatón: <em>"¿Cómo diseñar señalética disuasiva o cierres ecológicos sin alterar el paisaje natural?"</em>.</li>
      </ol>
      <div class="mt-2 text-[11px] text-emerald-800 font-semibold">Meta de aula: Vincular la conservación del gaviotín con soluciones territoriales concretas.</div>`
    };
  }

  if (sch.includes('Presbiteriana')) {
    return {
      type: 'action',
      badge: 'Brecha de Escala y Fragilidad',
      badgeClass: 'bg-amber-100 text-amber-800 border-amber-300',
      boxClass: 'bg-[#fef9ed] border-[#fcd34d] text-amber-950',
      title: 'Actividad de Aula (15-20 min): "El Peso de la Fragilidad"',
      content: `<strong>Diagnóstico de Brecha:</strong> Comprensión positiva de biodiversidad, pero se subestima la escala física y extrema vulnerabilidad de especies migratorias (ej. gaviotín chico: 50 gramos).<br>
      <strong>Paso a paso para el docente:</strong>
      <ol class="list-decimal list-inside mt-1.5 space-y-1 text-slate-700">
        <li><strong>Fase 1 (4 min):</strong> El profesor hace circular objetos de ~50 gramos (ej. una caja pequeña de fósforos, un huevo) para calibrar el peso real en las manos de los estudiantes.</li>
        <li><strong>Fase 2 (6 min):</strong> Pregunta disparadora: <em>"¿Cómo un ave tan liviana vuela miles de kilómetros y puede perder su nido por una simple pisada o una fogata en la playa?"</em>.</li>
        <li><strong>Fase 3 (5 min):</strong> Redacción en post-it de 1 norma comunitaria escolar que los vecinos y familias deberían respetar al visitar las playas de Antofagasta.</li>
      </ol>
      <div class="mt-2 text-[11px] text-emerald-800 font-semibold">Meta de aula: Conectar el dato cuantitativo (50 gr) con empatía y protección cívica activa.</div>`
    };
  }

  if (sch.includes('Héroes de la Concepción') || sch.includes('Heroes')) {
    return {
      type: 'action',
      badge: 'Brecha de Redes Tróficas',
      badgeClass: 'bg-[#f1f9e9] text-[#52801b] border-[#cbe99c]',
      boxClass: 'bg-[#f1f9e9] border-[#cbe99c] text-emerald-950',
      title: 'Actividad de Aula (15-20 min): "El Círculo de la Cadena Invisible"',
      content: `<strong>Diagnóstico de Brecha:</strong> Sobresaliente asimilación (+0.95 pts), pero persisten dudas sobre el rol de microorganismos de salinas y humedales como base de la cadena biológica.<br>
      <strong>Paso a paso para el docente:</strong>
      <ol class="list-decimal list-inside mt-1.5 space-y-1 text-slate-700">
        <li><strong>Fase 1 (3 min):</strong> Asignar roles en tarjetas: Bacterias extremófilas, fitoplancton, microcrustáceos, peces pequeños, aves marinas mayores (piqueros, pelícanos).</li>
        <li><strong>Fase 2 (7 min):</strong> Los estudiantes se enlazan con una cuerda o lana que representa el flujo biológico del desierto y mar.</li>
        <li><strong>Fase 3 (5 min):</strong> El docente "corta" la base por contaminación química y los estudiantes comprueban visualmente la caída de las especies superiores.</li>
      </ol>
      <div class="mt-2 text-[11px] text-emerald-800 font-semibold">Meta de aula: Fortalecer la visión ecosistémica para dotar de solidez científica las ideas del Climatón.</div>`
    };
  }

  if (sch.includes('Rómulo Peña') || sch.includes('Romulo')) {
    return {
      type: 'action',
      badge: 'Precisión de Conceptos Clave',
      badgeClass: 'bg-amber-100 text-amber-800 border-amber-300',
      boxClass: 'bg-[#fef9ed] border-[#fcd34d] text-amber-950',
      title: 'Actividad de Aula (15-20 min): "Trivia Rápida: Mitos vs Evidencia"',
      content: `<strong>Diagnóstico de Brecha:</strong> El grupo tiene entusiasmo e interés, pero en las opciones cerradas prefirieron distractores atractivos en vez de términos técnicos precisos.<br>
      <strong>Paso a paso para el docente:</strong>
      <ol class="list-decimal list-inside mt-1.5 space-y-1 text-slate-700">
        <li><strong>Fase 1 (6 min):</strong> El profesor presenta 4 afirmaciones de Verdadero/Falso en tarjetas o pizarra sobre fauna nativa vs fauna invasora.</li>
        <li><strong>Fase 2 (6 min):</strong> Desarmar en conjunto los distractores: aclarar la diferencia entre especie residente, migratoria y endémica de nuestra costa.</li>
        <li><strong>Fase 3 (5 min):</strong> Cada dupla redacta en 1 titular el problema ambiental local que su equipo quiere resolver en el Climatón.</li>
      </ol>
      <div class="mt-2 text-[11px] text-emerald-800 font-semibold">Meta de aula: Pasar de opiniones generales a propuestas fundamentadas con conceptos sólidos.</div>`
    };
  }

  if (sch.includes('Sabella')) {
    return {
      type: 'action',
      badge: 'Brecha de Luz Cenital y Apantallamiento',
      badgeClass: 'bg-amber-100 text-amber-800 border-amber-300',
      boxClass: 'bg-[#fef9ed] border-[#fcd34d] text-amber-950',
      title: 'Actividad de Aula (15-20 min): "Experimento de la Linterna Apantallada"',
      content: `<strong>Diagnóstico de Brecha:</strong> Dominio de nociones atmosféricas, pero un segmento de estudiantes no vincula la orientación de luminarias con la dispersión de luz hacia el cénit.<br>
      <strong>Paso a paso para el docente:</strong>
      <ol class="list-decimal list-inside mt-1.5 space-y-1 text-slate-700">
        <li><strong>Fase 1 (4 min):</strong> Con la sala en penumbra, encender una linterna apuntando hacia arriba e iluminando el techo (simulando postes mal orientados).</li>
        <li><strong>Fase 2 (6 min):</strong> Colocar una campana o cono de cartulina que proyecta 100% de la luz hacia el piso: comparar el deslumbramiento y la visibilidad de puntos en la pared.</li>
        <li><strong>Fase 3 (5 min):</strong> Discusión aplicada: <em>"¿Cómo podríamos auditar y readecuar las luces de nuestras canchas o colegios para proteger los cielos astronómicos?"</em>.</li>
      </ol>
      <div class="mt-2 text-[11px] text-emerald-800 font-semibold">Meta de aula: Comprender el principio físico de apantallamiento para diseñar soluciones tecnológicas viables.</div>`
    };
  }

  if (sch.includes('Greenhill')) {
    return {
      type: 'action',
      badge: 'Brecha de Espectro y Horarios',
      badgeClass: 'bg-[#f1f9e9] text-[#52801b] border-[#cbe99c]',
      boxClass: 'bg-[#f1f9e9] border-[#cbe99c] text-emerald-950',
      title: 'Actividad de Aula (15-20 min): "Auditoría de Luz y Horarios"',
      content: `<strong>Diagnóstico de Brecha:</strong> Rendimiento sobresaliente (+0.80 pts), con oportunidad de profundizar en cómo la temperatura de color (luz azul) y los horarios nocturnos afectan a las aves costeras.<br>
      <strong>Paso a paso para el docente:</strong>
      <ol class="list-decimal list-inside mt-1.5 space-y-1 text-slate-700">
        <li><strong>Fase 1 (5 min):</strong> Mostrar fotografías comparativas de luminarias cálidas (ámbar) vs frías (blancas/azules) y proyectar el mapa satelital nocturno de Antofagasta.</li>
        <li><strong>Fase 2 (6 min):</strong> Analizar por qué la luz blanca dispersa más en la atmósfera y cómo los horarios (ej. apagado de letreros 8 PM - 1 AM) salvan volantones de fardelas y gaviotines.</li>
        <li><strong>Fase 3 (5 min):</strong> Redactar una propuesta escolar de gestión horaria de luminarias para incorporarla en las tarjetas del Climatón.</li>
      </ol>
      <div class="mt-2 text-[11px] text-emerald-800 font-semibold">Meta de aula: Integrar astronomía y biología marina en una sola propuesta de eficiencia lumínica.</div>`
    };
  }

  if (sch.includes('Bet-el') || sch.includes('Bet-El')) {
    return {
      type: 'action',
      badge: 'Nivelación Inicial: Escudo Invisible',
      badgeClass: 'bg-amber-100 text-amber-800 border-amber-300',
      boxClass: 'bg-[#fef9ed] border-[#fcd34d] text-amber-950',
      title: 'Actividad de Aula (15-20 min): "Capa a Capa: Nuestro Escudo Invisible"',
      content: `<strong>Diagnóstico de Brecha:</strong> Establecimiento con evaluaciones pendientes en Hito 4. Requiere nivelación formativa en atmósfera, radiación UV y cielos limpios.<br>
      <strong>Paso a paso para el docente:</strong>
      <ol class="list-decimal list-inside mt-1.5 space-y-1 text-slate-700">
        <li><strong>Fase 1 (5 min):</strong> Dibujar en pizarra las capas atmosféricas y explicar por qué actúa como un escudo térmico y filtro de radiación perjudicial.</li>
        <li><strong>Fase 2 (7 min):</strong> Explicar por qué Antofagasta alberga los observatorios más avanzados del mundo y cómo la luz artificial no controlada "ciega" a los telescopios.</li>
        <li><strong>Fase 3 (6 min):</strong> Aplicación guiada de la evaluación diagnóstica para registrar a los estudiantes e iniciar la ideación del Climatón.</li>
      </ol>
      <div class="mt-2 text-[11px] text-emerald-800 font-semibold">Meta de aula: Nivelar los conceptos basales para que el equipo participe en igualdad de condiciones formativas.</div>`
    };
  }

  return {
    type: 'action',
    badge: 'Consolidación Pedagógica',
    badgeClass: 'bg-slate-100 text-slate-800 border-slate-300',
    boxClass: 'bg-slate-50 border-slate-200 text-slate-800',
    title: 'Actividad de Aula (15-20 min): "Plenario de Ideación Sostenible"',
    content: `<strong>Diagnóstico:</strong> Desempeño promedio con variación de ${s.delta_avg >= 0 ? '+' : ''}${s.delta_avg} pts.<br>
    <strong>Actividad sugerida:</strong> Plenario de 15 minutos repasando las 5 preguntas basales para afianzar el vocabulario técnico antes de pasar a la formulación de proyectos en el Climatón.`
  };
}

// 5. TAB: CASOS CRÍTICOS Y ORIENTACIONES PEDAGÓGICAS
function renderPedagogicalTab() {
  const hitoKey = (CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global') ? 'Global' : CURRENT_HITO;
  const rawSchools = (RAW_DATA.hitos[hitoKey] && RAW_DATA.hitos[hitoKey].schools) || [];
  
  // Filtrar estrictamente a las escuelas de la Línea Verde para el hito seleccionado
  const targetSchoolNames = new Set(LINEA_VERDE_SCHOOLS[hitoKey] || LINEA_VERDE_SCHOOLS['Global']);
  const schools = rawSchools.filter(s => targetSchoolNames.has(s.school));
  
  const container = document.getElementById('pedagogical-cards');
  container.innerHTML = '';

  schools.forEach(s => {
    const card = document.createElement('div');
    card.className = 'border border-slate-100 rounded-3xl p-5 bg-white card-shadow space-y-3.5 flex flex-col justify-between';

    const rec = getSchoolPedagogicalRecommendation(s, hitoKey);

    card.innerHTML = `
      <div class="space-y-3">
        <!-- Encabezado de la escuela -->
        <div class="flex items-center justify-between border-b border-slate-100 pb-3">
          <div>
            <h4 class="font-extrabold text-slate-900 text-sm sm:text-base">${escapeHtml(s.school)}</h4>
            <span class="text-[11px] text-slate-500 font-medium">Docente Líder: <strong>${escapeHtml(s.lider1 || 'Coordinador del equipo')}</strong> ${s.lider2 ? ' / ' + escapeHtml(s.lider2) : ''}</span>
          </div>
          <div class="text-right shrink-0">
            ${s.has_data !== false ? `
              <span class="text-xs font-bold px-2.5 py-0.5 rounded-full ${s.delta_avg >= 0 ? 'bg-[#f1f9e9] text-[#52801b] border border-[#cbe99c]' : 'bg-amber-50 text-amber-800 border border-amber-200'}">Delta: ${s.delta_avg >= 0 ? '+' : ''}${s.delta_avg}</span>
              <div class="text-[10px] text-slate-400 font-medium mt-0.5">${s.pct_mejora}% de mejora</div>
            ` : `
              <span class="text-xs font-bold text-amber-800 bg-amber-100 px-2.5 py-0.5 rounded-full border border-amber-200">Evaluación Pendiente</span>
            `}
          </div>
        </div>

        <!-- Insignias de contexto -->
        <div class="flex items-center gap-2 flex-wrap">
          <span class="px-2.5 py-0.5 rounded-full text-[10px] font-bold border ${rec.badgeClass}">${rec.badge}</span>
          <span class="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">${escapeHtml(s.dependencia || 'Municipal')}</span>
        </div>

        <!-- Cuadro de Actividad de Aula (15-20 min) -->
        <div class="p-4 rounded-2xl border text-xs leading-relaxed ${rec.boxClass} space-y-2">
          <div class="font-extrabold text-slate-900 text-xs flex items-center gap-1.5 border-b border-amber-200/60 pb-1.5">
            <i data-lucide="sparkles" class="w-3.5 h-3.5 text-amber-700 shrink-0"></i>
            <span>${rec.title}</span>
          </div>
          <div class="text-[11.5px] leading-relaxed">${rec.content}</div>
        </div>
      </div>

      <!-- Pie de tarjeta responsivo (garantiza no quebrar números ni contacto) -->
      <div class="pt-3 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 text-[11px] text-slate-500">
        <span class="inline-flex items-center px-2.5 py-1 rounded-full font-semibold bg-slate-100 text-slate-700 whitespace-nowrap w-fit">
          <i data-lucide="users" class="w-3.5 h-3.5 mr-1.5 text-slate-400"></i> ${s.n} estudiantes evaluados
        </span>
        <div class="flex items-center gap-2 flex-wrap font-medium">
          ${s.email1 ? `<a href="mailto:${escapeHtml(s.email1)}" class="hover:text-emerald-700 underline decoration-slate-300">✉ ${escapeHtml(s.email1)}</a>` : ''}
          ${s.tel1 ? `<span>· Tel: ${escapeHtml(s.tel1)}</span>` : ''}
        </div>
      </div>
    `;
    container.appendChild(card);
  });

  if (window.lucide) lucide.createIcons();
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

// 5. TAB: FICHAS INDIVIDUALES (Protegida)
function renderParticipantsTab() {
  const lockedView = document.getElementById('fichas-locked-view');
  const contentView = document.getElementById('fichas-content-view');

  if (!isFichasUnlocked()) {
    if (lockedView) lockedView.classList.remove('hidden');
    if (contentView) contentView.classList.add('hidden');
    if (window.lucide) lucide.createIcons();
    return;
  }

  if (lockedView) lockedView.classList.add('hidden');
  if (contentView) contentView.classList.remove('hidden');

  const isGlobal = CURRENT_HITO === 'Ambos' || CURRENT_HITO === 'Global';
  let filtered = getFilteredParticipants();

  // Búsqueda local de participantes en la ficha
  const searchInput = document.getElementById('participant-search-input');
  const searchVal = searchInput ? searchInput.value.toLowerCase().trim() : '';
  if (searchVal) {
    filtered = filtered.filter(p => (p.name && p.name.toLowerCase().includes(searchVal)) || (p.school && p.school.toLowerCase().includes(searchVal)));
  }

  const countEl = document.getElementById('participants-count-filtered');
  if (countEl) countEl.innerText = filtered.length;

  const tbody = document.getElementById('participants-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';

  filtered.forEach((p, idx) => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/70 transition cursor-pointer';
    tr.onclick = () => openModal(p);

    const deltaSign = p.delta >= 0 ? '+' : '';
    const badgeClass = p.status === 'Mejoró' ? 'badge-improved' : (p.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained');
    const statusLabel = p.status === 'Retrocedió' ? 'Requiere refuerzo' : (p.status === 'Mantuvo' ? 'Consolidó' : p.status);
    let realStatusLabel = p.real_status || '';
    if (realStatusLabel.includes('Dificultad Conceptual')) {
      realStatusLabel = 'Refuerzo conceptual sugerido';
    } else if (realStatusLabel === 'Rendimiento Estable') {
      realStatusLabel = 'Aprendizaje Consolidado';
    }

    let hitoBadge = '';
    if (isGlobal) {
      if (p.hitos_completados && p.hitos_completados.length > 1) {
        hitoBadge = `<span class="ml-1.5 px-2 py-0.5 rounded-full text-[10px] font-bold bg-[#f1f9e9] text-[#52801b] border border-[#cbe99c]">Ambos Hitos (H3 + H4)</span>`;
      } else if (p.hito === 'Hito 3') {
        hitoBadge = `<span class="ml-1.5 px-2 py-0.5 rounded-full text-[10px] font-bold bg-[#f1f9e9] text-[#52801b] border border-[#cbe99c]">Solo Hito 3</span>`;
      } else if (p.hito === 'Hito 4') {
        hitoBadge = `<span class="ml-1.5 px-2 py-0.5 rounded-full text-[10px] font-bold bg-[#e6f9fa] text-[#00737d] border border-[#99e1e7]">Solo Hito 4</span>`;
      }
    }

    const preDisplay = isGlobal && p.h3 && p.h4 ? 
      `<div class="font-bold text-slate-800">${p.pre_score} <span class="text-[10px] text-slate-400 font-normal">/ 5</span></div><div class="text-[10px] text-slate-400 font-normal">(${p.pre_score_raw} / 10 pts)</div>` : 
      `<span class="font-bold text-slate-800">${p.pre_score}</span> <span class="text-[10px] text-slate-400">/ 5</span>`;

    const postDisplay = isGlobal && p.h3 && p.h4 ? 
      `<div class="font-bold text-[#52801b]">${p.post_score} <span class="text-[10px] text-emerald-600 font-normal">/ 5</span></div><div class="text-[10px] text-emerald-600 font-normal">(${p.post_score_raw} / 10 pts)</div>` : 
      `<span class="font-bold text-[#52801b]">${p.post_score}</span> <span class="text-[10px] text-slate-400">/ 5</span>`;

    tr.innerHTML = `
      <td class="py-3.5 px-4 font-bold text-slate-900 flex items-center gap-2.5">
        <span class="w-2.5 h-2.5 rounded-full ${p.gender === 'Femenino' ? 'bg-pink-400' : 'bg-blue-400'}"></span>
        ${escapeHtml(p.name)}
      </td>
      <td class="py-3.5 px-3 text-slate-600 font-medium">${escapeHtml(p.school)} ${hitoBadge}</td>
      <td class="py-3.5 px-3 text-slate-500 text-[11px]">${escapeHtml(p.gender)}</td>
      <td class="py-3.5 px-3 text-center">${preDisplay}</td>
      <td class="py-3.5 px-3 text-center">${postDisplay}</td>
      <td class="py-3.5 px-3 text-center font-bold ${p.delta >= 0 ? 'text-[#52801b]' : 'text-amber-700'}">${deltaSign}${p.delta}</td>
      <td class="py-3.5 px-3 text-center">
        <span class="px-2.5 py-0.5 rounded-full text-[11px] font-bold ${badgeClass}">${statusLabel}</span>
        <div class="text-[10px] text-slate-400 mt-0.5 font-medium">${escapeHtml(realStatusLabel)}</div>
      </td>
      <td class="py-3.5 px-4 text-center">
        <button onclick="event.stopPropagation(); openModal(window.currentFiltered[${idx}])" class="px-3.5 py-1 text-xs font-bold text-slate-900 bg-[#88C23F] hover:bg-[#78ad35] rounded-full transition shadow-sm">
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
  if (!isFichasUnlocked()) {
    switchTab('participantes');
    return;
  }

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
    badgeContainer.innerHTML = `<span class="text-[10px] font-bold px-2.5 py-0.5 rounded-full bg-[#e6f9fa] text-[#00737d] border border-[#99e1e7]">Test-Retest Completo (Hito 3 y Hito 4)</span>`;
  } else if (h3Data) {
    badgeContainer.innerHTML = `<span class="text-[10px] font-bold px-2.5 py-0.5 rounded-full bg-[#f1f9e9] text-[#52801b] border border-[#cbe99c]">Evaluado en Hito 3</span>`;
  } else if (h4Data) {
    badgeContainer.innerHTML = `<span class="text-[10px] font-bold px-2.5 py-0.5 rounded-full bg-[#e6f9fa] text-[#00737d] border border-[#99e1e7]">Evaluado en Hito 4</span>`;
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

  const inactiveClasses = 'px-3.5 py-1 rounded-full text-xs font-medium text-slate-600 hover:text-slate-900 transition';

  if (btnConsolidado) btnConsolidado.className = mode === 'consolidado' ? 'px-3.5 py-1 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm transition' : inactiveClasses;
  if (btnH3) btnH3.className = mode === 'Hito 3' ? 'px-3.5 py-1 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm transition' : inactiveClasses;
  if (btnH4) btnH4.className = mode === 'Hito 4' ? 'px-3.5 py-1 rounded-full text-xs font-bold bg-[#00ADBB] text-white shadow-sm transition' : inactiveClasses;

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
      <div class="p-2.5 bg-[#f1f9e9]/60 rounded-xl border border-[#cbe99c] shadow-sm">
        <div class="text-[10px] text-[#52801b] font-bold uppercase tracking-wider">Hito 3 (Biodiversidad)</div>
        <div class="text-xl font-extrabold text-[#52801b] mt-1">${h3Data ? h3Data.post_score : '-'} <span class="text-xs text-slate-400 font-medium">/ 5 pts</span></div>
        <div class="text-[11px] text-slate-500 font-medium mt-0.5">${h3Data ? `PRE: ${h3Data.pre_score} · ${h3Data.delta >= 0 ? '+' : ''}${h3Data.delta} pts` : 'No registrado'}</div>
      </div>
      <div class="p-2.5 bg-[#e6f9fa]/60 rounded-xl border border-[#99e1e7] shadow-sm">
        <div class="text-[10px] text-[#00737d] font-bold uppercase tracking-wider">Hito 4 (Atmósfera)</div>
        <div class="text-xl font-extrabold text-[#00737d] mt-1">${h4Data ? h4Data.post_score : '-'} <span class="text-xs text-slate-400 font-medium">/ 5 pts</span></div>
        <div class="text-[11px] text-slate-500 font-medium mt-0.5">${h4Data ? `PRE: ${h4Data.pre_score} · ${h4Data.delta >= 0 ? '+' : ''}${h4Data.delta} pts` : 'No registrado'}</div>
      </div>
      <div class="p-2.5 bg-white rounded-xl border border-slate-100 shadow-sm">
        <div class="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Diagnóstico de Aprendizajes</div>
        <div class="mt-1">
          <span class="inline-block font-bold px-2.5 py-0.5 rounded-full text-[11px] ${p.status === 'Mejoró' ? 'badge-improved' : (p.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained')}">${p.status === 'Retrocedió' ? 'Requiere refuerzo' : (p.status === 'Mantuvo' ? 'Consolidó' : p.status)}</span>
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
        <i data-lucide="check-circle-2" class="w-4 h-4 text-[#88C23F]"></i>
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
        <i data-lucide="message-square" class="w-4 h-4 text-[#88C23F]"></i>
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
        <div class="text-xl font-extrabold text-[#52801b] mt-0.5">${dataH.post_score} / 5</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Diferencia</div>
        <div class="text-xl font-extrabold ${dataH.delta >= 0 ? 'text-[#52801b]' : 'text-amber-700'} mt-0.5">${deltaSign}${dataH.delta} pts</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Condición</div>
        <span class="mt-1 inline-block font-bold px-2.5 py-0.5 rounded-full text-[11px] ${dataH.status === 'Mejoró' ? 'badge-improved' : (dataH.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained')}">${dataH.status === 'Retrocedió' ? 'Requiere refuerzo' : (dataH.status === 'Mantuvo' ? 'Consolidó' : dataH.status)}</span>
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
        <div class="text-xl font-extrabold text-[#00737d] mt-0.5">${dataH.post_score} / 5</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Diferencia</div>
        <div class="text-xl font-extrabold ${dataH.delta >= 0 ? 'text-[#00737d]' : 'text-amber-700'} mt-0.5">${deltaSign}${dataH.delta} pts</div>
      </div>
      <div>
        <div class="text-[10px] text-slate-400 font-bold uppercase">Condición</div>
        <span class="mt-1 inline-block font-bold px-2.5 py-0.5 rounded-full text-[11px] ${dataH.status === 'Mejoró' ? 'badge-improved' : (dataH.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained')}">${dataH.status === 'Retrocedió' ? 'Requiere refuerzo' : (dataH.status === 'Mantuvo' ? 'Consolidó' : dataH.status)}</span>
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
  const themeBorder = theme === 'indigo' ? 'border-[#99e1e7] bg-[#e6f9fa]/20' : 'border-[#cbe99c] bg-[#f1f9e9]/20';
  const themeBadge = theme === 'indigo' ? 'bg-[#e6f9fa] text-[#00737d]' : 'bg-[#f1f9e9] text-[#52801b]';

  let itemsHtml = '';
  (hData.q_evo || []).forEach((evo, idx) => {
    if (idx >= qMeta.length) return;
    const isCorrectPost = (hData.post_mc[idx] || '').toLowerCase() === key[idx];
    const isCorrectPre = (hData.pre_mc[idx] || '').toLowerCase() === key[idx];

    let evoBadgeClass = 'bg-slate-100 text-slate-700';
    if (evo === 'Mejoró') evoBadgeClass = 'badge-improved font-bold';
    else if (evo === 'Retrocedió') evoBadgeClass = 'badge-regressed font-bold';
    else if (evo === 'Mantuvo Correcto') evoBadgeClass = 'badge-maintained font-medium';

    const evoText = evo === 'Retrocedió' ? 'Requiere refuerzo' : (evo === 'Mantuvo Correcto' ? 'Consolidó Correcto' : (evo === 'Mantuvo Incorrecto' ? 'Persiste Brecha' : evo));

    itemsHtml += `
      <div class="p-3.5 rounded-2xl border ${isCorrectPost ? 'bg-[#f1f9e9]/40 border-[#cbe99c]' : 'bg-amber-50/30 border-amber-200'} text-xs flex flex-wrap items-center justify-between gap-2">
        <div class="space-y-0.5 max-w-[75%]">
          <span class="font-bold text-slate-800">Q${idx + 1}: ${escapeHtml(qMeta[idx].title)}</span>
          <div class="text-[11px] text-slate-500 font-medium">
            PRE: <span class="font-bold ${isCorrectPre ? 'text-[#52801b]' : 'text-slate-700'}">${(hData.pre_mc[idx] || 'N/A').toUpperCase()}</span> ➔ 
            POST: <span class="font-bold ${isCorrectPost ? 'text-[#52801b]' : 'text-amber-700'}">${(hData.post_mc[idx] || 'N/A').toUpperCase()}</span>
            <span class="text-slate-400 font-normal ml-1">(Correcta: ${(key[idx] || '').toUpperCase()})</span>
          </div>
        </div>
        <div>
          <span class="text-[10px] px-3 py-1 rounded-full font-bold ${evoBadgeClass}">
            ${evoText}
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
          Puntaje: <span class="${theme === 'indigo' ? 'text-[#00737d]' : 'text-[#52801b]'} font-extrabold text-sm">${hData.post_score} / 5</span> 
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

  const themeBorder = theme === 'indigo' ? 'border-[#99e1e7] bg-[#e6f9fa]/20' : 'border-[#cbe99c] bg-[#f1f9e9]/20';
  const themeBadge = theme === 'indigo' ? 'bg-[#e6f9fa] text-[#00737d]' : 'bg-[#f1f9e9] text-[#52801b]';

  let itemsHtml = '';
  hData.open_analysis.forEach((oq, idx) => {
    let oBadge = 'badge-maintained';
    let oBadgeText = oq.status === 'Retrocedió' ? 'Requiere refuerzo' : (oq.status === 'Mantuvo' ? 'Consolidó' : oq.status);

    if (oq.status === 'Mejoró') {
      oBadge = 'badge-improved font-bold';
      oBadgeText = 'Mejora Cualitativa ↗';
    } else if (oq.status === 'Retrocedió') {
      oBadge = 'badge-regressed font-bold';
      oBadgeText = 'Requiere refuerzo en texto';
    } else if (oq.status === 'Mantuvo') {
      oBadge = 'bg-[#e6f9fa] text-[#00737d] border border-[#99e1e7] font-semibold';
      oBadgeText = 'Consolidó Respuesta =';
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
          <div class="bg-[#f1f9e9]/40 p-3 rounded-xl border border-[#cbe99c]">
            <span class="text-[10px] font-bold text-[#52801b] block uppercase tracking-wide mb-0.5">POST (En terreno):</span>
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
  if (!p) {
    alert('Por favor selecciona una ficha de estudiante primero.');
    return;
  }

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

  // Asegurar resolución de datos para Hito 3 y Hito 4
  const allData = getParticipantAllData(p.name);
  let h3Data = p.h3 || allData.h3;
  let h4Data = p.h4 || allData.h4;

  if (!h3Data && !h4Data) {
    if (p.hito === 'Hito 4' || (p.school && LINEA_VERDE_SCHOOLS['Hito 4'] && LINEA_VERDE_SCHOOLS['Hito 4'].includes(p.school))) {
      h4Data = p;
    } else {
      h3Data = p;
    }
  }
  if (!h3Data && p.hito === 'Hito 3') h3Data = p;
  if (!h4Data && p.hito === 'Hito 4') h4Data = p;

  const h3Meta = (RAW_DATA && RAW_DATA.hitos && RAW_DATA.hitos['Hito 3'] && RAW_DATA.hitos['Hito 3'].questions) || [];
  const h4Meta = (RAW_DATA && RAW_DATA.hitos && RAW_DATA.hitos['Hito 4'] && RAW_DATA.hitos['Hito 4'].questions) || [];

  // Remover cualquier elemento de impresión residual previo
  const oldDoc = document.getElementById('pdf-printable-document');
  if (oldDoc && oldDoc.parentNode) {
    oldDoc.parentNode.removeChild(oldDoc);
  }

  // Contenedor del documento para renderizado PDF
  // NOTA CLAVE: Ancho optimizado a 700px con box-sizing para encajar perfectamente
  // en el ancho imprimible A4 (194mm = ~733px a 96dpi) sin desbordar el extremo derecho.
  const printableDoc = document.createElement('div');
  printableDoc.id = 'pdf-printable-document';
  printableDoc.innerHTML = buildPDFDocumentHTML(p, h3Data, h4Data, currentModalMode, h3Meta, h4Meta);
  printableDoc.style.width = '700px';
  printableDoc.style.maxWidth = '700px';
  printableDoc.style.backgroundColor = '#ffffff';
  printableDoc.style.boxSizing = 'border-box';

  const cleanFileName = `Ficha_PAC_2026_${p.name.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]/g, '_')}.pdf`;

  const restoreUI = () => {
    if (btn) {
      btn.innerHTML = originalHtml;
      btn.disabled = false;
    }
    if (document.body.contains(printableDoc)) {
      document.body.removeChild(printableDoc);
    }
  };

  const fallbackPrint = () => {
    if (!document.body.contains(printableDoc)) {
      printableDoc.style.position = 'static';
      printableDoc.style.display = 'block';
      document.body.appendChild(printableDoc);
    }
    window.print();
    setTimeout(restoreUI, 1200);
  };

  // Esperar a que las imágenes (logos) estén cargadas antes de renderizar el canvas
  const images = Array.from(printableDoc.querySelectorAll('img'));
  const imagePromises = images.map(img => {
    if (img.complete) return Promise.resolve();
    return new Promise(res => {
      img.onload = res;
      img.onerror = res;
    });
  });

  Promise.all(imagePromises).then(() => {
    if (typeof html2pdf !== 'undefined') {
      const opt = {
        margin: [8, 8, 8, 8],
        filename: cleanFileName,
        image: { type: 'jpeg', quality: 0.98 },
        enableLinks: true,
        html2canvas: {
          scale: 2,
          useCORS: true,
          logging: false
        },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak: { mode: ['css', 'legacy'] }
      };

      html2pdf().set(opt).from(printableDoc).save().then(() => {
        restoreUI();
      }).catch(err => {
        console.warn('html2pdf falló, recurriendo a impresión nativa:', err);
        fallbackPrint();
      });
    } else {
      fallbackPrint();
    }
  }).catch(() => {
    fallbackPrint();
  });
}

// Construye la estructura HTML optimizada para impresión en PDF A4 (Exactamente 2 páginas)
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

  // Diagnóstico pedagógico dinámico y altamente accionable
  let orientacionDocente = '';
  let orientacionActividad = '';
  let orientacionGuias = [];

  const h3Gap = h3Data && (h3Data.delta < 0 || h3Data.post_score < 4 || (h3Data.open_analysis && h3Data.open_analysis.some(o => o.status === 'Retrocedió' || o.status === 'Requiere Refuerzo')));
  const h4Gap = h4Data && (h4Data.delta < 0 || h4Data.post_score < 4 || (h4Data.open_analysis && h4Data.open_analysis.some(o => o.status === 'Retrocedió' || o.status === 'Requiere Refuerzo')));

  if (h3Data && h4Data) {
    if (h3Gap && h4Gap) {
      orientacionDocente = 'Se observa un avance heterogéneo con elementos que requieren refuerzo en ambos hitos. En Biodiversidad precisa afianzar las tramas tróficas y el rol del humedal La Chimba; en Atmósfera requiere clarificar la distinción entre efecto invernadero natural vs. cambio climático y el impacto de luminarias en aves marinas.';
      orientacionActividad = 'Taller práctico de aula sobre redes ecológicas litorales y auditoría escolar de luminarias hacia el suelo (Decreto 1 MMA).';
      orientacionGuias = [
        { name: 'Guía Docente Hito 3: Biodiversidad Costera', file: 'Guia_Docente_Hito3_Biodiversidad_Costera.pdf' },
        { name: 'Guía Docente Hito 4: Cielos Limpios y Atmósfera', file: 'Guia_Docente_Hito4_Cielos_Atmosfera.pdf' }
      ];
    } else if (h3Gap) {
      orientacionDocente = 'Sólido dominio en contenidos de atmósfera con necesidad de afianzar conceptos ecológicos en Hito 3 (redes tróficas, especies descomponedoras y amenazas por perros asilvestrados al Gaviotín chico).';
      orientacionActividad = 'Simulación de camuflaje de nidos costeros y elaboración de un mapa de relaciones tróficas en el humedal La Chimba.';
      orientacionGuias = [
        { name: 'Guía Docente Hito 3: Biodiversidad Costera', file: 'Guia_Docente_Hito3_Biodiversidad_Costera.pdf' }
      ];
    } else if (h4Gap) {
      orientacionDocente = 'Sólido dominio en biodiversidad con brecha conceptual en Hito 4: requiere comprender el efecto invernadero como escudo regulador del clima y la orientación hacia el suelo en luminarias urbanas.';
      orientacionActividad = 'Experimento de retención térmica en frasco y campaña escolar de rescate de volantones de golondrinas de mar.';
      orientacionGuias = [
        { name: 'Guía Docente Hito 4: Cielos Limpios y Atmósfera', file: 'Guia_Docente_Hito4_Cielos_Atmosfera.pdf' }
      ];
    } else {
      orientacionDocente = 'Desempeño destacado y consolidado. El participante evidencia alta apropiación conceptual y argumentativa en ambos desafíos, posicionándose como un líder natural para formular iniciativas escolares.';
      orientacionActividad = 'Canalizar su motivación hacia la formulación de una Ficha de Proyecto para el Climatón escolar 2026.';
      orientacionGuias = [
        { name: 'Guía de Formulación de Proyectos Climatón 2026', file: 'Guia_Docente_Climaton_Accion_Climatica.pdf' }
      ];
    }
  } else if (h3Data) {
    if (h3Gap) {
      orientacionDocente = 'El estudiante evidencia avance formativo con necesidad de reforzar el rol de los descomponedores, las tramas tróficas del litoral y el impacto de los perros asilvestrados sobre el Gaviotín chico.';
      orientacionActividad = 'Taller práctico de redes ecológicas y modelamiento de hábitats protegidos para aves litorales de Antofagasta.';
      orientacionGuias = [
        { name: 'Guía Docente Hito 3: Biodiversidad Costera', file: 'Guia_Docente_Hito3_Biodiversidad_Costera.pdf' }
      ];
    } else {
      orientacionDocente = 'Apropiación sobresaliente de los contenidos de biodiversidad y ecosistemas costeros. Se sugiere incentivar el liderazgo en proyectos ambientales escolares.';
      orientacionActividad = 'Formular una iniciativa escolar de protección ambiental para presentar en el Climatón 2026.';
      orientacionGuias = [
        { name: 'Guía de Formulación de Proyectos Climatón 2026', file: 'Guia_Docente_Climaton_Accion_Climatica.pdf' }
      ];
    }
  } else if (h4Data) {
    if (h4Gap) {
      orientacionDocente = 'El participante requiere afianzar el rol de la atmósfera en la regulación climática y comprender la norma de emisión lumínica (Decreto 1 MMA) protegiendo cielos astronómicos y fauna silvestre.';
      orientacionActividad = 'Auditoría escolar de luminarias en el recinto educativo y diseño de afiches de concientización comunitaria.';
      orientacionGuias = [
        { name: 'Guía Docente Hito 4: Cielos Limpios y Atmósfera', file: 'Guia_Docente_Hito4_Cielos_Atmosfera.pdf' }
      ];
    } else {
      orientacionDocente = 'Dominio integral y reflexivo de la física atmosférica y la protección del cielo oscuro en la Región de Antofagasta.';
      orientacionActividad = 'Proponer soluciones de eficiencia lumínica escolar en el marco del Climatón 2026.';
      orientacionGuias = [
        { name: 'Guía de Formulación de Proyectos Climatón 2026', file: 'Guia_Docente_Climaton_Accion_Climatica.pdf' }
      ];
    }
  }

  // Filas para preguntas cerradas H3
  let h3MCRows = '';
  if (h3Data && h3Data.pre_mc && h3Data.post_mc) {
    const key3 = h3Meta.slice(0, 5).map(q => (q.correct_key || '').toUpperCase());
    h3Data.q_evo.forEach((evo, idx) => {
      if (idx >= h3Meta.length) return;
      const isPostCorr = (h3Data.post_mc[idx] || '').toUpperCase() === key3[idx];
      h3MCRows += `
        <tr>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; font-weight: bold; text-align: center; width: 25px;">Q${idx + 1}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0;">${escapeHtml(h3Meta[idx].title)}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; width: 35px;">${(h3Data.pre_mc[idx] || 'N/A').toUpperCase()}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: ${isPostCorr ? '#52801b' : '#e11d48'}; width: 35px;">${(h3Data.post_mc[idx] || 'N/A').toUpperCase()}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #52801b; width: 45px;">${key3[idx]}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; text-align: center; font-size: 9px; width: 85px;">${evo === 'Retrocedió' ? 'Req. refuerzo' : (evo === 'Mantuvo Correcto' ? 'Consolidó' : (evo === 'Mantuvo Incorrecto' ? 'Persiste Brecha' : evo))}</td>
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
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; font-weight: bold; text-align: center; width: 25px;">Q${idx + 1}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0;">${escapeHtml(h4Meta[idx].title)}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; width: 35px;">${(h4Data.pre_mc[idx] || 'N/A').toUpperCase()}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: ${isPostCorr ? '#00737d' : '#e11d48'}; width: 35px;">${(h4Data.post_mc[idx] || 'N/A').toUpperCase()}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #00737d; width: 45px;">${key4[idx]}</td>
          <td style="padding: 2.5px 4px; border: 1px solid #e2e8f0; text-align: center; font-size: 9px; width: 85px;">${evo === 'Retrocedió' ? 'Req. refuerzo' : (evo === 'Mantuvo Correcto' ? 'Consolidó' : (evo === 'Mantuvo Incorrecto' ? 'Persiste Brecha' : evo))}</td>
        </tr>
      `;
    });
  }

  // Tarjetas para preguntas abiertas H3 (diseño compacto)
  let h3OpenCards = '';
  if (h3Data && h3Data.open_analysis) {
    h3Data.open_analysis.forEach((oq, idx) => {
      h3OpenCards += `
        <div style="margin-bottom: 5px; padding: 4px 6px; border: 1px solid #e2e8f0; border-radius: 6px; background-color: #fafafa; box-sizing: border-box;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 2px;">
            <strong style="font-size: 9.5px; color: #0f172a;">P${idx + 1}: ${escapeHtml(oq.question)}</strong>
            <span style="font-size: 8px; font-weight: bold; padding: 1px 5px; border-radius: 8px; background-color: ${oq.status === 'Mejoró' ? '#f1f9e9' : (oq.status === 'Retrocedió' ? '#fef9ed' : '#e6f9fa')}; color: ${oq.status === 'Mejoró' ? '#52801b' : (oq.status === 'Retrocedió' ? '#b45309' : '#00737d')}; white-space: nowrap;">${oq.status === 'Retrocedió' ? 'Requiere refuerzo' : (oq.status === 'Mantuvo' ? 'Consolidó' : oq.status)}</span>
          </div>
          <div style="font-size: 8.5px; color: #64748b; margin-bottom: 2px;">Diagnóstico: <strong>${escapeHtml(oq.detail)}</strong></div>
          <table style="width: 100%; border-collapse: collapse; font-size: 9px; table-layout: fixed;">
            <tr>
              <td style="width: 50%; padding: 3px 5px; border: 1px solid #e2e8f0; background-color: #ffffff; vertical-align: top; word-break: break-word;">
                <span style="font-size: 8px; font-weight: bold; color: #94a3b8; display: block; text-transform: uppercase;">PRE:</span>
                <span style="font-style: italic; color: #334155;">${oq.pre_text && oq.pre_text.trim() ? '"' + escapeHtml(oq.pre_text) + '"' : '(Sin respuesta)'}</span>
              </td>
              <td style="width: 50%; padding: 3px 5px; border: 1px solid #e2e8f0; background-color: #f1f9e9; vertical-align: top; word-break: break-word;">
                <span style="font-size: 8px; font-weight: bold; color: #52801b; display: block; text-transform: uppercase;">POST:</span>
                <span style="color: #0f172a; font-weight: 500;">${oq.post_text && oq.post_text.trim() ? '"' + escapeHtml(oq.post_text) + '"' : '(Sin respuesta)'}</span>
              </td>
            </tr>
          </table>
        </div>
      `;
    });
  }

  // Tarjetas para preguntas abiertas H4 (diseño compacto)
  let h4OpenCards = '';
  if (h4Data && h4Data.open_analysis) {
    h4Data.open_analysis.forEach((oq, idx) => {
      h4OpenCards += `
        <div style="margin-bottom: 5px; padding: 4px 6px; border: 1px solid #e2e8f0; border-radius: 6px; background-color: #fafafa; box-sizing: border-box;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 2px;">
            <strong style="font-size: 9.5px; color: #0f172a;">P${idx + 1}: ${escapeHtml(oq.question)}</strong>
            <span style="font-size: 8px; font-weight: bold; padding: 1px 5px; border-radius: 8px; background-color: ${oq.status === 'Mejoró' ? '#e6f9fa' : (oq.status === 'Retrocedió' ? '#fef9ed' : '#e6f9fa')}; color: ${oq.status === 'Mejoró' ? '#00737d' : (oq.status === 'Retrocedió' ? '#b45309' : '#00737d')}; white-space: nowrap;">${oq.status === 'Retrocedió' ? 'Requiere refuerzo' : (oq.status === 'Mantuvo' ? 'Consolidó' : oq.status)}</span>
          </div>
          <div style="font-size: 8.5px; color: #64748b; margin-bottom: 2px;">Diagnóstico: <strong>${escapeHtml(oq.detail)}</strong></div>
          <table style="width: 100%; border-collapse: collapse; font-size: 9px; table-layout: fixed;">
            <tr>
              <td style="width: 50%; padding: 3px 5px; border: 1px solid #e2e8f0; background-color: #ffffff; vertical-align: top; word-break: break-word;">
                <span style="font-size: 8px; font-weight: bold; color: #94a3b8; display: block; text-transform: uppercase;">PRE:</span>
                <span style="font-style: italic; color: #334155;">${oq.pre_text && oq.pre_text.trim() ? '"' + escapeHtml(oq.pre_text) + '"' : '(Sin respuesta)'}</span>
              </td>
              <td style="width: 50%; padding: 3px 5px; border: 1px solid #e2e8f0; background-color: #e6f9fa; vertical-align: top; word-break: break-word;">
                <span style="font-size: 8px; font-weight: bold; color: #00737d; display: block; text-transform: uppercase;">POST:</span>
                <span style="color: #0f172a; font-weight: 500;">${oq.post_text && oq.post_text.trim() ? '"' + escapeHtml(oq.post_text) + '"' : '(Sin respuesta)'}</span>
              </td>
            </tr>
          </table>
        </div>
      `;
    });
  }

  return `
    <div style="font-family: Arial, Helvetica, sans-serif; color: #1e293b; padding: 0 4px; line-height: 1.3; background-color: #ffffff; box-sizing: border-box; width: 100%;">
      
      <!-- PÁGINA 1: Encabezado + Identificación + Punto 1 (Cuantitativo) + Punto 2 (Orientación Docente) + Punto 3 (Preguntas Cerradas) -->
      <div>
        <!-- Encabezado con Logos Oficiales -->
        <table style="width: 100%; border-bottom: 2px solid #88C23F; padding-bottom: 6px; margin-bottom: 8px; table-layout: fixed;">
          <tr>
            <td style="width: 60px; vertical-align: middle;">
              <img src="logo.png" style="height: 40px; width: auto; max-width: 55px; object-fit: contain;" alt="PAC 2026">
            </td>
            <td style="vertical-align: middle; padding-left: 8px;">
              <div style="font-size: 11px; font-weight: bold; text-transform: uppercase; color: #0f172a; letter-spacing: 0.5px;">Programa de Acción Climática 2026 · Línea Verde</div>
              <div style="font-size: 14px; font-weight: 900; color: #52801b; margin-top: 1px;">FICHA INDIVIDUAL DE EVALUACIÓN TEST-RETEST</div>
              <div style="font-size: 9px; color: #64748b;">Antofagasta: Laboratorio Vivo hacia el 2050 · Reporte de Aprendizajes</div>
            </td>
            <td style="width: 105px; text-align: right; vertical-align: middle;">
              <img src="logo_verdical.png" style="height: 32px; width: auto; max-width: 100px; object-fit: contain;" alt="Grupo Verdical">
            </td>
          </tr>
        </table>

        <!-- Ficha de Identificación del Participante -->
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 8px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; table-layout: fixed;">
          <tr>
            <td style="padding: 5px 8px; width: 50%; vertical-align: top;">
              <div style="font-size: 8px; text-transform: uppercase; font-weight: bold; color: #94a3b8;">Participante:</div>
              <div style="font-size: 13px; font-weight: 900; color: #0f172a;">${escapeHtml(p.name)}</div>
              <div style="font-size: 9.5px; color: #475569; margin-top: 1px;">Establecimiento: <strong>${escapeHtml(p.school)}</strong></div>
            </td>
            <td style="padding: 5px 8px; width: 50%; vertical-align: top; border-left: 1px solid #e2e8f0;">
              <div style="font-size: 8px; text-transform: uppercase; font-weight: bold; color: #94a3b8;">Detalles de Participación:</div>
              <div style="font-size: 9.5px; color: #334155;">Género: <strong>${escapeHtml(p.gender)}</strong> · Rol: <strong>${escapeHtml(p.role || 'Estudiante')}</strong></div>
              <div style="font-size: 9.5px; color: #52801b; font-weight: bold; margin-top: 1px;">${coverageText}</div>
              <div style="font-size: 9.5px; color: #64748b; margin-top: 1px;">Diagnóstico: <strong style="color: #0f172a;">${escapeHtml(p.real_status || p.status)}</strong></div>
            </td>
          </tr>
        </table>

        <!-- 1. RESUMEN CUANTITATIVO DE RENDIMIENTO -->
        <div style="margin-bottom: 8px;">
          <div style="font-size: 10.5px; font-weight: bold; text-transform: uppercase; color: #0f172a; margin-bottom: 4px; border-left: 3px solid #88C23F; padding-left: 6px;">1. Resumen Cuantitativo de Rendimiento (Puntajes PRE vs. POST)</div>
          <table style="width: 100%; border-collapse: collapse; font-size: 9.5px; table-layout: fixed;">
            <thead>
              <tr style="background-color: #f1f5f9; text-align: left;">
                <th style="padding: 3px 5px; border: 1px solid #cbd5e1; width: 44%;">Desafío / Hito Evaluado</th>
                <th style="padding: 3px 5px; border: 1px solid #cbd5e1; text-align: center; width: 13%;">Puntaje PRE</th>
                <th style="padding: 3px 5px; border: 1px solid #cbd5e1; text-align: center; width: 13%;">Puntaje POST</th>
                <th style="padding: 3px 5px; border: 1px solid #cbd5e1; text-align: center; width: 14%;">Diferencia</th>
                <th style="padding: 3px 5px; border: 1px solid #cbd5e1; text-align: center; width: 16%;">Condición</th>
              </tr>
            </thead>
            <tbody>
              ${h3Data ? `
                <tr>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; font-weight: bold;">Hito 3: Los otros vecinos de Antofagasta (Biodiversidad)</td>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center;">${h3Data.pre_score} / 5</td>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #52801b;">${h3Data.post_score} / 5</td>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: ${h3Data.delta >= 0 ? '#52801b' : '#e11d48'};">${h3Data.delta >= 0 ? '+' : ''}${h3Data.delta} pts</td>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center; font-weight: 600;">${h3Data.status === 'Retrocedió' ? 'Requiere refuerzo' : (h3Data.status === 'Mantuvo' ? 'Consolidó' : h3Data.status)}</td>
                </tr>
              ` : `
                <tr style="color: #94a3b8; font-style: italic;">
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0;">Hito 3: Los otros vecinos de Antofagasta</td>
                  <td colspan="4" style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center;">No convocado a este hito (escuela asignada a Hito 4)</td>
                </tr>
              `}
              ${h4Data ? `
                <tr>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; font-weight: bold;">Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos)</td>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center;">${h4Data.pre_score} / 5</td>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: #00737d;">${h4Data.post_score} / 5</td>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold; color: ${h4Data.delta >= 0 ? '#00737d' : '#e11d48'};">${h4Data.delta >= 0 ? '+' : ''}${h4Data.delta} pts</td>
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center; font-weight: 600;">${h4Data.status === 'Retrocedió' ? 'Requiere refuerzo' : (h4Data.status === 'Mantuvo' ? 'Consolidó' : h4Data.status)}</td>
                </tr>
              ` : `
                <tr style="color: #94a3b8; font-style: italic;">
                  <td style="padding: 3px 5px; border: 1px solid #e2e8f0;">Hito 4: Proteger nuestro escudo invisible</td>
                  <td colspan="4" style="padding: 3px 5px; border: 1px solid #e2e8f0; text-align: center;">No convocado a este hito (escuela asignada a Hito 3)</td>
                </tr>
              `}
              ${h3Data && h4Data ? `
                <tr style="background-color: #f1f9e9; font-weight: bold;">
                  <td style="padding: 3px 5px; border: 1px solid #cbd5e1; color: #3f6515;">Consolidado Combinado (10 preguntas cerradas)</td>
                  <td style="padding: 3px 5px; border: 1px solid #cbd5e1; text-align: center;">${totPre} / 10 pts</td>
                  <td style="padding: 3px 5px; border: 1px solid #cbd5e1; text-align: center; color: #52801b;">${totPost} / 10 pts</td>
                  <td style="padding: 3px 5px; border: 1px solid #cbd5e1; text-align: center; color: ${totDelta >= 0 ? '#52801b' : '#e11d48'};">${deltaSign}${totDelta} pts</td>
                  <td style="padding: 3px 5px; border: 1px solid #cbd5e1; text-align: center; color: #3f6515; font-weight: bold;">${p.status === 'Retrocedió' ? 'Requiere refuerzo' : (p.status === 'Mantuvo' ? 'Consolidó' : p.status)}</td>
                </tr>
              ` : ''}
            </tbody>
          </table>
        </div>

        <!-- 2. ORIENTACIÓN PARA EL PROFESOR/A LÍDER (NUEVA POSICIÓN COMO PUNTO 2) -->
        <div style="margin-bottom: 8px; padding: 6px 8px; background-color: #f1f9e9; border: 1px solid #cbe99c; border-radius: 6px;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 2px;">
            <div style="font-size: 10.5px; font-weight: bold; color: #3f6515; text-transform: uppercase;">2. Orientación Pedagógica y Plan de Acción Docente</div>
            <span style="font-size: 8px; background-color: #ffffff; color: #52801b; padding: 1px 6px; border-radius: 8px; font-weight: bold; border: 1px solid #cbe99c;">Metodología VERDICAL</span>
          </div>
          <p style="font-size: 9px; color: #14532d; margin: 0 0 3px 0; line-height: 1.3;">
            <strong>Diagnóstico Curricular:</strong> ${orientacionDocente}
          </p>
          <p style="font-size: 9px; color: #14532d; margin: 0 0 4px 0; line-height: 1.3;">
            <strong>Actividad Formativa Sugerida:</strong> ${orientacionActividad}
          </p>
          <div style="background-color: #ffffff; padding: 4px 6px; border-radius: 4px; border: 1px solid #d9f0b5; font-size: 8.5px; color: #334155;">
            <strong style="color: #274209;">Recursos y Guías en la Biblioteca del Dashboard:</strong>
            <span style="display: inline-block; margin-left: 4px;">
              ${orientacionGuias.map(g => `<a href="https://badijoly.github.io/pac-linea-verde/biblioteca/guias/${g.file}" target="_blank" style="color: #00737d; text-decoration: underline; font-weight: bold; margin-right: 6px;">📄 ${g.name} ↗</a>`).join('')}
              <a href="https://badijoly.github.io/pac-linea-verde/#biblioteca" target="_blank" style="color: #52801b; text-decoration: underline; font-weight: bold;">[Ver Biblioteca Completa ↗]</a>
            </span>
          </div>
        </div>

        <!-- 3. DESGLOSE DE PREGUNTAS CERRADAS (PUNTO 3) -->
        <div style="margin-bottom: 0px;">
          <div style="font-size: 10.5px; font-weight: bold; text-transform: uppercase; color: #0f172a; margin-bottom: 3px; border-left: 3px solid #88C23F; padding-left: 6px;">3. Desglose de Preguntas Cerradas (Alternativas)</div>
          
          ${h3Data ? `
            <div style="margin-bottom: 5px;">
              <div style="font-size: 9.5px; font-weight: bold; color: #52801b; margin-bottom: 2px;">Hito 3: Los otros vecinos de Antofagasta (Biodiversidad y Humedales) · Puntaje: ${h3Data.post_score} / 5</div>
              <table style="width: 100%; border-collapse: collapse; font-size: 9px; table-layout: fixed;">
                <thead>
                  <tr style="background-color: #f8fafc;">
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 25px; text-align: center;">#</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1;">Pregunta Evaluada</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 35px; text-align: center;">PRE</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 35px; text-align: center;">POST</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 45px; text-align: center;">Correcta</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 85px; text-align: center;">Evolución</th>
                  </tr>
                </thead>
                <tbody>
                  ${h3MCRows}
                </tbody>
              </table>
            </div>
          ` : ''}

          ${h4Data ? `
            <div style="margin-bottom: 0px;">
              <div style="font-size: 9.5px; font-weight: bold; color: #00737d; margin-bottom: 2px;">Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos Limpios) · Puntaje: ${h4Data.post_score} / 5</div>
              <table style="width: 100%; border-collapse: collapse; font-size: 9px; table-layout: fixed;">
                <thead>
                  <tr style="background-color: #f8fafc;">
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 25px; text-align: center;">#</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1;">Pregunta Evaluada</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 35px; text-align: center;">PRE</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 35px; text-align: center;">POST</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 45px; text-align: center;">Correcta</th>
                    <th style="padding: 2.5px 4px; border: 1px solid #cbd5e1; width: 85px; text-align: center;">Evolución</th>
                  </tr>
                </thead>
                <tbody>
                  ${h4MCRows}
                </tbody>
              </table>
            </div>
          ` : ''}
        </div>
      </div>

      <!-- PÁGINA 2: Salto limpio para Punto 4 (Preguntas Abiertas) + Pie Institucional -->
      <div style="page-break-before: always; break-before: page; padding-top: 4px;">
        
        <!-- 4. DESGLOSE DE PREGUNTAS ABIERTAS (PUNTO 4) -->
        <div style="margin-bottom: 8px;">
          <div style="font-size: 10.5px; font-weight: bold; text-transform: uppercase; color: #0f172a; margin-bottom: 4px; border-left: 3px solid #88C23F; padding-left: 6px;">4. Desglose de Preguntas Abiertas (Análisis Cualitativo de Apropiación)</div>
          
          ${h3Data ? `
            <div style="margin-bottom: 6px;">
              <div style="font-size: 9.5px; font-weight: bold; color: #52801b; margin-bottom: 3px;">Hito 3: Los otros vecinos de Antofagasta (Biodiversidad y Ecosistemas)</div>
              ${h3OpenCards}
            </div>
          ` : ''}

          ${h4Data ? `
            <div style="margin-bottom: 6px;">
              <div style="font-size: 9.5px; font-weight: bold; color: #00737d; margin-bottom: 3px;">Hito 4: Proteger nuestro escudo invisible (Atmósfera y Cielos Limpios)</div>
              ${h4OpenCards}
            </div>
          ` : ''}
        </div>

        <!-- Pie de Página Institucional -->
        <table style="width: 100%; border-top: 1px solid #e2e8f0; padding-top: 5px; font-size: 8px; color: #94a3b8; table-layout: fixed;">
          <tr>
            <td style="width: 50%;">Herramienta desarrollada por <strong>rodolfo.tapia@verdical.cl</strong> para <strong>Grupo Verdical SPA</strong></td>
            <td style="width: 50%; text-align: right;">PAC Línea Verde 2026 · Metodología VERDICAL · Antofagasta, Chile</td>
          </tr>
        </table>

      </div>

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


// ==========================================
// FUNCIONALIDADES DE LA BIBLIOTECA DE RECURSOS
// ==========================================
function filterBiblioteca(category) {
  const cards = document.querySelectorAll('.biblioteca-card');
  const pills = document.querySelectorAll('.biblioteca-pill');

  pills.forEach(p => {
    if (p.getAttribute('data-cat') === category) {
      p.className = 'biblioteca-pill px-3.5 py-1.5 rounded-full text-xs font-bold bg-[#88C23F] text-slate-950 shadow-sm transition';
    } else {
      p.className = 'biblioteca-pill px-3.5 py-1.5 rounded-full text-xs font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 hover:text-slate-900 transition';
    }
  });

  cards.forEach(card => {
    const cardCat = card.getAttribute('data-category');
    if (category === 'all' || cardCat === category || (category === 'guias' && cardCat === 'guias')) {
      card.classList.remove('hidden');
    } else {
      card.classList.add('hidden');
    }
  });
}

function searchBiblioteca(query) {
  const q = (query || '').toLowerCase().trim();
  const cards = document.querySelectorAll('.biblioteca-card');

  cards.forEach(card => {
    const title = (card.getAttribute('data-title') || '').toLowerCase();
    const desc = (card.getAttribute('data-desc') || '').toLowerCase();
    const tags = (card.getAttribute('data-tags') || '').toLowerCase();
    if (!q || title.includes(q) || desc.includes(q) || tags.includes(q)) {
      card.classList.remove('hidden');
    } else {
      card.classList.add('hidden');
    }
  });
}
