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
  document.getElementById('btn-reload').addEventListener('click', () => {
    loadDashboardData(true);
  });

  const fileInput = document.getElementById('file-input');
  fileInput.addEventListener('change', handleFileUpload);
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
      syncText.innerText = 'Error al cargar datos. Sube un archivo .xlsx';
      syncStatus.className = 'flex items-center gap-1.5 text-xs text-rose-800 bg-rose-50 px-3 py-1.5 rounded-lg border border-rose-200';
      return;
    }
  }

  populateFilterDropdowns();
  updateUI();
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
    el.classList.remove('border-emerald-600', 'text-emerald-700', 'font-semibold');
    el.classList.add('border-transparent', 'text-slate-600', 'font-medium');
  });

  const targetView = document.getElementById(`view-${tabId}`);
  const targetBtn = document.getElementById(`tab-btn-${tabId}`);
  if (targetView) targetView.classList.remove('hidden');
  if (targetBtn) {
    targetBtn.classList.remove('border-transparent', 'text-slate-600', 'font-medium');
    targetBtn.classList.add('border-emerald-600', 'text-emerald-700', 'font-semibold');
  }

  updateUI();
  if (window.lucide) lucide.createIcons();
}

// Cambiar Hito seleccionado
function setHito(hito) {
  CURRENT_HITO = hito;

  const btn4 = document.getElementById('btn-hito-4');
  const btn3 = document.getElementById('btn-hito-3');
  const btnAmbos = document.getElementById('btn-hito-ambos');

  [btn4, btn3, btnAmbos].forEach(btn => {
    btn.className = 'px-3 py-1.5 rounded-lg font-medium text-slate-600 hover:text-slate-900 transition';
  });

  if (hito === 'Hito 4') {
    btn4.className = 'px-3 py-1.5 rounded-lg font-semibold bg-white text-emerald-800 shadow-sm transition';
  } else if (hito === 'Hito 3') {
    btn3.className = 'px-3 py-1.5 rounded-lg font-semibold bg-white text-emerald-800 shadow-sm transition';
  } else {
    btnAmbos.className = 'px-3 py-1.5 rounded-lg font-semibold bg-white text-emerald-800 shadow-sm transition';
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

  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const hitoData = RAW_DATA.hitos[hitoKey];
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
    descEl.innerText = 'Evaluación del aprendizaje sobre la biodiversidad del desierto costero, el Humedal La Chimba, el Gaviotín Chico (Sternula lorata), cadenas tróficas y especies descomponedoras.';
    const h3 = RAW_DATA.hitos['Hito 3'];
    if (h3) hakeEl.innerText = `g = +${h3.hake_gain}`;
  } else {
    titleEl.innerText = 'Comparativa Consolidada: Hito 3 vs Hito 4';
    descEl.innerText = 'Contraste entre el diagnóstico basal de biodiversidad y humedales (Hito 3) versus la atmósfera y contaminación lumínica (Hito 4).';
    hakeEl.innerText = 'Global PAC';
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
  deltaEl.className = avgDelta >= 0 ? 'text-2xl font-bold text-emerald-600 mt-1' : 'text-2xl font-bold text-rose-600 mt-1';

  pctMejoraEl.innerText = `${pctMejora}%`;
  countMejoraEl.innerText = `${nMejora} participantes subieron`;

  pctEstancadosEl.innerText = `${pctEstancados}%`;
  countEstancadosEl.innerText = `${nMantuvo} mantuvieron, ${nRetrocedio} bajaron`;
}

// RENDER CHARTS CON CHART.JS
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
        backgroundColor: ['#10b981', '#94a3b8', '#f43f5e'],
        borderWidth: 2,
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
      cutout: '70%'
    }
  });

  const legendEl = document.getElementById('legend-status');
  legendEl.innerHTML = `
    <div>
      <span class="inline-block w-2.5 h-2.5 rounded-full bg-emerald-500 mr-1"></span>
      <span class="font-semibold text-slate-700">${nMejora}</span>
      <div class="text-[10px] text-slate-400">Mejoró (${N ? ((nMejora/N)*100).toFixed(0) : 0}%)</div>
    </div>
    <div>
      <span class="inline-block w-2.5 h-2.5 rounded-full bg-slate-400 mr-1"></span>
      <span class="font-semibold text-slate-700">${nMantuvo}</span>
      <div class="text-[10px] text-slate-400">Mantuvo (${N ? ((nMantuvo/N)*100).toFixed(0) : 0}%)</div>
    </div>
    <div>
      <span class="inline-block w-2.5 h-2.5 rounded-full bg-rose-500 mr-1"></span>
      <span class="font-semibold text-slate-700">${nRetrocedio}</span>
      <div class="text-[10px] text-slate-400">Retrocedió (${N ? ((nRetrocedio/N)*100).toFixed(0) : 0}%)</div>
    </div>
  `;

  // 2. Question Accuracy Bar Chart
  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const questionsMeta = RAW_DATA.hitos[hitoKey].questions;
  const key = RAW_DATA.hitos[hitoKey].questions.map(q => q.correct_key.toLowerCase());

  const prePcts = [];
  const postPcts = [];
  const labels = ['Pregunta 1', 'Pregunta 2', 'Pregunta 3', 'Pregunta 4', 'Pregunta 5'];

  for (let i = 0; i < 5; i++) {
    const preC = filtered.filter(p => p.pre_mc[i] === key[i]).length;
    const postC = filtered.filter(p => p.post_mc[i] === key[i]).length;
    prePcts.push(N ? ((preC / N) * 100).toFixed(1) : 0);
    postPcts.push(N ? ((postC / N) * 100).toFixed(1) : 0);
  }

  const ctxQuestions = document.getElementById('chart-questions-bar').getContext('2d');
  if (CHARTS['questions']) CHARTS['questions'].destroy();

  CHARTS['questions'] = new Chart(ctxQuestions, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: '% Aciertos PRE',
          data: prePcts,
          backgroundColor: '#94a3b8',
          borderRadius: 6
        },
        {
          label: '% Aciertos POST',
          data: postPcts,
          backgroundColor: '#059669',
          borderRadius: 6
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
          ticks: { callback: v => v + '%' }
        }
      },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11 } } }
      }
    }
  });

  // 3. Schools Delta Bar Chart
  const schoolsMeta = RAW_DATA.hitos[hitoKey].schools.filter(s => s.has_data !== false);
  const schoolLabels = schoolsMeta.map(s => s.school.replace('Escuela ', 'Esc. ').replace('Colegio ', 'Col. '));
  const schoolDeltas = schoolsMeta.map(s => s.delta_avg);

  const ctxSchools = document.getElementById('chart-schools-delta').getContext('2d');
  if (CHARTS['schools']) CHARTS['schools'].destroy();

  CHARTS['schools'] = new Chart(ctxSchools, {
    type: 'bar',
    data: {
      labels: schoolLabels,
      datasets: [{
        label: 'Delta Promedio (Puntos)',
        data: schoolDeltas,
        backgroundColor: schoolDeltas.map(d => d >= 0 ? '#10b981' : '#f43f5e'),
        borderRadius: 6
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
          ticks: { callback: v => (v >= 0 ? '+' : '') + v }
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
          backgroundColor: '#cbd5e1',
          borderRadius: 6
        },
        {
          label: 'Puntaje POST',
          data: gPost,
          backgroundColor: '#0d9488',
          borderRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: { beginAtZero: true, max: 5 }
      },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11 } } }
      }
    }
  });
}

// 2. TAB: ESTABLECIMIENTOS TABLE
function renderSchoolsTab() {
  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const schools = RAW_DATA.hitos[hitoKey].schools;
  const tbody = document.getElementById('schools-table-body');
  tbody.innerHTML = '';

  schools.forEach(s => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/80 transition';

    if (s.has_data === false || s.n === 0) {
      tr.innerHTML = `
        <td class="py-3 px-4 font-semibold text-slate-800">${s.school}</td>
        <td class="py-3 px-3">
          <span class="px-2 py-0.5 rounded text-[11px] bg-slate-100 text-slate-700 font-medium">${s.dependencia}</span>
        </td>
        <td class="py-3 px-3 text-center text-slate-400 italic font-semibold">0 (Sin datos)</td>
        <td class="py-3 px-3 text-center text-slate-400">-</td>
        <td class="py-3 px-3 text-center text-slate-400">-</td>
        <td class="py-3 px-3 text-center">
          <span class="px-2 py-0.5 rounded text-[10px] bg-amber-50 text-amber-800 border border-amber-200 font-medium">Sin datos en este hito</span>
        </td>
        <td class="py-3 px-3 text-center text-slate-400">-</td>
        <td class="py-3 px-3 text-center text-slate-400 text-[10px]">No evaluado</td>
        <td class="py-3 px-4">
          <div class="text-slate-800 font-medium">${s.lider1 || 'Coordinador del equipo'}</div>
          <div class="text-[10px] text-slate-400">${s.email1 || ''} ${s.tel1 ? '· Tel: ' + s.tel1 : ''}</div>
        </td>
      `;
    } else {
      const deltaSign = s.delta_avg >= 0 ? '+' : '';
      const deltaColor = s.delta_avg > 0 ? 'text-emerald-700 font-bold' : (s.delta_avg === 0 ? 'text-slate-600' : 'text-rose-600 font-bold');

      tr.innerHTML = `
        <td class="py-3 px-4 font-semibold text-slate-800">${s.school}</td>
        <td class="py-3 px-3">
          <span class="px-2 py-0.5 rounded text-[11px] bg-slate-100 text-slate-700 font-medium">${s.dependencia}</span>
        </td>
        <td class="py-3 px-3 text-center font-bold text-slate-800">${s.n}</td>
        <td class="py-3 px-3 text-center text-slate-600">${s.pre_avg !== null ? s.pre_avg.toFixed(2) : '-'}</td>
        <td class="py-3 px-3 text-center text-emerald-700 font-semibold">${s.post_avg !== null ? s.post_avg.toFixed(2) : '-'}</td>
        <td class="py-3 px-3 text-center ${deltaColor}">${deltaSign}${s.delta_avg.toFixed(2)}</td>
        <td class="py-3 px-3 text-center">
          <span class="px-2 py-0.5 rounded-full text-[11px] font-bold ${s.pct_mejora >= 50 ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-100 text-slate-700'}">
            ${s.pct_mejora}%
          </span>
        </td>
        <td class="py-3 px-3">
          <div class="flex items-center gap-1 text-[10px]">
            <span class="text-emerald-700 font-medium" title="Mejoró">${s.status_counts['Mejoró'] || 0}↗</span>
            <span class="text-slate-400">·</span>
            <span class="text-slate-600" title="Mantuvo">${s.status_counts['Mantuvo'] || 0}=</span>
            <span class="text-slate-400">·</span>
            <span class="text-rose-600 font-medium" title="Retrocedió">${s.status_counts['Retrocedió'] || 0}↘</span>
          </div>
        </td>
        <td class="py-3 px-4">
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
  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const questions = RAW_DATA.hitos[hitoKey].questions;
  const container = document.getElementById('questions-detail-container');
  container.innerHTML = '';

  questions.forEach((q) => {
    const card = document.createElement('div');
    card.className = 'border border-slate-200 rounded-xl p-4 bg-slate-50/50 hover:bg-white transition space-y-3';

    const distractorPills = Object.entries(q.post_distractors)
      .map(([opt, count]) => `<span class="px-2 py-0.5 rounded bg-rose-50 text-rose-700 border border-rose-200 font-medium text-[11px]">Opción ${opt}: ${count} veces</span>`)
      .join(' ') || '<span class="text-slate-400 text-xs">Sin errores registrados</span>';

    const diffBadge = q.diff >= 0 ? 
      `<span class="badge-improved text-xs px-2.5 py-0.5 rounded-full font-bold">+${q.diff} aciertos (+${(q.post_pct - q.pre_pct).toFixed(1)}%)</span>` :
      `<span class="badge-regressed text-xs px-2.5 py-0.5 rounded-full font-bold">${q.diff} aciertos (${(q.post_pct - q.pre_pct).toFixed(1)}%)</span>`;

    card.innerHTML = `
      <div class="flex flex-wrap items-start justify-between gap-2">
        <div class="space-y-1">
          <div class="flex items-center gap-2">
            <span class="w-6 h-6 rounded-full bg-emerald-600 text-white font-bold text-xs flex items-center justify-center">${q.num}</span>
            <span class="text-xs font-bold uppercase tracking-wider text-slate-500">Alternativa Correcta: [ ${q.correct_key} ]</span>
          </div>
          <h4 class="text-sm font-bold text-slate-800 pl-8">${q.title}</h4>
        </div>
        <div>${diffBadge}</div>
      </div>

      <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-white p-3 rounded-lg border border-slate-200 text-xs pl-8">
        <div>
          <span class="text-slate-400 text-[10px] block">Aciertos PRE</span>
          <span class="font-bold text-slate-700 text-sm">${q.pre_correct} (${q.pre_pct}%)</span>
        </div>
        <div>
          <span class="text-slate-400 text-[10px] block">Aciertos POST</span>
          <span class="font-bold text-emerald-700 text-sm">${q.post_correct} (${q.post_pct}%)</span>
        </div>
        <div>
          <span class="text-slate-400 text-[10px] block">Mantuvieron Correcto</span>
          <span class="font-bold text-slate-700 text-sm">${q.transitions['Mantuvo Correcto'] || 0}</span>
        </div>
        <div>
          <span class="text-slate-400 text-[10px] block">Retrocedieron (de correcta a errada)</span>
          <span class="font-bold text-rose-600 text-sm">${q.transitions['Retrocedió'] || 0}</span>
        </div>
      </div>

      <div class="pl-8 pt-1 text-xs">
        <span class="font-medium text-slate-600 mr-2">Distractores más seleccionados en el POST:</span>
        ${distractorPills}
      </div>
    `;
    container.appendChild(card);
  });
}

// 4. TAB: CUALITATIVO RIGUROSO
function renderQualitativeTab() {
  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const openQuestions = RAW_DATA.hitos[hitoKey].open_questions;
  const selectEl = document.getElementById('select-open-q');

  if (selectEl.children.length === 0 || selectEl.dataset.hito !== hitoKey) {
    selectEl.innerHTML = '';
    openQuestions.forEach((oq, idx) => {
      const opt = document.createElement('option');
      opt.value = idx;
      opt.textContent = `P${idx + 1}: ${oq.substring(0, 65)}...`;
      selectEl.appendChild(opt);
    });
    selectEl.dataset.hito = hitoKey;
  }

  const selectedIdx = parseInt(selectEl.value || 0);
  const filtered = getFilteredParticipants();
  const listEl = document.getElementById('qualitative-responses-list');
  listEl.innerHTML = '';

  filtered.forEach(p => {
    const qData = p.open_analysis[selectedIdx];
    if (!qData) return;

    const item = document.createElement('div');
    item.className = 'bg-white p-4 rounded-xl border border-slate-200 hover:border-slate-300 transition text-xs space-y-2.5';

    // Determinar badge según estado CUALITATIVO de esta pregunta específica
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
      `<div class="mt-2 text-[11px] font-semibold text-amber-800 bg-amber-50 px-2.5 py-1.5 rounded-lg border border-amber-200 flex items-center gap-1.5">
        <i data-lucide="alert-triangle" class="w-3.5 h-3.5 text-amber-600 shrink-0"></i>
        <span>Elemento de evaluación no concluyente por falta de respuesta escrita. Se sugiere revisión individual con el profesor/a líder.</span>
      </div>` : '';

    item.innerHTML = `
      <div class="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-2">
        <div>
          <span class="font-bold text-slate-900 text-sm">${p.name}</span>
          <span class="text-slate-500 ml-1.5">· ${p.school}</span>
        </div>
        <div class="flex items-center gap-2">
          <span class="px-2.5 py-0.5 rounded text-[11px] font-semibold ${badgeClass}">${badgeText}</span>
        </div>
      </div>

      <div class="text-[11px] text-slate-500 font-medium">
        Diagnóstico conceptual: <span class="text-slate-700">${qData.detail}</span>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-3 pt-1">
        <div class="bg-slate-50 p-3 rounded-lg border border-slate-200/80">
          <span class="text-[10px] font-bold text-slate-500 block uppercase tracking-wide mb-1">Respuesta PRE (Inicial):</span>
          <p class="text-slate-700 italic">${qData.pre_text && qData.pre_text.trim() ? '"' + qData.pre_text + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta o no sabe)</span>'}</p>
        </div>
        <div class="bg-emerald-50/30 p-3 rounded-lg border border-emerald-100">
          <span class="text-[10px] font-bold text-emerald-800 block uppercase tracking-wide mb-1">Respuesta POST (Final en terreno):</span>
          <p class="text-slate-800 font-medium">${qData.post_text && qData.post_text.trim() ? '"' + qData.post_text + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta o no sabe)</span>'}</p>
        </div>
      </div>

      ${reviewAlert}
    `;
    listEl.appendChild(item);
  });
}

// 5. TAB: CASOS CRÍTICOS & PEDAGÓGICO
function renderPedagogicalTab() {
  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const schools = RAW_DATA.hitos[hitoKey].schools;
  const container = document.getElementById('pedagogical-cards');
  container.innerHTML = '';

  schools.forEach(s => {
    const card = document.createElement('div');
    card.className = 'border border-slate-200 rounded-xl p-4 bg-white shadow-sm space-y-3';

    let recomendacion = '';
    let statusClass = 'bg-emerald-50 border-emerald-200 text-emerald-800';

    if (s.has_data === false || s.n === 0) {
      recomendacion = `<strong>Sin evaluaciones registradas en este hito:</strong> Este establecimiento no registra respuestas PRE ni POST para este desafío. Se sugiere coordinar directamente con el profesor/a líder (<strong>${s.lider1 || 'Coordinador'}</strong>) para verificar la ejecución de la actividad en terreno o aplicar la evaluación de manera diferida.`;
      statusClass = 'bg-amber-50 border-amber-200 text-amber-900';
    } else if (s.delta_avg > 0.7) {
      recomendacion = `<strong>Consolidación y Prototipado:</strong> Este equipo demostró una comprensión sobresaliente de los conceptos de la intervención (+${s.delta_avg} pts). Se sugiere a la líder <strong>${s.lider1 || 'Docente Líder'}</strong> canalizar este entusiasmo directamente al diseño de soluciones en la fase <em>Diseñador</em> de VERDICAL (ej. prototipos de pantallas para luminarias o maquetas de protección de cielos y humedales).`;
      statusClass = 'bg-emerald-50 border-emerald-200 text-emerald-800';
    } else if (s.delta_avg >= 0.2) {
      recomendacion = `<strong>Refuerzo de conceptos intermedios:</strong> El equipo tiene una asimilación positiva pero heterogénea (+${s.delta_avg} pts). Se recomienda a la líder <strong>${s.lider1 || 'Docente Líder'}</strong> realizar un plenario corto de 15 minutos repasando las diferencias entre efecto invernadero natural y emisiones artificiales, antes de la entrega final.`;
      statusClass = 'bg-amber-50 border-amber-200 text-amber-800';
    } else {
      recomendacion = `<strong>Atención a 'Efecto Techo' y Distractores:</strong> El puntaje basal fue muy elevado o se registraron confusiones sutiles en las opciones cerradas (${s.delta_avg} pts). No interpretar esto como falta de aprendizaje: cualitativamente los estudiantes manejan la terminología. Se aconseja pedirles que expliquen con sus propias palabras el impacto en la fauna local para afianzar la seguridad en sus conocimientos.`;
      statusClass = 'bg-slate-50 border-slate-200 text-slate-800';
    }

    card.innerHTML = `
      <div class="flex items-center justify-between border-b border-slate-100 pb-2">
        <div>
          <h4 class="font-bold text-slate-900 text-sm">${s.school}</h4>
          <span class="text-[11px] text-slate-600 font-medium">Líder: ${s.lider1 || 'Coordinador del equipo'} ${s.lider2 ? '/ ' + s.lider2 : ''}</span>
        </div>
        <div class="text-right">
          ${s.has_data !== false ? `
            <span class="text-xs font-bold ${s.delta_avg >= 0 ? 'text-emerald-600' : 'text-rose-600'}">Delta: ${s.delta_avg >= 0 ? '+' : ''}${s.delta_avg}</span>
            <div class="text-[10px] text-slate-400">${s.pct_mejora}% mejoró</div>
          ` : `
            <span class="text-xs font-bold text-amber-700 bg-amber-100 px-2 py-0.5 rounded">Sin datos</span>
          `}
        </div>
      </div>

      <div class="p-3 rounded-lg border text-xs leading-relaxed ${statusClass}">
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

// 6. TAB: FICHAS INDIVIDUALES
function renderParticipantsTab() {
  const filtered = getFilteredParticipants();
  document.getElementById('participants-count-filtered').innerText = filtered.length;

  const tbody = document.getElementById('participants-table-body');
  tbody.innerHTML = '';

  filtered.forEach((p, idx) => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/80 transition cursor-pointer';
    tr.onclick = () => openModal(p);

    const deltaSign = p.delta >= 0 ? '+' : '';
    const badgeClass = p.status === 'Mejoró' ? 'badge-improved' : (p.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained');

    tr.innerHTML = `
      <td class="py-3 px-4 font-semibold text-slate-900 flex items-center gap-2">
        <span class="w-2 h-2 rounded-full ${p.gender === 'Femenino' ? 'bg-pink-400' : 'bg-blue-400'}"></span>
        ${p.name}
      </td>
      <td class="py-3 px-3 text-slate-600">${p.school}</td>
      <td class="py-3 px-3 text-slate-500 text-[11px]">${p.gender}</td>
      <td class="py-3 px-3 text-center text-slate-600 font-medium">${p.pre_score} / 5</td>
      <td class="py-3 px-3 text-center text-emerald-700 font-bold">${p.post_score} / 5</td>
      <td class="py-3 px-3 text-center font-bold ${p.delta >= 0 ? 'text-emerald-600' : 'text-rose-600'}">${deltaSign}${p.delta}</td>
      <td class="py-3 px-3 text-center">
        <span class="px-2 py-0.5 rounded text-[11px] font-semibold ${badgeClass}">${p.status}</span>
        <div class="text-[10px] text-slate-500 mt-0.5 font-medium">${p.real_status || ''}</div>
      </td>
      <td class="py-3 px-4 text-center">
        <button onclick="event.stopPropagation(); openModal(window.currentFiltered[${idx}])" class="px-2.5 py-1 text-[11px] font-medium text-emerald-700 bg-emerald-50 hover:bg-emerald-100 rounded-md border border-emerald-200 transition">
          Ver Ficha
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  window.currentFiltered = filtered;
}

// MODAL INDIVIDUAL PARTICIPANTE
function openModal(p) {
  if (!p) return;

  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const questionsMeta = RAW_DATA.hitos[hitoKey].questions;
  const key = questionsMeta.map(q => q.correct_key);

  document.getElementById('modal-name').innerText = p.name;
  document.getElementById('modal-school').innerText = `${p.school} · Género: ${p.gender} · Rol: ${p.role || 'Estudiante'}`;
  document.getElementById('modal-role').innerText = p.real_status || 'Estudiante';

  document.getElementById('modal-pre-score').innerText = `${p.pre_score} / 5`;
  document.getElementById('modal-post-score').innerText = `${p.post_score} / 5`;
  document.getElementById('modal-delta-score').innerText = `${p.delta >= 0 ? '+' : ''}${p.delta} pts`;

  const badgeEl = document.getElementById('modal-status-badge');
  badgeEl.innerText = p.status;
  badgeEl.className = `mt-1 inline-block font-semibold px-2 py-0.5 rounded text-[11px] ${p.status === 'Mejoró' ? 'badge-improved' : (p.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained')}`;

  // Multiple choice questions list
  const mcList = document.getElementById('modal-mc-list');
  mcList.innerHTML = '';

  p.q_evo.forEach((evo, idx) => {
    const qRow = document.createElement('div');
    const isCorrectPost = p.post_mc[idx] === key[idx].toLowerCase();
    const isCorrectPre = p.pre_mc[idx] === key[idx].toLowerCase();

    qRow.className = `p-2.5 rounded-lg border text-xs flex items-center justify-between ${isCorrectPost ? 'bg-emerald-50/60 border-emerald-200' : 'bg-rose-50/40 border-rose-200'}`;

    qRow.innerHTML = `
      <div class="space-y-0.5 max-w-[70%]">
        <span class="font-bold text-slate-800">Q${idx + 1}: ${questionsMeta[idx].title}</span>
        <div class="text-[11px] text-slate-500">
          Respuesta PRE: <span class="font-bold ${isCorrectPre ? 'text-emerald-700' : 'text-slate-700'}">${p.pre_mc[idx].toUpperCase() || 'N/A'}</span> ➔ 
          Respuesta POST: <span class="font-bold ${isCorrectPost ? 'text-emerald-700' : 'text-rose-600'}">${p.post_mc[idx].toUpperCase() || 'N/A'}</span>
          <span class="text-slate-400">(Correcta: ${key[idx]})</span>
        </div>
      </div>
      <div>
        <span class="text-[10px] font-bold px-2 py-0.5 rounded ${evo === 'Mejoró' ? 'bg-emerald-100 text-emerald-800' : (evo === 'Retrocedió' ? 'bg-rose-100 text-rose-800' : 'bg-slate-100 text-slate-700')}">
          ${evo}
        </span>
      </div>
    `;
    mcList.appendChild(qRow);
  });

  // Open questions list
  const openList = document.getElementById('modal-open-list');
  openList.innerHTML = '';

  p.open_analysis.forEach((oq) => {
    const oRow = document.createElement('div');
    oRow.className = 'p-3 rounded-lg border border-slate-200 bg-slate-50 space-y-2';

    let oBadge = 'badge-maintained';
    if (oq.status === 'Mejoró') oBadge = 'badge-improved';
    else if (oq.status === 'Retrocedió') oBadge = 'badge-regressed';
    else if (oq.status === 'Sin Evidencia') oBadge = 'bg-slate-100 text-slate-600 border border-slate-300';

    const reviewTag = oq.requires_review ? 
      `<div class="text-[10px] text-amber-800 bg-amber-50 p-1.5 rounded border border-amber-200 font-medium">⚠️ Sin evidencia en texto: Requiere revisión individual con el profesor/a líder.</div>` : '';

    oRow.innerHTML = `
      <div class="flex items-center justify-between">
        <span class="font-bold text-slate-800 text-xs">${oq.question}</span>
        <span class="text-[10px] font-semibold px-2 py-0.5 rounded ${oBadge}">${oq.status}</span>
      </div>
      <div class="text-[10px] text-slate-500 font-medium">${oq.detail}</div>
      <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs">
        <div class="bg-white p-2.5 rounded border border-slate-200">
          <span class="text-[10px] font-bold text-slate-400 block uppercase">PRE:</span>
          <p class="italic text-slate-700">${oq.pre_text && oq.pre_text.trim() ? '"' + oq.pre_text + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta)</span>'}</p>
        </div>
        <div class="bg-white p-2.5 rounded border border-emerald-200">
          <span class="text-[10px] font-bold text-emerald-700 block uppercase">POST:</span>
          <p class="font-medium text-slate-800">${oq.post_text && oq.post_text.trim() ? '"' + oq.post_text + '"' : '<span class="text-slate-400 font-normal italic">(Sin respuesta)</span>'}</p>
        </div>
      </div>
      ${reviewTag}
    `;
    openList.appendChild(oRow);
  });

  document.getElementById('participant-modal').classList.remove('hidden');
  if (window.lucide) lucide.createIcons();
}

function closeModal() {
  document.getElementById('participant-modal').classList.add('hidden');
}

// Exportar tabla a CSV
function exportSchoolsTable() {
  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
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
