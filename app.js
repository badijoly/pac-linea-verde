/**
 * Lógica principal del Dashboard Interactivo de Evaluación Línea Verde - PAC 2026
 */

let RAW_DATA = null;
let CURRENT_HITO = 'Hito 4'; // Default to Hito 4 (Escudo invisible)
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
    // Intentar primero desde el servidor API local con timestamp para evitar caché
    const response = await fetch(`/api/data?t=${Date.now()}`);
    if (!response.ok) throw new Error('API local no respondió');
    RAW_DATA = await response.json();
    syncText.innerText = 'Conectado a Excel (En vivo)';
    syncStatus.className = 'flex items-center gap-1.5 text-xs text-emerald-800 bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200';
  } catch (err) {
    console.warn('Fallo API local, intentando cargar data.json estático...', err);
    try {
      const respStatic = await fetch(`./data.json?t=${Date.now()}`);
      if (!respStatic.ok) throw new Error('data.json no disponible');
      RAW_DATA = await respStatic.json();
      syncText.innerText = 'Datos locales cargados (data.json)';
      syncStatus.className = 'flex items-center gap-1.5 text-xs text-amber-800 bg-amber-50 px-3 py-1.5 rounded-lg border border-amber-200';
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
    const stripAcc = (s) => s ? s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim() : '';

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
      const first = stripAcc(name).split(' ')[0] || '';
      if (femaleNames.has(first)) return 'Femenino';
      if (first.endsWith('a')) return 'Femenino';
      return 'Masculino';
    };

    const evalOpenText = (text, hito) => {
      if (!text || text.trim().length < 3) return { level: 0, label: 'Sin respuesta', text: '' };
      const tc = stripAcc(text);
      if (['no se', 'nada', 'ninguna', 'no me acuerdo', 'nose'].some(b => tc.includes(b))) {
        return { level: 1, label: 'Básico / Desconocimiento', text };
      }
      const kw = hito === 'Hito 3' ? 
        ['ecosistema', 'biodiversidad', 'gaviotin', '50 gr', '50 gramos', 'microorganismo', 'flora', 'fauna', 'salina', 'desierto costero', 'cadena trofica', 'humedal', 'reserva', 'portada', 'endemica', 'migratoria'] :
        ['atmosfera', 'escudo', 'capa de ozono', 'gases', 'efecto invernadero', 'radiacion', 'uv', 'contaminacion luminica', 'astronomia', 'telescopio', 'paranal', 'luz azul', 'cielos limpios', 'patrimonio'];
      const matches = kw.filter(k => tc.includes(k)).length;
      if (matches >= 2 || tc.split(/\s+/).length >= 10) return { level: 3, label: 'Apropiación Científica', text };
      if (matches === 1 || tc.split(/\s+/).length >= 4) return { level: 2, label: 'Comprensión Inicial', text };
      return { level: 1, label: 'Básico / Desconocimiento', text };
    };

    // 1. Extraer Directorio de Establecimientos
    const schools = {};
    if (wb.SheetNames.includes('Participantes 2026')) {
      const sheet = wb.Sheets['Participantes 2026'];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[1]) continue;
        let sch = cleanStr(row[1]);
        if (sch.includes('Sabella')) sch = 'Liceo Andrés Sabella';
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
        let sch = cleanStr(row[1]);
        if (sch.includes('Sabella')) sch = 'Liceo Andrés Sabella';
        if (schools[sch]) {
          if (!schools[sch].lider1) {
            schools[sch].lider1 = cleanStr(row[2]);
            schools[sch].tel1 = cleanStr(row[3]);
            schools[sch].email1 = cleanStr(row[4]);
          }
          if (!schools[sch].lider2) {
            schools[sch].lider2 = cleanStr(row[5]);
            schools[sch].tel2 = cleanStr(row[6]);
            schools[sch].email2 = cleanStr(row[7]);
          }
        }
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
        let school = cleanStr(row[2]);
        if (school.includes('Sabella')) school = 'Liceo Andrés Sabella';
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
          const status = delta > 0 ? 'Mejoró' : (delta === 0 ? 'Mantuvo' : 'Retrocedió');

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
          for (let i = 0; i < 4; i++) {
            const tPre = d.PRE.open[i];
            const tPost = d.POST.open[i];
            const evPre = evalOpenText(tPre, hitoName);
            const evPost = evalOpenText(tPost, hitoName);
            openAnalysis.push({
              question: openTitles[i],
              pre_text: tPre,
              post_text: tPost,
              pre_eval: evPre,
              post_eval: evPost,
              level_delta: evPost.level - evPre.level
            });
          }

          const pObj = {
            name,
            role: d.role,
            school: sch,
            gender,
            pre_score: preScore,
            post_score: postScore,
            delta,
            status,
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
          certificacion: meta.certificacion || 'No'
        };
      });
      schoolRankings.sort((a, b) => b.delta_avg - a.delta_avg);

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
      h.schools.forEach(s => schoolsSet.add(s.school));
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

  // Dynamic legend for donut
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
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11 } } },
        tooltip: {
          callbacks: {
            title: (items) => {
              const idx = items[0].dataIndex;
              return `Q${idx + 1}: ${questionsMeta[idx].title.substring(0, 60)}...`;
            },
            label: (item) => `${item.dataset.label}: ${item.raw}%`
          }
        }
      }
    }
  });

  // 3. Schools Delta Bar Chart
  const schoolsMeta = RAW_DATA.hitos[hitoKey].schools;
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

    const deltaSign = s.delta_avg >= 0 ? '+' : '';
    const deltaColor = s.delta_avg > 0 ? 'text-emerald-700 font-bold' : (s.delta_avg === 0 ? 'text-slate-600' : 'text-rose-600 font-bold');

    tr.innerHTML = `
      <td class="py-3 px-4 font-semibold text-slate-800">${s.school}</td>
      <td class="py-3 px-3">
        <span class="px-2 py-0.5 rounded text-[11px] bg-slate-100 text-slate-700 font-medium">${s.dependencia}</span>
      </td>
      <td class="py-3 px-3 text-center font-bold text-slate-800">${s.n}</td>
      <td class="py-3 px-3 text-center text-slate-600">${s.pre_avg.toFixed(2)}</td>
      <td class="py-3 px-3 text-center text-emerald-700 font-semibold">${s.post_avg.toFixed(2)}</td>
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
        <div class="text-slate-800 font-medium">${s.lider1 || 'Sin registrar'}</div>
        <div class="text-[10px] text-slate-400">${s.email1 || ''} ${s.tel1 ? '· Tel: ' + s.tel1 : ''}</div>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// 3. TAB: PREGUNTAS Y DISTRACTORES
function renderQuestionsTab() {
  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const questions = RAW_DATA.hitos[hitoKey].questions;
  const container = document.getElementById('questions-detail-container');
  container.innerHTML = '';

  questions.forEach((q, i) => {
    const card = document.createElement('div');
    card.className = 'border border-slate-200 rounded-xl p-4 bg-slate-50/50 hover:bg-white transition space-y-3';

    // Distractor pills
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

// 4. TAB: CUALITATIVO
function renderQualitativeTab() {
  const hitoKey = CURRENT_HITO === 'Ambos' ? 'Hito 4' : CURRENT_HITO;
  const openQuestions = RAW_DATA.hitos[hitoKey].open_questions;
  const selectEl = document.getElementById('select-open-q');

  // Populate open questions select if empty or changed
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
    item.className = 'bg-white p-3.5 rounded-xl border border-slate-200 hover:border-slate-300 transition text-xs space-y-2';

    const deltaBadge = qData.level_delta > 0 ?
      '<span class="text-[10px] font-bold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">Crecimiento conceptual ↗</span>' :
      (qData.level_delta < 0 ? '<span class="text-[10px] font-medium px-2 py-0.5 rounded bg-rose-100 text-rose-800">Menor detalle</span>' : '<span class="text-[10px] text-slate-400">Nivel similar</span>');

    item.innerHTML = `
      <div class="flex items-center justify-between border-b border-slate-100 pb-2">
        <div>
          <span class="font-bold text-slate-900">${p.name}</span>
          <span class="text-slate-400 ml-1">(${p.school})</span>
        </div>
        <div class="flex items-center gap-2">
          ${deltaBadge}
          <span class="font-medium text-[11px] px-2 py-0.5 rounded ${p.status === 'Mejoró' ? 'badge-improved' : (p.status === 'Retrocedió' ? 'badge-regressed' : 'badge-maintained')}">${p.status}</span>
        </div>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-3 pt-1">
        <div class="bg-slate-50 p-2.5 rounded-lg border border-slate-100">
          <span class="text-[10px] font-semibold text-slate-400 block uppercase mb-1">Respuesta PRE:</span>
          <p class="text-slate-700 italic">"${qData.pre_text || '<span class=\"text-slate-400\">Sin respuesta registrada</span>'}"</p>
          <div class="mt-1 text-[10px] font-medium text-slate-500">Nivel basal: ${qData.pre_eval.label}</div>
        </div>
        <div class="bg-emerald-50/40 p-2.5 rounded-lg border border-emerald-100">
          <span class="text-[10px] font-semibold text-emerald-700 block uppercase mb-1">Respuesta POST:</span>
          <p class="text-slate-800 font-medium">"${qData.post_text || '<span class=\"text-slate-400\">Sin respuesta registrada</span>'}"</p>
          <div class="mt-1 text-[10px] font-semibold text-emerald-700">Nivel final: ${qData.post_eval.label}</div>
        </div>
      </div>
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

    // Generar sugerencia personalizada según su comportamiento
    let recomendacion = '';
    let statusClass = 'bg-emerald-50 border-emerald-200 text-emerald-800';

    if (s.delta_avg > 0.7) {
      recomendacion = `<strong>Consolidación y Prototipado:</strong> Este equipo demostró una comprensión sobresaliente de los conceptos de la intervención (+${s.delta_avg} pts). Se sugiere al profesor líder canalizar este entusiasmo directamente al diseño de soluciones en la fase <em>Diseñador</em> de VERDICAL (ej. prototipos solares o luminarias apantalladas).`;
      statusClass = 'bg-emerald-50 border-emerald-200 text-emerald-800';
    } else if (s.delta_avg >= 0.2) {
      recomendacion = `<strong>Refuerzo de conceptos intermedios:</strong> El equipo tiene una asimilación positiva pero heterogénea (+${s.delta_avg} pts). Se recomienda realizar un plenario corto de 15 minutos repasando las diferencias entre efecto invernadero natural y antropogénico, antes de la entrega final.`;
      statusClass = 'bg-amber-50 border-amber-200 text-amber-800';
    } else {
      recomendacion = `<strong>Atención a 'Efecto Techo' y Distractores:</strong> El puntaje basal fue muy elevado o se registraron confusiones sutiles en las opciones cerradas (${s.delta_avg} pts). No interpretar esto como falta de aprendizaje: cualitativamente los estudiantes manejan la terminología. Se aconseja pedirles que expliquen con sus propias palabras el impacto en la fauna local para afianzar la seguridad en sus conocimientos.`;
      statusClass = 'bg-slate-50 border-slate-200 text-slate-800';
    }

    card.innerHTML = `
      <div class="flex items-center justify-between border-b border-slate-100 pb-2">
        <div>
          <h4 class="font-bold text-slate-900 text-sm">${s.school}</h4>
          <span class="text-[11px] text-slate-500">Líder: ${s.lider1 || 'Coordinador del equipo'}</span>
        </div>
        <div class="text-right">
          <span class="text-xs font-bold ${s.delta_avg >= 0 ? 'text-emerald-600' : 'text-rose-600'}">Delta: ${s.delta_avg >= 0 ? '+' : ''}${s.delta_avg}</span>
          <div class="text-[10px] text-slate-400">${s.pct_mejora}% mejoró</div>
        </div>
      </div>

      <div class="p-3 rounded-lg border text-xs leading-relaxed ${statusClass}">
        ${recomendacion}
      </div>

      <div class="flex items-center justify-between text-[11px] text-slate-500 pt-1">
        <span>Estudiantes evaluados: <strong>${s.n}</strong></span>
        <span>${s.email1 ? '✉ ' + s.email1 : ''}</span>
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
  document.getElementById('modal-school').innerText = `${p.school} · Género: ${p.gender}`;
  document.getElementById('modal-role').innerText = p.role || 'Estudiante';

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

  p.open_analysis.forEach((oq, idx) => {
    const oRow = document.createElement('div');
    oRow.className = 'p-3 rounded-lg border border-slate-200 bg-slate-50 space-y-2';

    oRow.innerHTML = `
      <div class="font-bold text-slate-800 text-xs">${oq.question}</div>
      <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs">
        <div class="bg-white p-2 rounded border border-slate-200">
          <span class="text-[10px] font-bold text-slate-400 block uppercase">PRE:</span>
          <p class="italic text-slate-700">"${oq.pre_text || 'Sin respuesta'}"</p>
        </div>
        <div class="bg-white p-2 rounded border border-emerald-200">
          <span class="text-[10px] font-bold text-emerald-700 block uppercase">POST:</span>
          <p class="font-medium text-slate-800">"${oq.post_text || 'Sin respuesta'}"</p>
        </div>
      </div>
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
    csv += `"${s.school}","${s.dependencia}",${s.n},${s.pre_avg},${s.post_avg},${s.delta_avg},${s.pct_mejora}%,"${s.lider1}","${s.email1}","${s.tel1}"\n`;
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
