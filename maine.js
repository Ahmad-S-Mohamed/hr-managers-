/* ============================================================
   HR Analytics Dashboard – maine.js  (v2 – Fixed & Optimised)
   ============================================================ */

'use strict';

// ── Global state ──────────────────────────────────────────────
let DATA = {};
let allEmployees   = [];
let filteredEmployees = [];
let currentPage    = 1;
const PER_PAGE     = 12;
let sortKey        = null;
let sortDir        = 1;
let isDark         = false;
let dataLoaded     = false;          // guard: load Excel only once

// ── Colour palette ────────────────────────────────────────────
const COLORS = [
  '#4361ee','#7209b7','#06d6a0','#f4a261',
  '#0096c7','#ef476f','#118ab2','#ffd166'
];

// Cache every Chart instance so we can destroy before redraw
const chartInstances = {};

// ── Utility: safe number parse ────────────────────────────────
function toNum(val) {
  if (val === null || val === undefined) return 0;
  const s = String(val).replace(/[^\d.]/g, '');
  return parseFloat(s) || 0;
}

// ── Utility: Excel serial → JS Date → year string ────────────
function serialToYear(serial) {
  if (!serial) return null;
  // Numeric serial (Excel date)
  if (typeof serial === 'number') {
    // Excel epoch: 1 Jan 1900 = serial 1  (with leap-year bug)
    const msPerDay = 86400000;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + serial * msPerDay);
    return String(d.getUTCFullYear());
  }
  // String date (e.g. "2022-04-01" or "01/04/2022")
  const match = String(serial).match(/\b(\d{4})\b/);
  return match ? match[1] : null;
}

// ══════════════════════════════════════════════════════════════
//  LOAD EXCEL
// ══════════════════════════════════════════════════════════════
async function loadExcel() {
  if (dataLoaded) return;   // prevent duplicate fetches on tab switches

  showLoadingOverlay(true);

  try {
    const response = await fetch('data.xlsm');

    if (!response.ok) {
      throw new Error(`HTTP ${response.status} – لم يتم العثور على الملف data.xlsm`);
    }

    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });

    // ── Find sheet (case-insensitive fallback) ────────────────
    let sheetName = 'archive';
    if (!workbook.Sheets[sheetName]) {
      sheetName = workbook.SheetNames.find(n =>
        n.toLowerCase().includes('archive') || n.toLowerCase().includes('data')
      ) || workbook.SheetNames[0];
    }

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      throw new Error('لا توجد ورقة بيانات صالحة في الملف');
    }

    const employees = XLSX.utils.sheet_to_json(worksheet, { range: 0, defval: '' });

    if (!employees.length) {
      throw new Error('الملف فارغ أو لا يحتوي على بيانات');
    }

    buildData(employees);
    dataLoaded = true;

    // Build all chart groups
    buildOverviewCharts();
    buildSalaryCharts();
    buildContractCharts();
    buildInsightCharts();
    buildEmployeeTable();

  } catch (err) {
    console.error('LOAD EXCEL ERROR:', err);
    showErrorBanner(err.message);
  } finally {
    showLoadingOverlay(false);
  }
}

// ── Loading overlay ───────────────────────────────────────────
function showLoadingOverlay(show) {
  let overlay = document.getElementById('loading-overlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'loading-overlay';
    overlay.innerHTML = `
      <div class="loading-box">
        <div class="spinner"></div>
        <p style="margin-top:14px;font-size:14px;color:var(--text2)">جارٍ تحميل البيانات…</p>
      </div>`;
    document.body.appendChild(overlay);
  }
  overlay.style.display = show ? 'flex' : 'none';
}

function showErrorBanner(msg) {
  const banner = document.createElement('div');
  banner.className = 'error-banner';
  banner.innerHTML = `⚠ خطأ في تحميل البيانات: ${msg}`;
  document.querySelector('.main')?.prepend(banner);
}

// ══════════════════════════════════════════════════════════════
//  BUILD DATA
// ══════════════════════════════════════════════════════════════
function buildData(employees) {

  // ── Normalise rows ─────────────────────────────────────────
  DATA.employees = employees
    .filter(emp => emp['كود الموظف'] || emp['اسم الموظف'])   // skip blank rows
    .map(emp => {
      const salary = toNum(
        emp['اجمالي الراتب'] ??
        emp['إجمالي الراتب'] ??
        emp['الراتب الإجمالي'] ?? 0
      );

      const status = (
        emp['حالة  الموظف'] ||
        emp['حالة الموظف'] || ''
      ).trim();

      const hireRaw = emp['بداية العقد'] || emp['تاريخ التعيين'] || '';

      return {
        id:              String(emp['كود الموظف'] || ''),
        name:            String(emp['اسم الموظف'] || ''),
        dept:            String(emp['الإدارة'] || 'غير محدد').trim(),
        title:           String(emp['المسمى الوظيفي'] || '').trim(),
        status,
        location:        String(emp['الموقع'] || emp['موقع العمل'] || 'غير محدد').trim(),
        salary,
        contract_status: String(emp['حالة العقد'] || '').trim(),
        contract_type:   String(emp['نوع العقد'] || 'محدد المدة').trim(),
        gender:          String(emp['الجنس'] || '').trim(),
        marital:         String(emp['حالة اجتماعية'] || emp['الحالة الاجتماعية'] || '').trim(),
        work_type:       String(emp['نوع الدوام'] || 'دوام كامل').trim(),
        hire_date:       hireRaw,
        hire_year:       serialToYear(hireRaw),
        edu:             String(emp['المؤهل الدراسي'] || 'بكالوريوس').trim(),
        insurance:       String(emp['شركة التأمين'] || 'لا يوجد').trim(),
        insurance_cat:   String(emp['فئة التأمين'] || 'بدون').trim(),
        bank:            String(emp['اسم البنك'] || emp['البنك'] || 'بنك مصر').trim()
      };
    });

  const E = DATA.employees;

  // ── KPIs ──────────────────────────────────────────────────
  DATA.total      = E.length;
  DATA.active     = E.filter(e => e.status.includes('نشط')).length;
  DATA.probation  = E.filter(e => e.status.includes('تجربه') || e.status.includes('تجربة')).length;
  DATA.terminated = E.filter(e => e.status.includes('منتهي') || e.status.includes('مستقيل')).length;

  DATA.male      = E.filter(e => e.gender === 'ذكر').length;
  DATA.female    = E.filter(e => e.gender === 'انثى' || e.gender === 'أنثى').length;
  DATA.full_time = E.filter(e => e.work_type.includes('كامل')).length;
  DATA.part_time = E.filter(e => e.work_type.includes('جزئي')).length;

  // ── Marital ────────────────────────────────────────────────
  DATA.marital = { 'متزوج': 0, 'أعزب': 0 };
  E.forEach(e => {
    if (e.marital.includes('متزوج'))                          DATA.marital['متزوج']++;
    else if (e.marital.includes('أعزب') || e.marital.includes('عزباء')) DATA.marital['أعزب']++;
  });

  // ── Hires by year ──────────────────────────────────────────
  DATA.hires_by_year = {};
  E.forEach(e => {
    const y = e.hire_year;
    if (y) DATA.hires_by_year[y] = (DATA.hires_by_year[y] || 0) + 1;
  });

  // ── Department counts ──────────────────────────────────────
  DATA.dept_counts = countBy(E, 'dept');

  // ── Contract types ─────────────────────────────────────────
  DATA.contract_types = countBy(E, 'contract_type');

  // ── Average salary by job title ───────────────────────────
  DATA.job_salary = avgBy(E, 'title', 'salary');

  // ── Salary ranges ─────────────────────────────────────────
  DATA.salary_ranges = { 'أقل من 10k': 0, '10k–15k': 0, '15k–20k': 0, 'أكثر من 20k': 0 };
  E.forEach(e => {
    if      (e.salary < 10000)  DATA.salary_ranges['أقل من 10k']++;
    else if (e.salary <= 15000) DATA.salary_ranges['10k–15k']++;
    else if (e.salary <= 20000) DATA.salary_ranges['15k–20k']++;
    else                        DATA.salary_ranges['أكثر من 20k']++;
  });

  // ── Avg salary by dept ─────────────────────────────────────
  DATA.dept_salary = avgBy(E, 'dept', 'salary');

  // ── Avg salary by gender ───────────────────────────────────
  const normGender = e => (e.gender === 'انثى' || e.gender === 'أنثى') ? 'إناث' : 'ذكور';
  DATA.salary_by_gender = avgByFn(E, normGender, 'salary');

  // ── Education counts ───────────────────────────────────────
  DATA.edu_counts = countBy(E, 'edu');

  // ── Contracts ─────────────────────────────────────────────
  DATA.active_contracts  = E.filter(e => e.contract_status === 'ساري').length;
  DATA.expired_contracts = E.filter(e => e.contract_status === 'منتهي').length;
  DATA.near_expiry       = E.filter(e =>
    e.contract_status.includes('قارب') || e.contract_status.includes('تجديد')
  ).length;

  // ── Insurance ─────────────────────────────────────────────
  DATA.insurance     = countBy(E, 'insurance');
  DATA.insurance_cat = countBy(E, 'insurance_cat');

  // ── Location & bank ───────────────────────────────────────
  DATA.loc_counts = countBy(E, 'location');
  DATA.bank       = countBy(E, 'bank');

  // ── Avg salary by education ───────────────────────────────
  DATA.edu_salary = avgBy(E, 'edu', 'salary');

  // ── Table arrays ──────────────────────────────────────────
  allEmployees      = [...DATA.employees];
  filteredEmployees = [...allEmployees];

  updateKPIs();
}

// ── Helper: count occurrences of a field value ─────────────
function countBy(arr, field) {
  return arr.reduce((acc, item) => {
    const k = item[field] || 'غير محدد';
    acc[k] = (acc[k] || 0) + 1;
    return acc;
  }, {});
}

// ── Helper: average a numeric field grouped by a string field
function avgBy(arr, groupField, numField) {
  const sums = {}, counts = {};
  arr.forEach(item => {
    const k = item[groupField];
    if (!k) return;
    sums[k]   = (sums[k]   || 0) + (item[numField] || 0);
    counts[k] = (counts[k] || 0) + 1;
  });
  const result = {};
  for (const k in sums) result[k] = counts[k] ? Math.round(sums[k] / counts[k]) : 0;
  return result;
}

// ── Helper: average grouped by a key-function ──────────────
function avgByFn(arr, keyFn, numField) {
  const sums = {}, counts = {};
  arr.forEach(item => {
    const k = keyFn(item);
    sums[k]   = (sums[k]   || 0) + (item[numField] || 0);
    counts[k] = (counts[k] || 0) + 1;
  });
  const result = {};
  for (const k in sums) result[k] = counts[k] ? Math.round(sums[k] / counts[k]) : 0;
  return result;
}

// ── Update KPI cards ─────────────────────────────────────────
function updateKPIs() {
  const set = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  };

  const E = DATA.employees || [];

  // Overview
  set('kpi-total',      DATA.total      || 0);
  set('kpi-active',     DATA.active     || 0);
  set('kpi-probation',  DATA.probation  || 0);
  set('kpi-terminated', DATA.terminated || 0);

  // Average salary (overview card)
  const salaries = E.map(e => e.salary).filter(s => s > 0);
  const avgSal   = salaries.length ? Math.round(salaries.reduce((a,b) => a+b, 0) / salaries.length) : 0;
  set('kpi-avg-salary', avgSal.toLocaleString('ar-EG'));

  // Active contracts (overview card)
  set('kpi-contracts',    DATA.active_contracts || 0);
  const nearSubEl = document.getElementById('kpi-near-expiry-sub');
  if (nearSubEl) nearSubEl.textContent = `${DATA.near_expiry || 0} قاربت على الانتهاء`;

  // Salary tab KPIs
  set('sal-avg',   avgSal.toLocaleString('ar-EG'));
  set('sal-max',   (Math.max(...salaries, 0)).toLocaleString('ar-EG'));
  set('sal-min',   (Math.min(...salaries.filter(s=>s>0), 0)).toLocaleString('ar-EG'));
  set('sal-total', salaries.reduce((a,b) => a+b, 0).toLocaleString('ar-EG'));

  // Contracts tab KPIs
  set('kpi-active-contracts',  DATA.active_contracts  || 0);
  set('kpi-expired-contracts', DATA.expired_contracts || 0);
  set('kpi-near-expiry',       DATA.near_expiry       || 0);

  // Average contract duration (rough: from hire_year span)
  const durations = E.map(e => {
    const raw = e.hire_date;
    if (!raw || typeof raw !== 'number') return null;
    const msPerDay   = 86400000;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const start      = new Date(excelEpoch.getTime() + raw * msPerDay);
    const yearsAgo   = (Date.now() - start.getTime()) / (msPerDay * 365.25);
    return yearsAgo > 0 ? yearsAgo : null;
  }).filter(Boolean);
  const avgDur = durations.length
    ? (durations.reduce((a,b) => a+b, 0) / durations.length).toFixed(1)
    : '—';
  set('kpi-avg-contract', avgDur);
}

// ══════════════════════════════════════════════════════════════
//  THEME
// ══════════════════════════════════════════════════════════════
function toggleTheme() {
  isDark = !isDark;
  document.documentElement.setAttribute('data-theme', isDark ? 'dark' : '');
  // Refresh charts so their colours match the new theme
  Object.values(chartInstances).forEach(c => c?.update());
}

// ══════════════════════════════════════════════════════════════
//  SIDEBAR
// ══════════════════════════════════════════════════════════════
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('overlay').classList.toggle('show');
}
function closeSidebar() {
  document.getElementById('sidebar')?.classList.remove('open');
  document.getElementById('overlay')?.classList.remove('show');
}

// ══════════════════════════════════════════════════════════════
//  TABS
// ══════════════════════════════════════════════════════════════
const TABS = ['overview', 'employees', 'salaries', 'contracts', 'insights'];
const TITLES = {
  overview:  'لوحة نظرة عامة',
  employees: 'الموظفون',
  salaries:  'المرتبات والرواتب',
  contracts: 'العقود والتأمينات',
  insights:  'الرؤى والتوصيات'
};

function showTab(tab) {
  TABS.forEach(t => {
    const el = document.getElementById('tab-' + t);
    if (el) el.style.display = t === tab ? '' : 'none';
  });

  document.querySelectorAll('.nav-item').forEach((el, i) => {
    el.classList.toggle('active', TABS[i] === tab);
  });

  const ptEl = document.getElementById('page-title');
  const pbEl = document.getElementById('page-breadcrumb');
  if (ptEl) ptEl.textContent = TITLES[tab] || tab;
  if (pbEl) pbEl.textContent = TITLES[tab] || tab;

  closeSidebar();

  // Only rebuild charts if data is ready; small delay lets the DOM paint first
  if (!dataLoaded) return;
  const builders = {
    overview:  buildOverviewCharts,
    salaries:  buildSalaryCharts,
    contracts: buildContractCharts,
    insights:  buildInsightCharts
  };
  if (builders[tab]) setTimeout(builders[tab], 60);
  if (tab === 'employees') buildEmployeeTable();
}

// ══════════════════════════════════════════════════════════════
//  CHART HELPERS
// ══════════════════════════════════════════════════════════════
function getChartDefaults() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        rtl: true,
        textDirection: 'rtl',
        bodyFont:  { family: 'Tajawal', size: 12 },
        titleFont: { family: 'Tajawal', size: 13 }
      }
    }
  };
}

function axisDefaults() {
  return {
    y: {
      beginAtZero: true,
      grid: { color: 'rgba(128,128,128,0.08)' },
      ticks: { font: { family: 'Tajawal', size: 11 } }
    },
    x: {
      grid: { display: false },
      ticks: { font: { family: 'Tajawal', size: 11 } }
    }
  };
}

function createChart(id, config) {
  if (chartInstances[id]) {
    chartInstances[id].destroy();
    delete chartInstances[id];
  }
  const canvas = document.getElementById(id);
  if (!canvas) return;
  chartInstances[id] = new Chart(canvas, config);
}

function makeLegend(id, labels, colors) {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerHTML = labels
    .map((l, i) =>
      `<span class="legend-item">
         <span class="legend-dot" style="background:${colors[i % colors.length]}"></span>
         ${l}
       </span>`)
    .join('');
}

// ══════════════════════════════════════════════════════════════
//  OVERVIEW CHARTS
// ══════════════════════════════════════════════════════════════
function buildOverviewCharts() {
  if (!DATA.hires_by_year) return;

  // Sort years for chronological display
  const hirYears = Object.keys(DATA.hires_by_year).sort();
  const hirVals  = hirYears.map(y => DATA.hires_by_year[y]);

  createChart('hiresChart', {
    type: 'bar',
    data: {
      labels: hirYears,
      datasets: [{
        label: 'موظفون', data: hirVals,
        backgroundColor: COLORS[0], borderRadius: 7, borderSkipped: false
      }]
    },
    options: { ...getChartDefaults(), scales: axisDefaults() }
  });

  // Department doughnut
  const deptLabels = Object.keys(DATA.dept_counts || {});
  const deptVals   = Object.values(DATA.dept_counts || {});
  createChart('deptChart', {
    type: 'doughnut',
    data: { labels: deptLabels, datasets: [{ data: deptVals, backgroundColor: COLORS, borderWidth: 2, borderColor: 'var(--surface)' }] },
    options: { ...getChartDefaults(), cutout: '62%' }
  });
  makeLegend('dept-legend', deptLabels.map((l, i) => `${l} (${deptVals[i]})`), COLORS);

  // Gender
  createChart('genderChart', {
    type: 'doughnut',
    data: {
      labels: ['ذكور', 'إناث'],
      datasets: [{ data: [DATA.male || 0, DATA.female || 0], backgroundColor: [COLORS[0], COLORS[5]], borderWidth: 2, borderColor: 'var(--surface)' }]
    },
    options: { ...getChartDefaults(), cutout: '62%' }
  });
  makeLegend('gender-legend', [`ذكور (${DATA.male || 0})`, `إناث (${DATA.female || 0})`], [COLORS[0], COLORS[5]]);

  // Work type
  createChart('workTypeChart', {
    type: 'doughnut',
    data: {
      labels: ['دوام كامل', 'دوام جزئي'],
      datasets: [{ data: [DATA.full_time || 0, DATA.part_time || 0], backgroundColor: [COLORS[2], COLORS[3]], borderWidth: 2, borderColor: 'var(--surface)' }]
    },
    options: { ...getChartDefaults(), cutout: '62%' }
  });
  makeLegend('worktype-legend', [`دوام كامل (${DATA.full_time || 0})`, `دوام جزئي (${DATA.part_time || 0})`], [COLORS[2], COLORS[3]]);

  // Marital
  const mat = DATA.marital || { 'متزوج': 0, 'أعزب': 0 };
  createChart('maritalChart', {
    type: 'doughnut',
    data: {
      labels: ['متزوج', 'أعزب'],
      datasets: [{ data: [mat['متزوج'] || 0, mat['أعزب'] || 0], backgroundColor: [COLORS[1], COLORS[4]], borderWidth: 2, borderColor: 'var(--surface)' }]
    },
    options: { ...getChartDefaults(), cutout: '62%' }
  });
  makeLegend('marital-legend', [`متزوج (${mat['متزوج'] || 0})`, `أعزب (${mat['أعزب'] || 0})`], [COLORS[1], COLORS[4]]);

  // Avg salary by job (horizontal)
  const jobLabels = Object.keys(DATA.job_salary || {});
  const jobVals   = Object.values(DATA.job_salary || {});
  createChart('salaryJobChart', {
    type: 'bar',
    data: {
      labels: jobLabels,
      datasets: [{
        label: 'متوسط الراتب', data: jobVals,
        backgroundColor: COLORS.slice(0, jobLabels.length),
        borderRadius: 5, borderSkipped: false
      }]
    },
    options: {
      ...getChartDefaults(),
      indexAxis: 'y',
      scales: {
        x: { beginAtZero: true, grid: { color: 'rgba(128,128,128,0.08)' }, ticks: { font: { family: 'Tajawal', size: 11 } } },
        y: { grid: { display: false }, ticks: { font: { family: 'Tajawal', size: 11 } } }
      }
    }
  });

  // Contract types
  const ctLabels = Object.keys(DATA.contract_types || {});
  const ctVals   = Object.values(DATA.contract_types || {});
  createChart('contractTypeChart', {
    type: 'bar',
    data: {
      labels: ctLabels,
      datasets: [{
        label: 'عدد العقود', data: ctVals,
        backgroundColor: COLORS, borderRadius: 6, borderSkipped: false
      }]
    },
    options: { ...getChartDefaults(), scales: axisDefaults() }
  });
}

// ══════════════════════════════════════════════════════════════
//  SALARY CHARTS
// ══════════════════════════════════════════════════════════════
function buildSalaryCharts() {
  if (!DATA.salary_ranges) return;

  const srLabels = Object.keys(DATA.salary_ranges);
  const srVals   = Object.values(DATA.salary_ranges);
  createChart('salaryRangeChart', {
    type: 'bar',
    data: { labels: srLabels, datasets: [{ label: 'عدد الموظفين', data: srVals, backgroundColor: COLORS, borderRadius: 6, borderSkipped: false }] },
    options: { ...getChartDefaults(), scales: axisDefaults() }
  });

  const dsLabels = Object.keys(DATA.dept_salary || {});
  const dsVals   = Object.values(DATA.dept_salary || {});
  createChart('deptSalaryChart', {
    type: 'bar',
    data: { labels: dsLabels, datasets: [{ label: 'متوسط الراتب', data: dsVals, backgroundColor: COLORS, borderRadius: 6, borderSkipped: false }] },
    options: { ...getChartDefaults(), scales: axisDefaults() }
  });

  const gLabels = Object.keys(DATA.salary_by_gender || {});
  const gVals   = Object.values(DATA.salary_by_gender || {});
  createChart('genderSalaryChart', {
    type: 'bar',
    data: { labels: gLabels, datasets: [{ label: 'متوسط الراتب', data: gVals, backgroundColor: [COLORS[5], COLORS[0]], borderRadius: 6, borderSkipped: false }] },
    options: { ...getChartDefaults(), scales: axisDefaults() }
  });

  const eduLabels = Object.keys(DATA.edu_salary || {});
  const eduVals   = Object.values(DATA.edu_salary || {});
  createChart('eduSalaryChart', {
    type: 'bar',
    data: { labels: eduLabels, datasets: [{ label: 'متوسط الراتب', data: eduVals, backgroundColor: [COLORS[0], COLORS[2], COLORS[3]], borderRadius: 6, borderSkipped: false }] },
    options: { ...getChartDefaults(), scales: axisDefaults() }
  });
}

// ══════════════════════════════════════════════════════════════
//  CONTRACT CHARTS
// ══════════════════════════════════════════════════════════════
function buildContractCharts() {
  if (!DATA.insurance) return;

  createChart('contractStatusChart', {
    type: 'doughnut',
    data: {
      labels: ['ساري', 'منتهي', 'قارب على الانتهاء'],
      datasets: [{
        data: [DATA.active_contracts || 0, DATA.expired_contracts || 0, DATA.near_expiry || 0],
        backgroundColor: ['#06d6a0', '#ef476f', '#f4a261'],
        borderWidth: 2, borderColor: 'var(--surface)'
      }]
    },
    options: { ...getChartDefaults(), cutout: '62%' }
  });
  makeLegend('contract-status-legend',
    [`ساري (${DATA.active_contracts || 0})`, `منتهي (${DATA.expired_contracts || 0})`, `قارب على الانتهاء (${DATA.near_expiry || 0})`],
    ['#06d6a0', '#ef476f', '#f4a261']
  );

  const ct2Labels = Object.keys(DATA.contract_types || {});
  const ct2Vals   = Object.values(DATA.contract_types || {});
  createChart('contractTypeChart2', {
    type: 'bar',
    data: { labels: ct2Labels, datasets: [{ label: 'عدد', data: ct2Vals, backgroundColor: COLORS, borderRadius: 6, borderSkipped: false }] },
    options: { ...getChartDefaults(), scales: axisDefaults() }
  });

  const insLabels = Object.keys(DATA.insurance || {});
  const insVals   = Object.values(DATA.insurance || {});
  createChart('insuranceChart', {
    type: 'doughnut',
    data: { labels: insLabels, datasets: [{ data: insVals, backgroundColor: COLORS, borderWidth: 2, borderColor: 'var(--surface)' }] },
    options: { ...getChartDefaults(), cutout: '62%' }
  });
  makeLegend('insurance-legend', insLabels.map((l, i) => `${l} (${insVals[i]})`), COLORS);

  const icLabels = Object.keys(DATA.insurance_cat || {});
  const icVals   = Object.values(DATA.insurance_cat || {});
  createChart('insuranceCatChart', {
    type: 'doughnut',
    data: { labels: icLabels, datasets: [{ data: icVals, backgroundColor: COLORS, borderWidth: 2, borderColor: 'var(--surface)' }] },
    options: { ...getChartDefaults(), cutout: '62%' }
  });
  makeLegend('insurancecat-legend', icLabels.map((l, i) => `فئة ${l} (${icVals[i]})`), COLORS);
}

// ══════════════════════════════════════════════════════════════
//  INSIGHT CHARTS
// ══════════════════════════════════════════════════════════════
function buildInsightCharts() {
  if (!DATA.loc_counts) return;

  const locLabels = Object.keys(DATA.loc_counts);
  const locVals   = Object.values(DATA.loc_counts);
  createChart('locationChart', {
    type: 'bar',
    data: { labels: locLabels, datasets: [{ label: 'عدد الموظفين', data: locVals, backgroundColor: COLORS, borderRadius: 6, borderSkipped: false }] },
    options: { ...getChartDefaults(), scales: axisDefaults() }
  });

  const bankLabels = Object.keys(DATA.bank || {});
  const bankVals   = Object.values(DATA.bank || {});
  createChart('bankChart', {
    type: 'doughnut',
    data: { labels: bankLabels, datasets: [{ data: bankVals, backgroundColor: COLORS, borderWidth: 2, borderColor: 'var(--surface)' }] },
    options: { ...getChartDefaults(), cutout: '62%' }
  });
  makeLegend('bank-legend', bankLabels.map((l, i) => `${l} (${bankVals[i]})`), COLORS);
}

// ══════════════════════════════════════════════════════════════
//  EMPLOYEE TABLE
// ══════════════════════════════════════════════════════════════
const avatarColors = ['#4361ee', '#7209b7', '#06d6a0', '#f4a261', '#0096c7', '#ef476f'];

function statusBadge(s) {
  if (s.includes('نشط'))                              return `<span class="badge badge-active">${s}</span>`;
  if (s.includes('تجربه') || s.includes('تجربة'))    return `<span class="badge badge-probation">${s}</span>`;
  return                                                     `<span class="badge badge-terminated">${s}</span>`;
}
function contractBadge(s) {
  if (s === 'ساري')    return `<span class="badge badge-saree">${s}</span>`;
  if (s === 'منتهي')   return `<span class="badge badge-expired">${s}</span>`;
  return                      `<span class="badge badge-near">${s}</span>`;
}

function buildEmployeeTable() {
  const start = (currentPage - 1) * PER_PAGE;
  const rows  = filteredEmployees.slice(start, start + PER_PAGE);
  const tbody = document.getElementById('emp-tbody');
  const empty = document.getElementById('empty-state');
  const rcEl  = document.getElementById('row-count');

  if (!tbody) return;

  if (rows.length === 0) {
    tbody.innerHTML = '';
    if (empty) empty.style.display = '';
    const pagEl = document.getElementById('pagination');
    if (pagEl) pagEl.innerHTML = '';
    return;
  }

  if (empty) empty.style.display = 'none';

  tbody.innerHTML = rows.map((e, i) => {
    const initials = e.name ? e.name.charAt(0) : 'م';
    const color    = avatarColors[(start + i) % avatarColors.length];
    return `<tr>
      <td class="td-id">${e.id}</td>
      <td>
        <div class="emp-name">
          <div class="emp-avatar" style="background:${color}18;color:${color}">${initials}</div>
          <span>${e.name}</span>
        </div>
      </td>
      <td>${e.dept}</td>
      <td class="td-secondary">${e.title}</td>
      <td>${statusBadge(e.status)}</td>
      <td class="td-secondary">${e.location}</td>
      <td class="td-salary">${(e.salary || 0).toLocaleString('ar-EG')} ج.م</td>
      <td>${contractBadge(e.contract_status)}</td>
    </tr>`;
  }).join('');

  if (rcEl) rcEl.textContent = `(${filteredEmployees.length} موظف)`;
  buildPagination();
}

function buildPagination() {
  const total = Math.ceil(filteredEmployees.length / PER_PAGE);
  const pagEl = document.getElementById('pagination');
  if (!pagEl) return;
  if (total <= 1) { pagEl.innerHTML = ''; return; }

  let html = '';
  if (currentPage > 1) html += `<button class="page-btn" onclick="goPage(${currentPage - 1})">‹</button>`;

  for (let i = 1; i <= total; i++) {
    if (i === 1 || i === total || Math.abs(i - currentPage) <= 1) {
      html += `<button class="page-btn ${i === currentPage ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
    } else if (Math.abs(i - currentPage) === 2) {
      html += `<span class="page-ellipsis">…</span>`;
    }
  }

  if (currentPage < total) html += `<button class="page-btn" onclick="goPage(${currentPage + 1})">›</button>`;
  html += `<span class="page-info">${(currentPage - 1) * PER_PAGE + 1}–${Math.min(currentPage * PER_PAGE, filteredEmployees.length)} من ${filteredEmployees.length}</span>`;
  pagEl.innerHTML = html;
}

function goPage(p) { currentPage = p; buildEmployeeTable(); }

// ── Filters ───────────────────────────────────────────────────
function filterTable() {
  const dept   = document.getElementById('filter-dept')?.value   || '';
  const status = document.getElementById('filter-status')?.value || '';
  const gender = document.getElementById('filter-gender')?.value || '';
  const loc    = document.getElementById('filter-loc')?.value    || '';
  const search = (document.getElementById('globalSearch')?.value || '').toLowerCase();

  filteredEmployees = allEmployees.filter(e =>
    (!dept   || e.dept === dept) &&
    (!status || e.status.includes(status)) &&
    (!gender || e.gender === gender) &&
    (!loc    || e.location === loc) &&
    (!search || e.name.toLowerCase().includes(search)  ||
                e.id.toLowerCase().includes(search)    ||
                e.title.toLowerCase().includes(search))
  );
  currentPage = 1;
  buildEmployeeTable();
}

function resetFilters() {
  ['filter-dept', 'filter-status', 'filter-gender', 'filter-loc'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const gs = document.getElementById('globalSearch');
  if (gs) gs.value = '';
  filteredEmployees = [...allEmployees];
  currentPage = 1;
  buildEmployeeTable();
}

function handleSearch(v) {
  const empTab = document.getElementById('tab-employees');
  if (empTab && empTab.style.display !== 'none') filterTable();
}

// ── Sort ──────────────────────────────────────────────────────
function sortTable(key) {
  if (sortKey === key) sortDir *= -1;
  else { sortKey = key; sortDir = 1; }
  filteredEmployees.sort((a, b) => {
    const av = a[key] ?? '';
    const bv = b[key] ?? '';
    if (typeof av === 'number') return (av - bv) * sortDir;
    return String(av).localeCompare(String(bv), 'ar') * sortDir;
  });
  buildEmployeeTable();
}

// ── Export CSV ────────────────────────────────────────────────
function exportCSV() {
  const headers = ['الكود', 'الاسم', 'الإدارة', 'المسمى', 'الحالة', 'الموقع', 'الراتب', 'حالة العقد', 'الجنس'];
  const rows    = filteredEmployees.map(e =>
    [e.id, e.name, e.dept, e.title, e.status, e.location, e.salary, e.contract_status, e.gender]
      .map(v => `"${String(v || '').replace(/"/g, '""')}"`)
  );
  const csv  = [headers, ...rows].map(r => r.join(',')).join('\r\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: 'employees.csv' });
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ══════════════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════════════
window.addEventListener('DOMContentLoaded', () => {
  loadExcel();
});
