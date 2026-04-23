/* =============================================
   DataViz – app.js
   XLS/XLSX Visualizer
   ============================================= */

'use strict';

// ─── Constants ────────────────────────────────
const MAX_TABLE_PREVIEW_ROWS = 200;

// ─── State ────────────────────────────────────
const state = {
  workbook: null,
  sheetData: [],      // Array of {header, rows} per sheet
  currentSheet: 0,
  vizStyle: 'dashboard',   // 'dashboard' | 'focus' | 'compare'
  chartType: 'bar',
  xAxisCol: '',
  yAxisCols: [],
  charts: [],          // Chart.js instances
  isDark: true,
};

// ─── Colour Palettes ──────────────────────────
const PALETTES = {
  vivid: [
    'rgba(108,99,255,0.85)', 'rgba(0,210,255,0.85)', 'rgba(35,209,139,0.85)',
    'rgba(245,166,35,0.85)', 'rgba(255,95,95,0.85)',  'rgba(255,159,67,0.85)',
    'rgba(164,69,240,0.85)', 'rgba(0,184,212,0.85)',  'rgba(255,107,107,0.85)',
    'rgba(78,205,196,0.85)',
  ],
  soft: [
    'rgba(130,119,255,0.7)', 'rgba(64,220,255,0.7)', 'rgba(72,219,160,0.7)',
    'rgba(255,204,102,0.7)', 'rgba(255,130,130,0.7)','rgba(255,189,110,0.7)',
    'rgba(192,120,255,0.7)', 'rgba(66,211,227,0.7)', 'rgba(255,145,145,0.7)',
    'rgba(108,230,218,0.7)',
  ],
};

function getPalette(n) {
  const base = PALETTES.vivid;
  const result = [];
  for (let i = 0; i < n; i++) result.push(base[i % base.length]);
  return result;
}

// ─── DOM Refs ─────────────────────────────────
const dom = {
  fileInput:      () => document.getElementById('fileInput'),
  dropZone:       () => document.getElementById('dropZone'),
  browseBtn:      () => document.getElementById('browseBtn'),
  loadSample:     () => document.getElementById('loadSample'),
  uploadSection:  () => document.getElementById('uploadSection'),
  vizSection:     () => document.getElementById('vizSection'),
  fileName:       () => document.getElementById('fileName'),
  rowCount:       () => document.getElementById('rowCount'),
  colCount:       () => document.getElementById('colCount'),
  resetBtn:       () => document.getElementById('resetBtn'),
  sheetSelector:  () => document.getElementById('sheetSelector'),
  sheetSelect:    () => document.getElementById('sheetSelect'),
  tabs:           () => document.querySelectorAll('.tab'),
  chartChips:     () => document.querySelectorAll('.chip'),
  xAxisCol:       () => document.getElementById('xAxisCol'),
  yAxisCols:      () => document.getElementById('yAxisCols'),
  chartArea:      () => document.getElementById('chartArea'),
  tableToggle:    () => document.getElementById('tableToggle'),
  tableWrapper:   () => document.getElementById('tableWrapper'),
  dataTable:      () => document.getElementById('dataTable'),
  toastContainer: () => document.getElementById('toastContainer'),
  themeToggle:    () => document.getElementById('themeToggle'),
  sunIcon:        () => document.getElementById('sunIcon'),
  moonIcon:       () => document.getElementById('moonIcon'),
};

// ─── Init ─────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  restoreTheme();
});

function bindEvents() {
  // File upload
  dom.browseBtn().addEventListener('click', e => { e.stopPropagation(); dom.fileInput().click(); });
  dom.dropZone().addEventListener('click', () => dom.fileInput().click());
  dom.fileInput().addEventListener('change', e => handleFile(e.target.files[0]));

  // Drag & drop
  dom.dropZone().addEventListener('dragover', e => { e.preventDefault(); dom.dropZone().classList.add('drag-over'); });
  dom.dropZone().addEventListener('dragleave', () => dom.dropZone().classList.remove('drag-over'));
  dom.dropZone().addEventListener('drop', e => {
    e.preventDefault();
    dom.dropZone().classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  });

  // Keyboard on drop zone
  dom.dropZone().addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); dom.fileInput().click(); }
  });

  // Load sample
  dom.loadSample().addEventListener('click', loadSampleData);

  // Reset
  dom.resetBtn().addEventListener('click', resetApp);

  // Sheet select
  dom.sheetSelect().addEventListener('change', () => {
    state.currentSheet = parseInt(dom.sheetSelect().value, 10);
    onSheetChange();
  });

  // Style tabs
  dom.tabs().forEach(tab => {
    tab.addEventListener('click', () => {
      dom.tabs().forEach(t => { t.classList.remove('active'); t.setAttribute('aria-selected', 'false'); });
      tab.classList.add('active');
      tab.setAttribute('aria-selected', 'true');
      state.vizStyle = tab.dataset.style;
      renderCharts();
    });
  });

  // Chart type chips
  dom.chartChips().forEach(chip => {
    chip.addEventListener('click', () => {
      dom.chartChips().forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      state.chartType = chip.dataset.chart;
      renderCharts();
    });
  });

  // Table toggle
  dom.tableToggle().addEventListener('click', () => {
    const w = dom.tableWrapper();
    const btn = dom.tableToggle();
    const isHidden = w.classList.toggle('hidden');
    btn.innerHTML = isHidden
      ? `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 3H5a2 2 0 0 0-2 2v4m6-6h10a2 2 0 0 1 2 2v4M9 3v18m0 0h10a2 2 0 0 0 2-2V9M9 21H5a2 2 0 0 0-2-2V9m0 0h18"/></svg> Show Data Table`
      : `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 3H5a2 2 0 0 0-2 2v4m6-6h10a2 2 0 0 1 2 2v4M9 3v18m0 0h10a2 2 0 0 0 2-2V9M9 21H5a2 2 0 0 0-2-2V9m0 0h18"/></svg> Hide Data Table`;
  });

  // Theme toggle
  dom.themeToggle().addEventListener('click', toggleTheme);

  // X/Y axis changes
  dom.xAxisCol().addEventListener('change', () => {
    state.xAxisCol = dom.xAxisCol().value;
    renderCharts();
  });
}

// ─── Theme ────────────────────────────────────
function restoreTheme() {
  const saved = localStorage.getItem('dataviz-theme');
  if (saved === 'light') applyTheme('light');
  else applyTheme('dark');
}

function toggleTheme() {
  applyTheme(state.isDark ? 'light' : 'dark');
}

function applyTheme(mode) {
  state.isDark = mode === 'dark';
  document.body.classList.toggle('light', !state.isDark);
  dom.sunIcon().classList.toggle('hidden', !state.isDark);
  dom.moonIcon().classList.toggle('hidden', state.isDark);
  localStorage.setItem('dataviz-theme', mode);
  if (state.charts.length) renderCharts(); // re-render for theme-aware colours
}

// ─── File Handling ────────────────────────────
function handleFile(file) {
  if (!file) return;
  if (!/\.(xls|xlsx)$/i.test(file.name)) {
    showToast('Please upload a .xls or .xlsx file.', 'error');
    return;
  }
  dom.fileName().textContent = file.name;
  readWorkbook(file);
}

function readWorkbook(file) {
  dom.chartArea().innerHTML = '<div class="spinner-overlay"><div class="spinner"></div><span>Parsing file…</span></div>';
  showVizSection();

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      state.workbook = wb;
      parseWorkbook(wb);
    } catch (err) {
      showToast('Error parsing file: ' + err.message, 'error');
      resetApp();
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseWorkbook(wb) {
  state.sheetData = wb.SheetNames.map(name => {
    const ws = wb.Sheets[name];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!raw.length) return { name, header: [], rows: [] };
    const header = raw[0].map(h => String(h));
    const rows   = raw.slice(1).map(r => {
      const obj = {};
      header.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i] : ''; });
      return obj;
    });
    return { name, header, rows };
  });

  // Populate sheet selector
  const sheetSel = dom.sheetSelect();
  sheetSel.innerHTML = '';
  state.sheetData.forEach((s, i) => {
    const opt = document.createElement('option');
    opt.value = i; opt.textContent = s.name;
    sheetSel.appendChild(opt);
  });
  dom.sheetSelector().classList.toggle('hidden', state.sheetData.length <= 1);

  state.currentSheet = 0;
  onSheetChange();
}

function onSheetChange() {
  const sheet = state.sheetData[state.currentSheet];
  if (!sheet || !sheet.rows.length) {
    showToast('This sheet has no data.', 'warning');
    return;
  }

  dom.rowCount().textContent = sheet.rows.length + ' rows';
  dom.colCount().textContent = sheet.header.length + ' columns';

  buildColumnPickers(sheet);
  buildDataTable(sheet);
  renderCharts();
}

// ─── Column Pickers ───────────────────────────
function buildColumnPickers(sheet) {
  const xSel = dom.xAxisCol();
  xSel.innerHTML = '';
  sheet.header.forEach(h => {
    const o = document.createElement('option');
    o.value = h; o.textContent = h;
    xSel.appendChild(o);
  });
  state.xAxisCol = sheet.header[0] || '';
  xSel.value = state.xAxisCol;

  // Y axis checkboxes – default: all numeric columns
  // Pre-compute numeric columns with a single pass through rows to avoid O(n*m) repeated checks
  const numericColSet = new Set();
  for (const row of sheet.rows) {
    for (const h of sheet.header) {
      if (!numericColSet.has(h) && (typeof row[h] === 'number' || (row[h] !== '' && !isNaN(Number(row[h]))))) {
        numericColSet.add(h);
      }
    }
  }
  const numericCols = sheet.header.filter(h => numericColSet.has(h));
  const yCols = numericCols.length ? numericCols : sheet.header.filter(h => h !== state.xAxisCol);
  state.yAxisCols = yCols.slice(0, 4); // default first 4

  const container = dom.yAxisCols();
  container.innerHTML = '';
  yCols.forEach(h => {
    const item = document.createElement('label');
    item.className = 'checkbox-item' + (state.yAxisCols.includes(h) ? ' selected' : '');
    item.innerHTML = `<input type="checkbox" ${state.yAxisCols.includes(h) ? 'checked' : ''} value="${escapeAttr(h)}" /> ${escapeHtml(h)}`;
    item.querySelector('input').addEventListener('change', evt => {
      item.classList.toggle('selected', evt.target.checked);
      if (evt.target.checked) { if (!state.yAxisCols.includes(h)) state.yAxisCols.push(h); }
      else { state.yAxisCols = state.yAxisCols.filter(c => c !== h); }
      renderCharts();
    });
    container.appendChild(item);
  });
}

// ─── Data Table ───────────────────────────────
function buildDataTable(sheet) {
  const table = dom.dataTable();
  table.innerHTML = '';

  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  sheet.header.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    hr.appendChild(th);
  });
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  const preview = sheet.rows.slice(0, MAX_TABLE_PREVIEW_ROWS);
  preview.forEach(row => {
    const tr = document.createElement('tr');
    sheet.header.forEach(h => {
      const td = document.createElement('td');
      td.textContent = row[h] !== undefined ? row[h] : '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

// ─── Chart Rendering ──────────────────────────
function renderCharts() {
  destroyCharts();
  const area = dom.chartArea();
  area.innerHTML = '';

  const sheet = state.sheetData[state.currentSheet];
  if (!sheet || !sheet.rows.length) return;

  area.className = 'chart-area style-' + state.vizStyle;

  if (state.vizStyle === 'dashboard') renderDashboard(sheet, area);
  else if (state.vizStyle === 'focus')  renderFocus(sheet, area);
  else if (state.vizStyle === 'compare') renderCompare(sheet, area);
}

// --- Dashboard style ---
function renderDashboard(sheet, area) {
  const yCols = state.yAxisCols;
  if (!yCols.length) { showEmptyHint(area); return; }

  // Summary stat cards
  yCols.forEach((col, idx) => {
    const vals = numericValues(sheet.rows, col);
    if (!vals.length) return;
    const sum  = vals.reduce((a, b) => a + b, 0);
    const avg  = sum / vals.length;
    const max  = Math.max(...vals);
    const min  = Math.min(...vals);

    const card = document.createElement('div');
    card.className = 'stat-card';
    card.style.setProperty('--gradient', `linear-gradient(90deg, ${PALETTES.vivid[idx % PALETTES.vivid.length]}, ${PALETTES.soft[idx % PALETTES.soft.length]})`);
    card.innerHTML = `
      <div class="stat-label">${escapeHtml(col)}</div>
      <div class="stat-value">${formatNum(avg)}</div>
      <div class="stat-desc">avg &nbsp;·&nbsp; Σ ${formatNum(sum)} &nbsp;·&nbsp; ↑ ${formatNum(max)} &nbsp;·&nbsp; ↓ ${formatNum(min)}</div>
    `;
    area.appendChild(card);
  });

  // One chart per numeric column
  yCols.forEach((col, idx) => {
    const card = createChartCard(col, `vs ${state.xAxisCol}`);
    area.appendChild(card);
    const canvas = card.querySelector('canvas');
    const labels = sheet.rows.map(r => String(r[state.xAxisCol] || ''));
    const values = sheet.rows.map(r => toNumber(r[col]));
    const color  = PALETTES.vivid[idx % PALETTES.vivid.length];
    const chart  = buildChart(canvas, state.chartType, labels, [{
      label: col,
      data: values,
      backgroundColor: color,
      borderColor: color.replace('0.85', '1'),
      borderWidth: 2,
      fill: state.chartType === 'line',
      tension: 0.35,
      pointRadius: 3,
    }]);
    state.charts.push(chart);
  });
}

// --- Focus style ---
function renderFocus(sheet, area) {
  const yCols = state.yAxisCols;
  if (!yCols.length) { showEmptyHint(area); return; }

  const card = createChartCard('Data Overview', `${yCols.length} series`);
  area.appendChild(card);
  const canvas = card.querySelector('canvas');
  const labels = sheet.rows.map(r => String(r[state.xAxisCol] || ''));
  const datasets = yCols.map((col, idx) => {
    const color = PALETTES.vivid[idx % PALETTES.vivid.length];
    return {
      label: col,
      data: sheet.rows.map(r => toNumber(r[col])),
      backgroundColor: PALETTES.soft[idx % PALETTES.soft.length],
      borderColor: color,
      borderWidth: 2,
      fill: state.chartType === 'line',
      tension: 0.35,
      pointRadius: 3,
    };
  });
  const chart = buildChart(canvas, state.chartType, labels, datasets);
  state.charts.push(chart);
}

// --- Compare style ---
function renderCompare(sheet, area) {
  const yCols = state.yAxisCols;
  if (!yCols.length) { showEmptyHint(area); return; }

  const labels = sheet.rows.map(r => String(r[state.xAxisCol] || ''));

  // Left panel: primary chart
  const leftCard = createChartCard('Primary View', yCols[0] || '');
  area.appendChild(leftCard);
  const leftCanvas = leftCard.querySelector('canvas');
  const leftDatasets = yCols.slice(0, Math.ceil(yCols.length / 2)).map((col, i) => {
    const c = PALETTES.vivid[i % PALETTES.vivid.length];
    return { label: col, data: sheet.rows.map(r => toNumber(r[col])), backgroundColor: PALETTES.soft[i], borderColor: c, borderWidth: 2, fill: false, tension: 0.35, pointRadius: 3 };
  });
  state.charts.push(buildChart(leftCanvas, state.chartType, labels, leftDatasets));

  // Right panel: alternative chart type or remaining columns
  const altType = state.chartType === 'bar' ? 'line' : state.chartType === 'line' ? 'bar' : 'pie';
  const rightCard = createChartCard('Alternate View', altType.charAt(0).toUpperCase() + altType.slice(1));
  area.appendChild(rightCard);
  const rightCanvas = rightCard.querySelector('canvas');
  const rightDatasets = yCols.slice(Math.ceil(yCols.length / 2)).length
    ? yCols.slice(Math.ceil(yCols.length / 2)).map((col, i) => {
        const c = PALETTES.vivid[(i + 5) % PALETTES.vivid.length];
        return { label: col, data: sheet.rows.map(r => toNumber(r[col])), backgroundColor: PALETTES.soft[i + 5], borderColor: c, borderWidth: 2, fill: false, tension: 0.35, pointRadius: 3 };
      })
    : yCols.map((col, i) => {
        const c = PALETTES.vivid[(i + 3) % PALETTES.vivid.length];
        return { label: col, data: sheet.rows.map(r => toNumber(r[col])), backgroundColor: PALETTES.soft[i + 3], borderColor: c, borderWidth: 2, fill: false, tension: 0.35, pointRadius: 3 };
      });
  state.charts.push(buildChart(rightCanvas, altType, labels, rightDatasets));
}

// ─── Chart Factory ────────────────────────────
function buildChart(canvas, type, labels, datasets) {
  const isDark = state.isDark;
  const gridColor   = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const tickColor   = isDark ? '#8892b0' : '#556080';
  const legendColor = isDark ? '#e8eaf6' : '#1a1d2e';

  // For pie/polar/radar use all colours
  if (['pie', 'doughnut', 'polarArea'].includes(type)) {
    datasets = datasets.map(ds => ({
      ...ds,
      backgroundColor: getPalette(ds.data.length),
      borderColor: isDark ? '#1a1d2e' : '#ffffff',
      borderWidth: 2,
    }));
  }

  const config = {
    type,
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 500, easing: 'easeInOutQuart' },
      plugins: {
        legend: {
          labels: {
            color: legendColor,
            font: { family: "'Inter', sans-serif", size: 12 },
            padding: 16,
            usePointStyle: true,
            pointStyleWidth: 10,
          },
        },
        tooltip: {
          backgroundColor: isDark ? '#22263a' : '#ffffff',
          borderColor: isDark ? '#2e3454' : '#d8ddf0',
          borderWidth: 1,
          titleColor: isDark ? '#e8eaf6' : '#1a1d2e',
          bodyColor: isDark ? '#8892b0' : '#556080',
          padding: 12,
          cornerRadius: 10,
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${formatNum(ctx.parsed.y ?? ctx.parsed)}`,
            // Note: radial charts (pie, polarArea) expose ctx.parsed directly as a number,
            // while cartesian charts use ctx.parsed.y — the ?? fallback handles both cases.
          },
        },
      },
      scales: ['pie', 'doughnut', 'polarArea', 'radar'].includes(type) ? {} : {
        x: {
          grid: { color: gridColor },
          ticks: { color: tickColor, font: { size: 11 }, maxTicksLimit: 20, maxRotation: 45 },
          border: { color: gridColor },
        },
        y: {
          grid: { color: gridColor },
          ticks: { color: tickColor, font: { size: 11 } },
          border: { color: gridColor },
        },
      },
    },
  };

  // Radar scales
  if (type === 'radar') {
    config.options.scales = {
      r: {
        grid: { color: gridColor },
        ticks: { color: tickColor, backdropColor: 'transparent' },
        pointLabels: { color: tickColor, font: { size: 11 } },
      },
    };
  }

  return new Chart(canvas, config);
}

// ─── Chart Card ───────────────────────────────
function createChartCard(title, subtitle) {
  const card = document.createElement('div');
  card.className = 'chart-card';
  card.innerHTML = `
    <div class="chart-card-header">
      <div>
        <div class="chart-card-title">${escapeHtml(title)}</div>
        <div class="chart-card-subtitle">${escapeHtml(subtitle)}</div>
      </div>
    </div>
    <div class="chart-card-body">
      <canvas></canvas>
    </div>
  `;
  return card;
}

function destroyCharts() {
  state.charts.forEach(c => c.destroy());
  state.charts = [];
}

function showEmptyHint(area) {
  area.innerHTML = `<div class="spinner-overlay"><span>Select at least one Y-axis column to visualize.</span></div>`;
}

// ─── Section Toggle ───────────────────────────
function showVizSection() {
  dom.uploadSection().classList.add('hidden');
  dom.vizSection().classList.remove('hidden');
}

function resetApp() {
  destroyCharts();
  state.workbook = null;
  state.sheetData = [];
  state.currentSheet = 0;
  state.yAxisCols = [];
  dom.fileInput().value = '';
  dom.chartArea().innerHTML = '';
  dom.dataTable().innerHTML = '';
  dom.tableWrapper().classList.add('hidden');
  dom.tableToggle().innerHTML = `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 3H5a2 2 0 0 0-2 2v4m6-6h10a2 2 0 0 1 2 2v4M9 3v18m0 0h10a2 2 0 0 0 2-2V9M9 21H5a2 2 0 0 0-2-2V9m0 0h18"/></svg> Show Data Table`;
  dom.vizSection().classList.add('hidden');
  dom.uploadSection().classList.remove('hidden');
}

// ─── Sample Data ──────────────────────────────
function loadSampleData() {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const sales  = [12400,14800,11200,17600,15300,18900,22100,19800,16400,21200,23800,25600];
  const costs  = [8200, 9100, 7800, 10200,9600, 11400,12800,11600,9900, 12300,14200,15100];
  const profit = sales.map((s, i) => s - costs[i]);
  const units  = [310, 370, 280, 440, 380, 472, 552, 495, 410, 530, 595, 640];

  const ws_data = [
    ['Month','Sales','Costs','Profit','Units'],
    ...months.map((m, i) => [m, sales[i], costs[i], profit[i], units[i]]),
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, 'Sales Data');

  // Second sheet – regional
  const regions  = ['North','South','East','West','Central'];
  const q1 = [45600,38200,52100,41300,29800];
  const q2 = [48900,41200,55800,44600,32100];
  const q3 = [52300,43800,58200,47200,35400];
  const q4 = [56100,46200,62400,51000,38700];
  const ws2_data = [
    ['Region','Q1','Q2','Q3','Q4'],
    ...regions.map((r, i) => [r, q1[i], q2[i], q3[i], q4[i]]),
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(ws2_data);
  XLSX.utils.book_append_sheet(wb, ws2, 'Regional');

  dom.fileName().textContent = 'sample-data.xlsx';
  state.workbook = wb;
  dom.chartArea().innerHTML = '<div class="spinner-overlay"><div class="spinner"></div><span>Loading sample…</span></div>';
  showVizSection();
  setTimeout(() => { parseWorkbook(wb); showToast('Sample data loaded!', 'success'); }, 300);
}

// ─── Utilities ────────────────────────────────
function numericValues(rows, col) {
  return rows.map(r => toNumber(r[col])).filter(v => !isNaN(v) && v !== null);
}

function toNumber(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function formatNum(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(2) + 'M';
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1) + 'k';
  return Number.isInteger(n) ? n.toString() : n.toFixed(2);
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeAttr(str) {
  return String(str).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ─── Toast Notifications ─────────────────────
function showToast(msg, type = 'info') {
  const container = dom.toastContainer();
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = msg;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'fadeOutToast 0.3s ease forwards';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}
