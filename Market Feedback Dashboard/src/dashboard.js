/**
 * Market Feedback Dashboard — Main Logic
 * All chart rendering, filter management, and interactions.
 */

// ─── Configuration ──────────────────────────────────────────────────────────
const BENCHMARK_BRAND = DASHBOARD_DATA.metadata.benchmarkBrand;
const MAX_BRANDS = 5;
const BRAND_COLORS = [
  '#6366f1', '#f59e0b', '#10b981', '#f43f5e', '#8b5cf6',
  '#06b6d4', '#ec4899', '#14b8a6', '#f97316', '#84cc16'
];
const MUTED_COLORS = [
  '#64748b', '#78716c', '#71717a', '#6b7280', '#737373',
  '#a8a29e', '#a1a1aa', '#9ca3af', '#a3a3a3', '#94a3b8'
];
const MONTH_ORDER = DASHBOARD_DATA.metadata.months;

// ─── Price Positioning Configuration ────────────────────────────────────────
const PP_BENCHMARK_BRAND = 'Tata tiscon';
const PP_JSW_BRANDS = ['JSW ONE TMT 550', 'JSW ONE TMT 550D'];

// Curated brand list: { key, label, dataBrands (array), type }
const PP_BRAND_CONFIG = [
  { key: 'tiscon', label: 'TISCON (TSL)', dataBrands: ['Tata tiscon'], type: 'benchmark' },
  { key: 'neo', label: 'Neo', dataBrands: ['JSW Neosteel'], type: 'competitor' },
  { key: 'jspl', label: 'JSPL', dataBrands: ['Jindal Steel & Power Ltd. (JSPL)', 'Local JSPL  tmt'], type: 'competitor' },
  { key: 'sail', label: 'SAIL', dataBrands: ['SAIL', 'Sail Seqr', 'Prime Gold Sail  Jvc tmt'], type: 'competitor' },
  { key: 'jsw550d', label: 'JSW One Fe 550D', dataBrands: ['JSW ONE TMT 550D'], type: 'jsw' },
  { key: 'jsw550', label: 'JSW One Fe 550', dataBrands: ['JSW ONE TMT 550'], type: 'jsw' },
  { key: 'branded', label: 'Branded Prem Sec.', dataBrands: ['Rathi', 'Rathi TMT', 'Kamdhenu', 'Kamdhenu Next', 'Goel', 'Rana tmt', 'GK', 'Jindal UHD', 'Local Jindal UHD'], type: 'aggregate' },
];

// Region-to-state mapping (from JSW Price List regions to dashboard states)
const REGION_STATE_MAP = {
  'Delhi': ['DELHI'],
  'Haryana': ['HARYANA'],
  'Punjab': ['PUNJAB'],
  'Rajasthan': ['RAJASTHAN'],
  'Chandigarh': ['CHANDIGARH'],
  'Himachal Pradesh': ['HIMACHAL PRADESH'],
  'Uttarakhand': ['UTTARAKHAND'],
  'Jammu & Kashmir': ['JAMMU AND KASHMIR'],
  'Uttar Pradesh': ['UTTAR PRADESH'],
  'Chhattisgarh': ['CHHATTISGARH'],
  'Madhya Pradesh': ['MADHYA PRADESH'],
  'Odisha': ['ODISHA'],
  'Jharkhand': ['JHARKHAND'],
  'Bihar': ['BIHAR'],
  'West Bengal': ['WEST BENGAL'],
  'Gujarat': ['GUJARAT'],
  'Maharashtra': ['MAHARASHTRA'],
};

// ─── State ──────────────────────────────────────────────────────────────────
const State = {
  filters: {
    months: [],
    state: 'ALL',
    district: 'ALL',
    brands: [],
    timePeriod: 'month', // 'month' or 'week'
  },
  charts: {
    priceTrend: null,
    priceDistribution: null,
    dealerCount: null,
  },
  table: {
    page: 1,
    perPage: 50,
    sortCol: 'm',
    sortDir: 'desc',
    search: '',
  },
  filteredData: [],
  pricePositioning: {
    region: null,
    chart: null,
    initialized: false,
  },
};

// ─── Utilities ──────────────────────────────────────────────────────────────
function mean(arr) {
  if (!arr.length) return 0;
  let sum = 0;
  for (let i = 0; i < arr.length; i++) sum += arr[i];
  return sum / arr.length;
}

function median(arr) {
  if (!arr.length) return 0;
  const s = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

function quantile(arr, q) {
  const s = [...arr].sort((a, b) => a - b);
  const pos = (s.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (s[base + 1] !== undefined) return s[base] + rest * (s[base + 1] - s[base]);
  return s[base];
}

function formatINR(val) {
  if (val == null || isNaN(val)) return '--';
  return '₹' + Math.round(val).toLocaleString('en-IN');
}

function titleCase(s) {
  if (!s) return '';
  return s.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
}

function brandColor(brand, idx) {
  if (brand === BENCHMARK_BRAND && BENCHMARK_BRAND) return '#6366f1';
  return BRAND_COLORS[idx % BRAND_COLORS.length];
}

function monthIndex(m) {
  return MONTH_ORDER.indexOf(m);
}

// ─── Filter Logic ───────────────────────────────────────────────────────────
function applyFilters() {
  const { months, state, district } = State.filters;
  let data = DASHBOARD_DATA.filterableData;

  if (months.length > 0) {
    const mset = new Set(months);
    data = data.filter(r => mset.has(r.m));
  }
  if (state !== 'ALL') {
    data = data.filter(r => r.s === state);
  }
  if (district !== 'ALL') {
    data = data.filter(r => r.d === district);
  }

  State.filteredData = data;
  State.table.page = 1;

  renderKPIs();
  renderPriceTrend();
  renderPriceDistribution();
  renderHeatmap();
  renderDealerCount();
  renderTable();
}

// ─── Filter UI Initialization ───────────────────────────────────────────────
function initFilters() {
  const meta = DASHBOARD_DATA.metadata;

  // Month multi-select
  initMultiSelect('month-filter', meta.months, State.filters.months, (selected) => {
    State.filters.months = selected;
    applyFilters();
  }, 99);

  // Time Period dropdown
  document.getElementById('time-period-filter').addEventListener('change', (e) => {
    State.filters.timePeriod = e.target.value;
    applyFilters();
  });

  // State dropdown
  const stateSelect = document.getElementById('state-filter');
  meta.states.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = titleCase(s);
    stateSelect.appendChild(opt);
  });
  stateSelect.addEventListener('change', () => {
    State.filters.state = stateSelect.value;
    State.filters.district = 'ALL';
    populateDistricts();
    applyFilters();
  });

  // District dropdown
  populateDistricts();
  document.getElementById('district-filter').addEventListener('change', (e) => {
    State.filters.district = e.target.value;
    applyFilters();
  });

  // Brand multi-select — default top 5
  const defaultBrands = meta.topBrands.slice(0, 5);
  State.filters.brands = [...defaultBrands];
  initMultiSelect('brand-filter', meta.topBrands, defaultBrands, (selected) => {
    State.filters.brands = selected;
    applyFilters();
  }, MAX_BRANDS);

  // Reset
  document.getElementById('reset-filters').addEventListener('click', resetFilters);
}

function populateDistricts() {
  const sel = document.getElementById('district-filter');
  sel.innerHTML = '<option value="ALL">All Districts</option>';
  const st = State.filters.state;
  if (st !== 'ALL' && DASHBOARD_DATA.metadata.districts[st]) {
    DASHBOARD_DATA.metadata.districts[st].forEach(d => {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = titleCase(d);
      sel.appendChild(opt);
    });
    sel.disabled = false;
  } else {
    sel.disabled = true;
  }
}

function resetFilters() {
  State.filters.months = [];
  State.filters.state = 'ALL';
  State.filters.district = 'ALL';
  State.filters.timePeriod = 'month';
  document.getElementById('time-period-filter').value = 'month';
  State.filters.brands = DASHBOARD_DATA.metadata.topBrands.slice(0, 5);

  document.getElementById('state-filter').value = 'ALL';
  populateDistricts();

  // Reset multi-selects
  resetMultiSelect('month-filter', []);
  resetMultiSelect('brand-filter', State.filters.brands);

  applyFilters();
}

// ─── Multi-Select Component ─────────────────────────────────────────────────
function initMultiSelect(containerId, options, defaultSelected, onChange, maxItems) {
  const container = document.getElementById(containerId);
  const trigger = container.querySelector('.multi-select-trigger');
  const dropdown = container.querySelector('.multi-select-dropdown');
  const searchInput = container.querySelector('.multi-select-search');
  const optionsContainer = container.querySelector('.multi-select-options');

  let selected = new Set(defaultSelected);

  function renderOptions(filter = '') {
    optionsContainer.innerHTML = '';
    const f = filter.toLowerCase();
    options.forEach(opt => {
      if (f && !opt.toLowerCase().includes(f)) return;
      const div = document.createElement('div');
      div.className = 'multi-select-option' +
        (selected.has(opt) ? ' selected' : '') +
        (!selected.has(opt) && selected.size >= maxItems ? ' disabled' : '');
      div.innerHTML = `<span>${selected.has(opt) ? '✓' : '○'}</span> ${opt}`;
      div.addEventListener('click', () => {
        if (selected.has(opt)) {
          selected.delete(opt);
        } else if (selected.size < maxItems) {
          selected.add(opt);
        }
        renderOptions(filter);
        renderTrigger();
        onChange([...selected]);
      });
      optionsContainer.appendChild(div);
    });
  }

  function renderTrigger() {
    trigger.innerHTML = '';
    if (selected.size === 0) {
      trigger.innerHTML = `<span class="placeholder">${containerId === 'month-filter' ? 'All Months' : 'Select brands...'}</span>`;
      return;
    }
    selected.forEach(val => {
      const chip = document.createElement('span');
      chip.className = 'multi-select-chip';
      const label = val.length > 14 ? val.substring(0, 12) + '..' : val;
      chip.innerHTML = `${label} <span class="remove">&times;</span>`;
      chip.querySelector('.remove').addEventListener('click', (e) => {
        e.stopPropagation();
        selected.delete(val);
        renderTrigger();
        renderOptions(searchInput.value);
        onChange([...selected]);
      });
      trigger.appendChild(chip);
    });
  }

  trigger.addEventListener('click', () => {
    dropdown.classList.toggle('open');
    if (dropdown.classList.contains('open')) {
      searchInput.value = '';
      searchInput.focus();
      renderOptions();
    }
  });

  searchInput.addEventListener('input', () => renderOptions(searchInput.value));

  // Close on outside click
  document.addEventListener('click', (e) => {
    if (!container.contains(e.target)) dropdown.classList.remove('open');
  });

  // Store reset function
  container._reset = (newSelected) => {
    selected = new Set(newSelected);
    renderTrigger();
    renderOptions();
  };

  renderTrigger();
  renderOptions();
}

function resetMultiSelect(containerId, newSelected) {
  const container = document.getElementById(containerId);
  if (container._reset) container._reset(newSelected);
}

// ─── KPI Cards ──────────────────────────────────────────────────────────────
function renderKPIs() {
  const data = State.filteredData;
  if (!data.length) return;

  const amounts = data.map(r => r.a);
  const avgPrice = mean(amounts);

  // Brand averages
  const brandTotals = {};
  const brandCounts = {};
  data.forEach(r => {
    brandTotals[r.b] = (brandTotals[r.b] || 0) + r.a;
    brandCounts[r.b] = (brandCounts[r.b] || 0) + 1;
  });
  const brandAvgs = {};
  for (const b in brandTotals) brandAvgs[b] = brandTotals[b] / brandCounts[b];
  const sorted = Object.entries(brandAvgs).sort((a, b) => a[1] - b[1]);

  const cheapest = sorted[0];
  const expensive = sorted[sorted.length - 1];
  const spread = expensive[1] - cheapest[1];

  // Coverage
  const statesSet = new Set(data.map(r => r.s).filter(Boolean));
  const districtsSet = new Set(data.map(r => r.d).filter(Boolean));

  // BIS %
  const bisCount = data.filter(r => r.q === 'BIS').length;
  const knownQ = data.filter(r => r.q === 'BIS' || r.q === 'NonBIS').length;

  document.getElementById('kpi-avg-price').textContent = formatINR(avgPrice);
  document.getElementById('kpi-avg-price-sub').textContent = `Across ${Object.keys(brandAvgs).length} brands, ${data.length} entries`;

  document.getElementById('kpi-spread').textContent = formatINR(spread);
  document.getElementById('kpi-spread-sub').textContent = `${cheapest[0]} to ${expensive[0]}`;

  document.getElementById('kpi-cheapest').textContent = formatINR(cheapest[1]);
  document.getElementById('kpi-cheapest-sub').textContent = cheapest[0];

  document.getElementById('kpi-expensive').textContent = formatINR(expensive[1]);
  document.getElementById('kpi-expensive-sub').textContent = expensive[0];

  document.getElementById('kpi-coverage').textContent = `${statesSet.size} / ${districtsSet.size}`;
  document.getElementById('kpi-coverage-sub').textContent = 'States / Districts';

  document.getElementById('kpi-datapoints').textContent = data.length.toLocaleString();
  document.getElementById('kpi-datapoints-sub').textContent = `Across ${Object.keys(brandAvgs).length} brands`;
}

// ─── Chart 1: Price Trend Line ──────────────────────────────────────────────
function renderPriceTrend() {
  const ctx = document.getElementById('canvas-price-trend');
  if (State.charts.priceTrend) State.charts.priceTrend.destroy();

  const data = State.filteredData;
  const brands = State.filters.brands;
  const isWeekView = State.filters.timePeriod === 'week' && State.filters.months.length > 0;

  let labels, trendKey, subtitleText;

  if (isWeekView) {
    // Weekly view: X-axis = week labels within selected months
    // Build sorted week labels from filtered data
    const weekSet = new Set();
    data.forEach(r => {
      if (State.filters.months.includes(r.m) && r.wl) weekSet.add(r.wl);
    });
    // Sort by month order then week number
    labels = [...weekSet].sort((a, b) => {
      const [mA, wA] = [a.substring(0, a.lastIndexOf('-W')), parseInt(a.split('-W')[1])];
      const [mB, wB] = [b.substring(0, b.lastIndexOf('-W')), parseInt(b.split('-W')[1])];
      const mi = monthIndex(mA) - monthIndex(mB);
      return mi !== 0 ? mi : wA - wB;
    });
    trendKey = 'wl'; // use week label as grouping key
    subtitleText = `Weekly average within ${State.filters.months.join(', ')}`;
  } else {
    // Monthly view (default)
    labels = State.filters.months.length > 0
      ? MONTH_ORDER.filter(m => State.filters.months.includes(m))
      : MONTH_ORDER;
    trendKey = 'm'; // use month as grouping key
    subtitleText = 'Monthly average dealer landing price (₹/MT, Excl. GST)';
  }

  // Update subtitle
  const subtitleEl = document.querySelector('#chart-price-trend .chart-subtitle');
  if (subtitleEl) subtitleEl.textContent = subtitleText;

  // Compute brand trends from filtered data using the chosen key
  const trends = {};
  data.forEach(r => {
    const key = r[trendKey];
    if (!key) return;
    if (!trends[r.b]) trends[r.b] = {};
    if (!trends[r.b][key]) trends[r.b][key] = [];
    trends[r.b][key].push(r.a);
  });

  // Market average line
  const marketByPeriod = {};
  data.forEach(r => {
    const key = r[trendKey];
    if (!key) return;
    if (!marketByPeriod[key]) marketByPeriod[key] = [];
    marketByPeriod[key].push(r.a);
  });

  const datasets = [];

  // Market average (dashed)
  datasets.push({
    label: 'Market Average',
    data: labels.map(l => marketByPeriod[l] ? Math.round(mean(marketByPeriod[l])) : null),
    borderColor: '#94a3b8',
    borderDash: [5, 5],
    borderWidth: 1.5,
    pointRadius: 0,
    spanGaps: true,
    order: 10,
  });

  // Brand lines
  brands.forEach((brand, i) => {
    const isBenchmark = brand === BENCHMARK_BRAND && BENCHMARK_BRAND;
    datasets.push({
      label: brand,
      data: labels.map(l => {
        const vals = trends[brand]?.[l];
        return vals ? Math.round(mean(vals)) : null;
      }),
      borderColor: brandColor(brand, i),
      backgroundColor: isBenchmark ? 'rgba(99,102,241,0.1)' : 'transparent',
      borderWidth: isBenchmark ? 3 : 2,
      pointRadius: isBenchmark ? 5 : 3,
      pointBackgroundColor: brandColor(brand, i),
      tension: 0.3,
      spanGaps: true,
      order: isBenchmark ? 0 : i + 1,
    });
  });

  State.charts.priceTrend = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#94a3b8', font: { family: 'Inter', size: 11 }, boxWidth: 12, padding: 12 }
        },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#f8fafc',
          bodyColor: '#94a3b8',
          borderColor: '#334155',
          borderWidth: 1,
          padding: 10,
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${formatINR(ctx.parsed.y)}`
          }
        },
      },
      scales: {
        x: {
          ticks: { color: '#64748b', font: { size: 11 } },
          grid: { color: 'rgba(51,65,85,0.3)' }
        },
        y: {
          ticks: {
            color: '#64748b',
            font: { size: 11 },
            callback: v => '₹' + (v / 1000).toFixed(0) + 'k'
          },
          grid: { color: 'rgba(51,65,85,0.3)' },
          title: { display: true, text: 'Price (₹/MT, Excl. GST)', color: '#64748b', font: { size: 11 } }
        }
      }
    }
  });
}

// ─── Chart 2: Price Distribution Box Plot ───────────────────────────────────
function renderPriceDistribution() {
  const ctx = document.getElementById('canvas-price-dist');
  if (State.charts.priceDistribution) State.charts.priceDistribution.destroy();

  const data = State.filteredData;
  const brands = State.filters.brands;

  // Collect price arrays per brand
  const brandPrices = {};
  data.forEach(r => {
    if (brands.includes(r.b)) {
      if (!brandPrices[r.b]) brandPrices[r.b] = [];
      brandPrices[r.b].push(r.a);
    }
  });

  State.charts.priceDistribution = new Chart(ctx, {
    type: 'boxplot',
    data: {
      labels: brands,
      datasets: [{
        label: 'Price Distribution',
        data: brands.map(b => brandPrices[b] || []),
        backgroundColor: brands.map((b, i) =>
          b === BENCHMARK_BRAND ? 'rgba(99,102,241,0.3)' : `${brandColor(b, i)}33`
        ),
        borderColor: brands.map((b, i) => brandColor(b, i)),
        borderWidth: brands.map(b => b === BENCHMARK_BRAND ? 2.5 : 1.5),
        outlierRadius: 2,
        itemRadius: 0,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#f8fafc',
          bodyColor: '#94a3b8',
          borderColor: '#334155',
          borderWidth: 1,
        }
      },
      scales: {
        x: {
          ticks: { color: '#64748b', font: { size: 11 }, maxRotation: 45 },
          grid: { display: false }
        },
        y: {
          ticks: {
            color: '#64748b',
            font: { size: 11 },
            callback: v => '₹' + (v / 1000).toFixed(0) + 'k'
          },
          grid: { color: 'rgba(51,65,85,0.3)' },
          title: { display: true, text: 'Price (₹/MT, Excl. GST)', color: '#64748b', font: { size: 11 } }
        }
      }
    }
  });
}

// ─── Chart 3: State Heatmap ─────────────────────────────────────────────────
function renderHeatmap() {
  const container = document.getElementById('heatmap-container');
  const data = State.filteredData;
  const brands = State.filters.brands;

  // Compute state × brand average
  const agg = {};
  data.forEach(r => {
    if (brands.includes(r.b)) {
      const key = r.s + '|' + r.b;
      if (!agg[key]) agg[key] = [];
      agg[key].push(r.a);
    }
  });

  const states = [...new Set(data.map(r => r.s))].filter(Boolean).sort();
  if (!states.length || !brands.length) {
    container.innerHTML = '<p style="color:var(--color-text-muted);padding:16px;">No data for current filters.</p>';
    return;
  }

  // Build matrix and find min/max
  const matrix = {};
  let gMin = Infinity, gMax = -Infinity;
  states.forEach(s => {
    matrix[s] = {};
    brands.forEach(b => {
      const vals = agg[s + '|' + b];
      if (vals && vals.length) {
        const avg = Math.round(mean(vals));
        matrix[s][b] = avg;
        if (avg < gMin) gMin = avg;
        if (avg > gMax) gMax = avg;
      }
    });
  });

  const range = gMax - gMin || 1;

  let html = '<table class="heatmap-table"><thead><tr><th style="text-align:left;">State</th>';
  brands.forEach(b => {
    const cls = b === BENCHMARK_BRAND ? ' benchmark' : '';
    html += `<th class="${cls}">${b}</th>`;
  });
  html += '</tr></thead><tbody>';

  states.forEach(s => {
    html += `<tr><td class="state-cell">${titleCase(s)}</td>`;
    brands.forEach(b => {
      const val = matrix[s]?.[b];
      if (val != null) {
        const level = Math.max(1, Math.min(6, Math.ceil(((val - gMin) / range) * 6)));
        html += `<td class="heatmap-cell" data-level="${level}" title="${b}: ${formatINR(val)} in ${titleCase(s)}">${formatINR(val)}</td>`;
      } else {
        html += '<td class="heatmap-empty">—</td>';
      }
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}

// ─── Chart 5: Dealer Count Grouped Bar ──────────────────────────────────────
function renderDealerCount() {
  const ctx = document.getElementById('canvas-dealer-count');
  if (State.charts.dealerCount) State.charts.dealerCount.destroy();

  const data = State.filteredData;
  const brands = State.filters.brands;

  // Count unique dealers per state×brand
  const sets = {};
  data.forEach(r => {
    if (brands.includes(r.b)) {
      const key = r.s + '|' + r.b;
      if (!sets[key]) sets[key] = new Set();
      sets[key].add(r.c);
    }
  });

  // Top 10 states by total dealer count
  const stateTotals = {};
  for (const [key, set] of Object.entries(sets)) {
    const state = key.split('|')[0];
    stateTotals[state] = (stateTotals[state] || 0) + set.size;
  }
  const topStates = Object.entries(stateTotals)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(e => e[0]);

  const datasets = brands.map((brand, i) => ({
    label: brand,
    data: topStates.map(s => sets[s + '|' + brand]?.size || 0),
    backgroundColor: brandColor(brand, i),
    borderRadius: 3,
  }));

  State.charts.dealerCount = new Chart(ctx, {
    type: 'bar',
    data: { labels: topStates.map(titleCase), datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#94a3b8', font: { family: 'Inter', size: 11 }, boxWidth: 12, padding: 12 }
        },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#f8fafc',
          bodyColor: '#94a3b8',
          borderColor: '#334155',
          borderWidth: 1,
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y} dealers`
          }
        }
      },
      scales: {
        x: {
          ticks: { color: '#64748b', font: { size: 11 }, maxRotation: 45 },
          grid: { display: false }
        },
        y: {
          ticks: { color: '#64748b', font: { size: 11 } },
          grid: { color: 'rgba(51,65,85,0.3)' },
          title: { display: true, text: 'Unique Dealers', color: '#64748b', font: { size: 11 } }
        }
      }
    }
  });
}

// ─── Raw Data Table ─────────────────────────────────────────────────────────
function renderTable() {
  const data = State.filteredData;
  const { page, perPage, sortCol, sortDir, search } = State.table;

  // Filter by search
  let filtered = data;
  if (search) {
    const q = search.toLowerCase();
    filtered = data.filter(r =>
      (r.b && r.b.toLowerCase().includes(q)) ||
      (r.c && r.c.toLowerCase().includes(q)) ||
      (r.s && r.s.toLowerCase().includes(q)) ||
      (r.d && r.d.toLowerCase().includes(q))
    );
  }

  // Sort
  filtered = [...filtered].sort((a, b) => {
    let va = a[sortCol] ?? '';
    let vb = b[sortCol] ?? '';
    if (sortCol === 'a') return sortDir === 'asc' ? va - vb : vb - va;
    if (sortCol === 'm') {
      return sortDir === 'asc'
        ? monthIndex(va) - monthIndex(vb)
        : monthIndex(vb) - monthIndex(va);
    }
    va = String(va);
    vb = String(vb);
    const cmp = va.localeCompare(vb);
    return sortDir === 'asc' ? cmp : -cmp;
  });

  // Paginate
  const total = filtered.length;
  const totalPages = Math.ceil(total / perPage) || 1;
  const clamped = Math.min(page, totalPages);
  const start = (clamped - 1) * perPage;
  const pageData = filtered.slice(start, start + perPage);

  // Render rows
  const tbody = document.getElementById('table-body');
  tbody.innerHTML = pageData.map(r => {
    const isBenchmark = r.b === BENCHMARK_BRAND ? ' benchmark-row' : '';
    return `<tr class="${isBenchmark}">
      <td>${r.b || ''}</td>
      <td>${formatINR(r.a)}</td>
      <td>${r.q === 'BIS' ? '<span style="color:#22c55e">BIS</span>' : r.q === 'NonBIS' ? '<span style="color:#ef4444">Non-BIS</span>' : '—'}</td>
      <td>${r.t || '—'}</td>
      <td>${titleCase(r.s)}</td>
      <td>${titleCase(r.d)}</td>
      <td>${r.m || ''}</td>
      <td style="font-size:12px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r.c || ''}">${r.c || ''}</td>
    </tr>`;
  }).join('');

  // Pagination info
  document.getElementById('pagination-info').textContent =
    `Showing ${start + 1}–${Math.min(start + perPage, total)} of ${total.toLocaleString()}`;

  // Pagination buttons
  const btnsContainer = document.getElementById('pagination-btns');
  btnsContainer.innerHTML = '';

  const addBtn = (label, pg, disabled = false, active = false) => {
    const btn = document.createElement('button');
    btn.textContent = label;
    btn.disabled = disabled;
    if (active) btn.className = 'active';
    btn.addEventListener('click', () => { State.table.page = pg; renderTable(); });
    btnsContainer.appendChild(btn);
  };

  addBtn('←', clamped - 1, clamped <= 1);

  // Show max 7 page buttons
  let startP = Math.max(1, clamped - 3);
  let endP = Math.min(totalPages, startP + 6);
  if (endP - startP < 6) startP = Math.max(1, endP - 6);

  for (let i = startP; i <= endP; i++) {
    addBtn(i, i, false, i === clamped);
  }

  addBtn('→', clamped + 1, clamped >= totalPages);

  // Update sort indicators
  document.querySelectorAll('.data-table thead th').forEach(th => {
    th.classList.toggle('sorted', th.dataset.col === sortCol);
    const icon = th.querySelector('.sort-icon');
    if (icon) icon.textContent = th.dataset.col === sortCol ? (sortDir === 'asc' ? '↑' : '↓') : '↕';
  });
}

function initTableListeners() {
  // Search
  document.getElementById('table-search').addEventListener('input', (e) => {
    State.table.search = e.target.value;
    State.table.page = 1;
    renderTable();
  });

  // Rows per page
  document.getElementById('rows-per-page').addEventListener('change', (e) => {
    State.table.perPage = parseInt(e.target.value);
    State.table.page = 1;
    renderTable();
  });

  // Sort headers
  document.querySelectorAll('.data-table thead th[data-col]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (State.table.sortCol === col) {
        State.table.sortDir = State.table.sortDir === 'asc' ? 'desc' : 'asc';
      } else {
        State.table.sortCol = col;
        State.table.sortDir = 'asc';
      }
      renderTable();
    });
  });

  // Export CSV
  document.getElementById('export-csv').addEventListener('click', exportCSV);
}

function exportCSV() {
  const data = State.filteredData;
  const headers = ['Brand', 'Amount (INR, Excl. GST)', 'Quality', 'Delivery', 'State', 'District', 'Month', 'Company'];
  const rows = data.map(r => [
    r.b, r.a, r.q || '', r.t || '', r.s, r.d || '', r.m, r.c || ''
  ]);
  const csv = [headers, ...rows].map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `market-feedback-${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ─── Sidebar Navigation ────────────────────────────────────────────────────
function initSidebar() {
  document.querySelectorAll('.sidebar-nav a').forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const target = document.getElementById(link.dataset.section);
      if (target) {
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        // Update active state
        document.querySelectorAll('.sidebar-nav a').forEach(l => l.classList.remove('active'));
        link.classList.add('active');
      }
    });
  });
}

// ─── Tab Switching ──────────────────────────────────────────────────────────
function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tabId);
  });
  document.querySelectorAll('.tab-content').forEach(tc => {
    tc.classList.remove('active');
  });
  const target = document.getElementById('tab-' + tabId);
  if (target) target.classList.add('active');

  const sidebar = document.querySelector('.sidebar');
  const main = document.querySelector('.main-content');
  if (tabId === 'price-positioning') {
    sidebar.style.display = 'none';
    main.style.marginLeft = '0';
    if (!State.pricePositioning.initialized) {
      initPricePositioning();
    }
  } else {
    sidebar.style.display = 'flex';
    main.style.marginLeft = '240px';
  }
}

// ─── Price Positioning ──────────────────────────────────────────────────────
function initPricePositioning() {
  State.pricePositioning.initialized = true;
  const regionSelect = document.getElementById('pp-region-filter');

  // Only show regions that have data
  const availableStates = new Set(DASHBOARD_DATA.filterableData.map(r => r.s));
  const regionEntries = Object.entries(REGION_STATE_MAP).filter(([_, states]) =>
    states.some(s => availableStates.has(s))
  );

  regionSelect.innerHTML = '';
  regionEntries.forEach(([region]) => {
    const opt = document.createElement('option');
    opt.value = region;
    opt.textContent = region;
    regionSelect.appendChild(opt);
  });

  const defaultRegion = regionEntries.find(([r]) => r === 'Delhi')?.[0]
    || regionEntries[0]?.[0];
  if (defaultRegion) {
    regionSelect.value = defaultRegion;
    State.pricePositioning.region = defaultRegion;
  }

  regionSelect.addEventListener('change', () => {
    State.pricePositioning.region = regionSelect.value;
    renderPricePositioning();
  });

  renderPricePositioning();
}

function aggregatePPData(region) {
  const states = REGION_STATE_MAP[region];
  if (!states) return null;

  const stateSet = new Set(states);
  const latestMonth = MONTH_ORDER[MONTH_ORDER.length - 1];

  // Filter to selected state(s) and latest month
  const stateData = DASHBOARD_DATA.filterableData.filter(
    r => stateSet.has(r.s) && r.m === latestMonth
  );
  if (!stateData.length) return null;

  // Collect prices per data brand
  const brandPrices = {};
  stateData.forEach(r => {
    if (!brandPrices[r.b]) brandPrices[r.b] = [];
    brandPrices[r.b].push(r.a);
  });

  // Build entries from curated brand config
  const entries = [];
  PP_BRAND_CONFIG.forEach(cfg => {
    const allPrices = [];
    cfg.dataBrands.forEach(db => {
      if (brandPrices[db]) allPrices.push(...brandPrices[db]);
    });
    if (!allPrices.length) return;

    entries.push({
      key: cfg.key,
      label: cfg.label,
      type: cfg.type,
      dealerPrice: Math.round(mean(allPrices)),
      dataPoints: allPrices.length,
    });
  });

  // Get benchmark price
  const benchmarkEntry = entries.find(e => e.type === 'benchmark');
  const benchmarkPrice = benchmarkEntry ? benchmarkEntry.dealerPrice : null;

  // Calculate gap vs benchmark (positive = cheaper than TISCON)
  entries.forEach(e => {
    if (e.type === 'benchmark') {
      e.gap = null;
    } else if (benchmarkPrice != null) {
      e.gap = benchmarkPrice - e.dealerPrice;
    } else {
      e.gap = null;
    }
  });

  return { region, month: latestMonth, entries, benchmarkPrice };
}

function renderPricePositioning() {
  const region = State.pricePositioning.region;
  const data = aggregatePPData(region);

  const noData = document.getElementById('pp-no-data');
  const content = document.getElementById('pp-content');
  const regionName = document.getElementById('pp-region-name');

  regionName.textContent = region || '';

  if (!data || !data.entries.length) {
    noData.style.display = 'block';
    content.style.display = 'none';
    return;
  }

  noData.style.display = 'none';
  content.style.display = 'grid';

  document.getElementById('pp-date-box').textContent = `As on ${data.month}`;

  renderPPTable(data);
  renderPPChart(data);
}

function renderPPTable(data) {
  const tbody = document.getElementById('pp-table-body');
  tbody.innerHTML = data.entries.map(entry => {
    let rowClass = '';
    if (entry.type === 'benchmark') rowClass = 'pp-row-benchmark';
    else if (entry.type === 'jsw') rowClass = 'pp-row-jsw';

    let gapHtml = '';
    if (entry.type === 'benchmark') {
      gapHtml = '<span class="pp-benchmark-badge">BENCHMARK</span>';
    } else if (entry.gap != null) {
      const absGap = Math.abs(entry.gap);
      if (entry.gap > 0) {
        gapHtml = `<span class="pp-gap-negative">&#9660; ${formatINR(absGap)}</span>`;
      } else if (entry.gap < 0) {
        gapHtml = `<span class="pp-gap-positive">&#9650; ${formatINR(absGap)}</span>`;
      } else {
        gapHtml = '<span style="color:var(--color-text-muted);">--</span>';
      }
    } else {
      gapHtml = '<span style="color:var(--color-text-muted);">N/A</span>';
    }

    return `<tr class="${rowClass}">
      <td>${entry.label}</td>
      <td>${formatINR(entry.dealerPrice)}</td>
      <td>${gapHtml}</td>
    </tr>`;
  }).join('');
}

function renderPPChart(data) {
  const ctx = document.getElementById('canvas-pp-bar');
  if (State.pricePositioning.chart) {
    State.pricePositioning.chart.destroy();
  }

  // Exclude benchmark from bars (shown as dashed line instead)
  const chartEntries = data.entries.filter(e => e.type !== 'benchmark');
  const labels = chartEntries.map(e => e.label);
  const prices = chartEntries.map(e => e.dealerPrice);
  const barColors = chartEntries.map(e => {
    if (e.type === 'jsw') return '#ef4444';
    return '#475569';
  });

  // Custom plugin for data labels above bars
  const dataLabelPlugin = {
    id: 'ppDataLabels',
    afterDatasetsDraw(chart) {
      const { ctx: c, data: d } = chart;
      const ds = d.datasets[0];
      const meta = chart.getDatasetMeta(0);
      c.save();
      c.textAlign = 'center';
      c.font = '600 11px Inter';
      meta.data.forEach((bar, i) => {
        const val = ds.data[i];
        const label = (val / 1000).toFixed(1) + 'K';
        c.fillStyle = chartEntries[i].type === 'jsw' ? '#ef4444' : '#94a3b8';
        c.fillText(label, bar.x, bar.y - 8);
      });
      c.restore();
    }
  };

  // Benchmark dashed line annotation
  const annotations = {};
  if (data.benchmarkPrice) {
    annotations.benchmarkLine = {
      type: 'line',
      yMin: data.benchmarkPrice,
      yMax: data.benchmarkPrice,
      borderColor: '#fbbf24',
      borderWidth: 2,
      borderDash: [6, 4],
      label: {
        display: true,
        content: `TISCON ${(data.benchmarkPrice / 1000).toFixed(1)}K`,
        position: 'end',
        backgroundColor: 'transparent',
        color: '#fbbf24',
        font: { size: 11, weight: '600', family: 'Inter' },
        padding: 4,
      }
    };
  }

  const allPrices = [...prices];
  if (data.benchmarkPrice) allPrices.push(data.benchmarkPrice);
  const minPrice = Math.min(...allPrices);
  const maxPrice = Math.max(...allPrices);

  State.pricePositioning.chart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Dealer Price (₹/MT)',
        data: prices,
        backgroundColor: barColors,
        borderRadius: 4,
        maxBarThickness: 48,
      }]
    },
    plugins: [dataLabelPlugin],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 24 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#f8fafc',
          bodyColor: '#94a3b8',
          borderColor: '#334155',
          borderWidth: 1,
          padding: 10,
          callbacks: {
            label: (c) => `Dealer Price: ${formatINR(c.parsed.y)}`
          }
        },
        annotation: { annotations },
      },
      scales: {
        x: {
          ticks: { color: '#64748b', font: { size: 11, family: 'Inter' }, maxRotation: 35 },
          grid: { display: false }
        },
        y: {
          ticks: {
            color: '#64748b',
            font: { size: 11 },
            callback: v => '₹' + (v / 1000).toFixed(0) + 'K'
          },
          grid: { color: 'rgba(51,65,85,0.3)' },
          beginAtZero: false,
          min: Math.floor((minPrice * 0.90) / 1000) * 1000,
          max: Math.ceil((maxPrice * 1.06) / 1000) * 1000,
        }
      }
    }
  });
}

// ─── Initialize ─────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initFilters();
  initTableListeners();
  initSidebar();
  applyFilters();

  // Remove loading overlay
  setTimeout(() => {
    document.getElementById('loading').classList.add('hidden');
  }, 300);
});
