const MONTHS_SHORT = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const MONTHS_LONG = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const fmtAR = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 2 });
const fmtMoney = new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', maximumFractionDigits: 0 });
const fmtInt = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });

const QUARTERS = {
  Q1: [0, 1, 2],
  Q2: [3, 4, 5],
  Q3: [6, 7, 8],
  Q4: [9, 10, 11],
};

const state = {
  data: null,
  year: null,
  month: 'all',
  quarter: 'all',
  search: '',
  type: 'all',
  category: 'all',
  clientMetric: 'facturacion',
  topN: 15,
  productMetric: 'kg',
  tab: 'resumen',
  sortBy: { clientes: { key: 'totalFacturacion', dir: 'desc' }, tipos: { key: 'facturacion', dir: 'desc' }, productos: { key: 'total', dir: 'desc' } },
};

const charts = {};

function $(sel) { return document.querySelector(sel); }
function $$(sel) { return [...document.querySelectorAll(sel)]; }

async function loadData() {
  try {
    if (window.__DATA__) {
      state.data = window.__DATA__;
    } else {
      const r = await fetch('data.json');
      if (!r.ok) throw new Error('No se pudo cargar data.json');
      state.data = await r.json();
    }
    init();
  } catch (e) {
    const err = $('#error');
    err.hidden = false;
    err.innerHTML = `<div><strong>Error:</strong> ${e.message}<br/><br/>
      Ejecutá <code style="background:#000;padding:.2rem .4rem;border-radius:6px;">node build-data.js</code> en la carpeta del proyecto para regenerar los datos.</div>`;
    $('#loader').hidden = true;
  }
}

function init() {
  const years = Object.keys(state.data.years).sort();
  const ySel = $('#yearSelect');
  ySel.innerHTML = years.map(y => `<option value="${y}">${y}</option>`).join('');
  state.year = years[years.length - 1];
  ySel.value = state.year;

  const mSel = $('#monthSelect');
  mSel.innerHTML = '<option value="all">Todos</option>' +
    MONTHS_LONG.map((m, i) => `<option value="${i}">${m}</option>`).join('');

  wireEvents();
  refreshFilters();
  updateFilterVisibility();
  render();

  $('#loader').hidden = true;
  $('#sourceInfo').textContent = `Fuentes: ${state.data.sources.map(s => s.file).join(' · ')}`;
}

function updateFilterVisibility() {
  const hideType = state.tab === 'productos';
  $('#typeFilterLabel').style.display = hideType ? 'none' : '';
  // Si se ocultó el filtro y tenía un valor, resetearlo
  if (hideType && state.type !== 'all') {
    state.type = 'all';
    $('#typeSelect').value = 'all';
  }
}

function wireEvents() {
  $$('#tabs button').forEach(b => b.addEventListener('click', () => {
    state.tab = b.dataset.tab;
    $$('#tabs button').forEach(x => x.classList.toggle('active', x === b));
    $$('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `tab-${state.tab}`));
    updateFilterVisibility();
    render();
  }));

  $('#yearSelect').addEventListener('change', e => { state.year = e.target.value; refreshFilters(); render(); });
  $('#quarterSelect').addEventListener('change', e => { state.quarter = e.target.value; render(); });
  $('#monthSelect').addEventListener('change', e => { state.month = e.target.value; render(); });
  $('#searchInput').addEventListener('input', e => { state.search = e.target.value.toLowerCase(); render(); });
  $('#typeSelect').addEventListener('change', e => { state.type = e.target.value; render(); });
  $('#clientMetric').addEventListener('change', e => { state.clientMetric = e.target.value; render(); });
  $('#topN').addEventListener('change', e => { state.topN = Number(e.target.value); render(); });
  $('#productMetric').addEventListener('change', e => { state.productMetric = e.target.value; render(); });
  $('#categorySelect').addEventListener('change', e => { state.category = e.target.value; render(); });
}

function currentYearData() { return state.data.years[state.year] || {}; }

function refreshFilters() {
  const y = currentYearData();
  const types = new Set();
  (y.byClientType || []).forEach(c => types.add(c.tipo || 'Sin clasificar'));
  // Si el año actual no tiene tipos, buscar en otros años
  if (types.size === 0) {
    for (const yr of Object.values(state.data.years)) {
      (yr.byClientType || []).forEach(c => types.add(c.tipo || 'Sin clasificar'));
    }
  }
  const tSel = $('#typeSelect');
  tSel.innerHTML = '<option value="all">Todos</option>' +
    [...types].sort().map(t => `<option value="${t}">${t}</option>`).join('');
  if (!types.has(state.type)) state.type = 'all';
  tSel.value = state.type;
  tSel.disabled = types.size === 0;

  const cats = new Set();
  (y.products || []).forEach(p => cats.add(p.categoria || '—'));
  const cSel = $('#categorySelect');
  cSel.innerHTML = '<option value="all">Todas</option>' +
    [...cats].sort().map(c => `<option value="${c}">${c}</option>`).join('');
  if (!cats.has(state.category)) state.category = 'all';
  cSel.value = state.category;
}

function monthFilterArray() {
  const base = state.quarter !== 'all' && QUARTERS[state.quarter]
    ? QUARTERS[state.quarter]
    : Array.from({ length: 12 }, (_, i) => i);
  if (state.month === 'all') return base;
  const m = Number(state.month);
  return base.includes(m) ? [m] : [];
}

// === Render ===
function render() {
  const y = currentYearData();
  const months = monthFilterArray();
  renderSearchMatches(y, months);
  renderResumen();
  renderAnalisis();
  renderMensual();
  renderClientes();
  renderTipos();
  renderProductos();
}

function renderResumen() {
  const y = currentYearData();
  const mg = y.monthlyGeneral || [];
  const months = monthFilterArray();
  const hasClientFilter = state.search || state.type !== 'all';

  if (!hasClientFilter) {
    // Sin filtros de cliente/búsqueda: usar datos agregados de monthlyGeneral
    const filtered = mg.filter(r => months.includes(r.monthIndex));
    const sum = (key) => filtered.reduce((a, r) => a + (Number(r[key]) || 0), 0);
    const kpis = [
      { label: 'Facturación', value: fmtMoney.format(sum('Facturación')) },
      { label: 'Kg vendidos', value: fmtAR.format(sum('Kg Vendidos')) },
      { label: 'Pedidos', value: fmtInt.format(sum('Pedidos del mes')) },
      { label: 'Clientes activos (prom.)', value: filtered.length ? fmtInt.format(sum('Clientes activos') / filtered.length) : '—' },
      { label: 'Nuevos clientes', value: fmtInt.format(sum('Nuevos clientes')) },
      { label: 'Reactivados', value: fmtInt.format(sum('Clientes reactivados')) },
    ];
    const prevYear = String(Number(state.year) - 1);
    if (state.data.years[prevYear]) {
      const mgPrev = state.data.years[prevYear].monthlyGeneral || [];
      const prevFiltered = mgPrev.filter(r => months.includes(r.monthIndex));
      const sumPrev = (key) => prevFiltered.reduce((a, r) => a + (Number(r[key]) || 0), 0);
      const deltaSpecs = [
        { idx: 0, key: 'Facturación' },
        { idx: 1, key: 'Kg Vendidos' },
        { idx: 2, key: 'Pedidos del mes' },
      ];
      deltaSpecs.forEach(({ idx, key }) => {
        const prev = sumPrev(key);
        if (prev > 0) {
          const delta = ((sum(key) - prev) / prev) * 100;
          kpis[idx].delta = `${delta >= 0 ? '▲' : '▼'} ${fmtAR.format(Math.abs(delta))}% vs ${prevYear}`;
          kpis[idx].deltaClass = delta >= 0 ? 'up' : 'down';
        }
      });
    }
    renderKpiGrid(kpis);

    const labels = months.map(i => MONTHS_SHORT[i]);
    const pick = (key) => months.map(i => {
      const row = mg.find(r => r.monthIndex === i);
      return row ? Number(row[key]) || 0 : 0;
    });
    const prevMgData = state.data.years[prevYear]?.monthlyGeneral || null;
    const pickPrev = (key) => months.map(i => {
      if (!prevMgData) return 0;
      const row = prevMgData.find(r => r.monthIndex === i);
      return row ? Number(row[key]) || 0 : 0;
    });
    const prevDs = (label, key, color) => prevMgData
      ? [{ label: `${label} ${prevYear}`, data: pickPrev(key), color, muted: true }]
      : [];
    drawBar('chartFacturacion', labels, [
      { label: `Facturación ${state.year}`, data: pick('Facturación'), color: '#7c8cff' },
      ...prevDs('Facturación', 'Facturación', '#7c8cff'),
    ], { money: true });
    drawBar('chartKg', labels, [
      { label: `Kg ${state.year}`, data: pick('Kg Vendidos'), color: '#4bd6b0' },
      ...prevDs('Kg', 'Kg Vendidos', '#4bd6b0'),
    ]);
    drawBar('chartPedidos', labels, [
      { label: `Pedidos ${state.year}`, data: pick('Pedidos del mes'), color: '#ffb457' },
      ...prevDs('Pedidos', 'Pedidos del mes', '#ffb457'),
    ]);
    drawLine('chartClientes', labels, [
      { label: 'Activos', data: pick('Clientes activos'), color: '#7c8cff' },
      { label: 'Nuevos', data: pick('Nuevos clientes'), color: '#4bd6b0' },
      { label: 'Reactivados', data: pick('Clientes reactivados'), color: '#ffb457' },
    ]);
  } else {
    // Con filtro de búsqueda o tipo de cliente: recalcular desde datos detallados
    const typeMap = new Map();
    (y.byClientType || []).forEach(c => typeMap.set(c.cliente, c.tipo || 'Sin clasificar'));

    // Filtrar clientes según búsqueda y tipo
    const allClients = (y.byClient || []);
    let clients = allClients.slice();
    if (state.type !== 'all') {
      clients = clients.filter(c => (typeMap.get(c.cliente) || 'Sin clasificar') === state.type);
    }
    if (state.search) {
      clients = clients.filter(c => c.cliente.toLowerCase().includes(state.search));
    }

    // Agregar datos mensuales desde los clientes filtrados
    const monthlyAgg = months.map(i => {
      let fact = 0, uni = 0, ped = 0;
      const clientesActivos = new Set();
      clients.forEach(c => {
        const m = c.byMonth[i];
        if (m) {
          fact += m.facturacion || 0;
          uni += m.unidades || 0;
          ped += m.pedidos || 0;
          if ((m.unidades || 0) > 0 || (m.pedidos || 0) > 0 || (m.facturacion || 0) > 0) {
            clientesActivos.add(c.cliente);
          }
        }
      });
      return { monthIndex: i, facturacion: fact, unidades: uni, pedidos: ped, clientesActivos: clientesActivos.size };
    });

    // Totales globales (sin filtro) para calcular proporciones
    const mgFiltered = mg.filter(r => months.includes(r.monthIndex));
    const globalUni = months.map(i =>
      allClients.reduce((a, c) => a + (c.byMonth[i]?.unidades || 0), 0)
    );
    const globalFact = mgFiltered.map(r => Number(r['Facturación']) || 0);
    const globalKg = mgFiltered.map(r => Number(r['Kg Vendidos']) || 0);

    // Si byClient no tiene facturación, estimar proporcionalmente por unidades
    const directFact = monthlyAgg.reduce((a, m) => a + m.facturacion, 0);
    let monthlyFact, monthlyKg;
    if (directFact === 0 && clients.length > 0) {
      // Estimar facturación: (unidades filtradas / unidades totales) × facturación total
      monthlyFact = months.map((mi, idx) => {
        const filtUni = monthlyAgg[idx].unidades;
        const totUni = globalUni[idx];
        return totUni > 0 ? (filtUni / totUni) * globalFact[idx] : 0;
      });
    } else {
      monthlyFact = monthlyAgg.map(m => m.facturacion);
    }

    // Kg: productos no están vinculados a clientes, estimar por proporción de unidades
    const matchedProducts = (y.products || []).filter(p =>
      state.search &&
      (p.producto.toLowerCase().includes(state.search) ||
       (p.categoria || '').toLowerCase().includes(state.search))
    );
    if (matchedProducts.length > 0) {
      // La búsqueda coincide con productos: usar Kg directos
      monthlyKg = months.map(i =>
        matchedProducts.reduce((a, p) => a + (p.byMonth[i]?.kg || 0), 0)
      );
    } else {
      // Sin match de productos (búsqueda de cliente o filtro tipo): estimar Kg proporcionalmente
      monthlyKg = months.map((mi, idx) => {
        const filtUni = monthlyAgg[idx].unidades;
        const totUni = globalUni[idx];
        return totUni > 0 ? (filtUni / totUni) * globalKg[idx] : 0;
      });
    }

    const totalFact = monthlyFact.reduce((a, v) => a + v, 0);
    const totalUni = monthlyAgg.reduce((a, m) => a + m.unidades, 0);
    const totalPed = monthlyAgg.reduce((a, m) => a + m.pedidos, 0);
    const totalKg = monthlyKg.reduce((a, v) => a + v, 0);
    const avgClientes = monthlyAgg.length
      ? monthlyAgg.reduce((a, m) => a + m.clientesActivos, 0) / monthlyAgg.length
      : 0;

    const isEstimated = directFact === 0 && clients.length > 0;
    const filterDesc = [state.type !== 'all' ? state.type : null, state.search ? `"${state.search}"` : null].filter(Boolean).join(' · ');
    const kpis = [
      { label: isEstimated ? 'Facturación (est.)' : 'Facturación', value: fmtMoney.format(totalFact) },
      { label: matchedProducts.length > 0 ? 'Kg vendidos' : 'Kg vendidos (est.)', value: fmtAR.format(totalKg) },
      { label: 'Pedidos', value: fmtInt.format(totalPed) },
      { label: 'Clientes activos (prom.)', value: fmtInt.format(avgClientes) },
      { label: 'Unidades', value: fmtInt.format(totalUni) },
      { label: 'Filtro', value: filterDesc },
    ];

    // Comparación vs año anterior con mismos filtros
    const prevYear = String(Number(state.year) - 1);
    let prevMonthlyFact = null, prevMonthlyKg = null, prevMonthlyPed = null;
    if (state.data.years[prevYear]) {
      const prevY = state.data.years[prevYear];
      const prevTypeMap = new Map();
      (prevY.byClientType || []).forEach(c => prevTypeMap.set(c.cliente, c.tipo || 'Sin clasificar'));
      const prevAllClients = prevY.byClient || [];
      let prevClients = prevAllClients.slice();
      if (state.type !== 'all') prevClients = prevClients.filter(c => (prevTypeMap.get(c.cliente) || 'Sin clasificar') === state.type);
      if (state.search) prevClients = prevClients.filter(c => c.cliente.toLowerCase().includes(state.search));
      const prevMg = prevY.monthlyGeneral || [];
      const prevRatioByMonth = months.map(mi => {
        const filtUni = prevClients.reduce((s, c) => s + (c.byMonth[mi]?.unidades || 0), 0);
        const totUni = prevAllClients.reduce((s, c) => s + (c.byMonth[mi]?.unidades || 0), 0);
        return totUni > 0 ? filtUni / totUni : 0;
      });
      const prevDirectFact = months.reduce((a, i) => a + prevClients.reduce((s, c) => s + (c.byMonth[i]?.facturacion || 0), 0), 0);
      prevMonthlyFact = months.map((mi, idx) => {
        if (prevDirectFact > 0) {
          return prevClients.reduce((s, c) => s + (c.byMonth[mi]?.facturacion || 0), 0);
        }
        const mgRow = prevMg.find(r => r.monthIndex === mi);
        const mgFact = mgRow ? Number(mgRow['Facturación']) || 0 : 0;
        return prevRatioByMonth[idx] * mgFact;
      });
      prevMonthlyPed = months.map(i => prevClients.reduce((s, c) => s + (c.byMonth[i]?.pedidos || 0), 0));
      prevMonthlyKg = months.map((mi, idx) => {
        const mgRow = prevMg.find(r => r.monthIndex === mi);
        const mgKg = mgRow ? Number(mgRow['Kg Vendidos']) || 0 : 0;
        return prevRatioByMonth[idx] * mgKg;
      });
      const prevFact = prevMonthlyFact.reduce((a, v) => a + v, 0);
      const prevPed = prevMonthlyPed.reduce((a, v) => a + v, 0);
      const prevKg = prevMonthlyKg.reduce((a, v) => a + v, 0);
      const setDelta = (idx, curr, prev) => {
        if (prev > 0) {
          const delta = ((curr - prev) / prev) * 100;
          kpis[idx].delta = `${delta >= 0 ? '▲' : '▼'} ${fmtAR.format(Math.abs(delta))}% vs ${prevYear}`;
          kpis[idx].deltaClass = delta >= 0 ? 'up' : 'down';
        }
      };
      setDelta(0, totalFact, prevFact);
      setDelta(1, totalKg, prevKg);
      setDelta(2, totalPed, prevPed);
    }
    renderKpiGrid(kpis);

    const labels = months.map(i => MONTHS_SHORT[i]);
    const prevDs = (label, data, color) => prevMonthlyFact
      ? [{ label: `${label} ${prevYear}`, data, color, muted: true }]
      : [];
    drawBar('chartFacturacion', labels, [
      { label: `Facturación ${state.year}`, data: monthlyFact, color: '#7c8cff' },
      ...(prevMonthlyFact ? prevDs('Facturación', prevMonthlyFact, '#7c8cff') : []),
    ], { money: true });
    drawBar('chartKg', labels, [
      { label: `Kg ${state.year}`, data: monthlyKg, color: '#4bd6b0' },
      ...(prevMonthlyKg ? prevDs('Kg', prevMonthlyKg, '#4bd6b0') : []),
    ]);
    drawBar('chartPedidos', labels, [
      { label: `Pedidos ${state.year}`, data: monthlyAgg.map(m => m.pedidos), color: '#ffb457' },
      ...(prevMonthlyPed ? prevDs('Pedidos', prevMonthlyPed, '#ffb457') : []),
    ]);
    drawLine('chartClientes', labels, [
      { label: 'Activos', data: monthlyAgg.map(m => m.clientesActivos), color: '#7c8cff' },
    ]);
  }
}

function renderSearchMatches(y, months) {
  const el = $('#searchMatches');
  if (!state.search) { el.hidden = true; return; }

  const typeMap = new Map();
  (y.byClientType || []).forEach(c => typeMap.set(c.cliente, c.tipo || 'Sin clasificar'));

  // Clientes que coinciden
  const matchedClients = (y.byClient || [])
    .filter(c => c.cliente.toLowerCase().includes(state.search))
    .filter(c => state.type === 'all' || (typeMap.get(c.cliente) || 'Sin clasificar') === state.type)
    .map(c => {
      const uni = months.reduce((a, i) => a + (c.byMonth[i]?.unidades || 0), 0);
      const tipo = typeMap.get(c.cliente) || '';
      return { name: c.cliente, detail: tipo ? `${tipo} · ${fmtInt.format(uni)} uni` : `${fmtInt.format(uni)} uni` };
    })
    .sort((a, b) => a.name.localeCompare(b.name));

  // Productos que coinciden
  const matchedProducts = (y.products || [])
    .filter(p =>
      p.producto.toLowerCase().includes(state.search) ||
      (p.categoria || '').toLowerCase().includes(state.search)
    )
    .map(p => {
      const kg = months.reduce((a, i) => a + (p.byMonth[i]?.kg || 0), 0);
      const uni = months.reduce((a, i) => a + (p.byMonth[i]?.unidades || 0), 0);
      return { name: p.producto, detail: `${p.categoria} · ${fmtAR.format(kg)} kg · ${fmtInt.format(uni)} uni` };
    })
    .sort((a, b) => a.name.localeCompare(b.name));

  if (!matchedClients.length && !matchedProducts.length) {
    el.hidden = false;
    el.innerHTML = `<h4>Resultados para "${state.search}"</h4><p style="color:var(--muted);font-size:.85rem;margin:0;">Sin coincidencias</p>`;
    return;
  }

  let html = `<h4>Resultados para "${state.search}"</h4>`;
  if (matchedClients.length) {
    html += `<div class="match-group"><div class="match-label">Clientes (${matchedClients.length})</div><div class="match-tags">`;
    html += matchedClients.map(c => `<span class="match-tag">${c.name}<span class="tag-detail">${c.detail}</span></span>`).join('');
    html += '</div></div>';
  }
  if (matchedProducts.length) {
    html += `<div class="match-group"><div class="match-label">Productos (${matchedProducts.length})</div><div class="match-tags">`;
    html += matchedProducts.map(p => `<span class="match-tag">${p.name}<span class="tag-detail">${p.detail}</span></span>`).join('');
    html += '</div></div>';
  }
  el.hidden = false;
  el.innerHTML = html;
}

function renderKpiGrid(kpis) {
  $('#kpiGrid').innerHTML = kpis.map(k => `
    <div class="kpi">
      <div class="label">${k.label}</div>
      <div class="value">${k.value}</div>
      ${k.delta ? `<div class="delta ${k.deltaClass||''}">${k.delta}</div>` : ''}
    </div>`).join('');
}

// === Análisis y proyecciones ===
function renderAnalisis() {
  const container = $('#analisisSection');
  if (!container) return;

  const y = currentYearData();
  const prevYear = String(Number(state.year) - 1);
  const prevY = state.data.years[prevYear];
  const mg = y.monthlyGeneral || [];
  const mgPrev = prevY ? (prevY.monthlyGeneral || []) : [];
  const num = v => (v == null ? 0 : Number(v) || 0);
  const hasValue = r => r && r['Facturación'] != null && Number(r['Facturación']) > 0;

  const completed = mg.filter(hasValue).sort((a, b) => a.monthIndex - b.monthIndex);

  if (!completed.length) {
    container.innerHTML = `<div class="card"><h3>Análisis y proyecciones</h3>
      <p style="color:var(--muted);margin:0;">Sin meses registrados en el año ${state.year}.</p></div>`;
    return;
  }
  if (!prevY || !mgPrev.length) {
    container.innerHTML = `<div class="card"><h3>Análisis y proyecciones</h3>
      <p style="color:var(--muted);margin:0;">No hay datos de ${prevYear} para comparar ni proyectar.</p></div>`;
    return;
  }

  const metrics = [
    { key: 'Facturación', label: 'Facturación', fmt: fmtMoney, money: true },
    { key: 'Kg Vendidos', label: 'Kg vendidos', fmt: fmtAR, money: false },
    { key: 'Pedidos del mes', label: 'Pedidos', fmt: fmtInt, money: false },
    { key: 'Clientes activos', label: 'Clientes activos', fmt: fmtInt, money: false, avg: true },
  ];

  // Último mes completado
  const last = completed[completed.length - 1];
  const lastIdx = last.monthIndex;
  const prevSame = mgPrev.find(r => r.monthIndex === lastIdx);

  const monthCmpHTML = metrics.map(m => {
    const curr = num(last[m.key]);
    const prev = prevSame ? num(prevSame[m.key]) : 0;
    return rowHTML(m, curr, prev, state.year, prevYear);
  }).join('');

  // YTD: meses completados vs mismo período año anterior
  const completedIdxs = completed.map(r => r.monthIndex);
  const prevSamePeriod = mgPrev.filter(r => completedIdxs.includes(r.monthIndex));
  const ytdCurr = {}, ytdPrev = {};
  metrics.forEach(m => {
    const sumC = completed.reduce((a, r) => a + num(r[m.key]), 0);
    const sumP = prevSamePeriod.reduce((a, r) => a + num(r[m.key]), 0);
    ytdCurr[m.key] = m.avg ? sumC / (completed.length || 1) : sumC;
    ytdPrev[m.key] = m.avg ? sumP / (prevSamePeriod.length || 1) : sumP;
  });

  const ytdHTML = metrics.map(m => rowHTML(m, ytdCurr[m.key], ytdPrev[m.key], state.year, prevYear)).join('');

  // Tasas de crecimiento (acumulada y reciente)
  const growthYTD = {}, growthRecent = {}, growthCombined = {};
  metrics.forEach(m => {
    growthYTD[m.key] = ytdPrev[m.key] > 0 ? (ytdCurr[m.key] / ytdPrev[m.key]) - 1 : 0;
  });
  const recentCount = Math.min(3, completed.length);
  const recentCurr = completed.slice(-recentCount);
  const recentIdxs = recentCurr.map(r => r.monthIndex);
  const recentPrev = mgPrev.filter(r => recentIdxs.includes(r.monthIndex));
  metrics.forEach(m => {
    const cs = recentCurr.reduce((a, r) => a + num(r[m.key]), 0);
    const ps = recentPrev.reduce((a, r) => a + num(r[m.key]), 0);
    growthRecent[m.key] = ps > 0 ? (cs / ps) - 1 : growthYTD[m.key];
  });
  // Crecimiento combinado: 60% tendencia reciente + 40% acumulada anual
  metrics.forEach(m => {
    growthCombined[m.key] = growthRecent[m.key] * 0.6 + growthYTD[m.key] * 0.4;
  });

  // Proyección meses restantes: mes_año_anterior × (1 + crecimiento combinado)
  const projections = [];
  for (let i = 0; i < 12; i++) {
    if (completedIdxs.includes(i)) continue;
    const prevRow = mgPrev.find(r => r.monthIndex === i);
    if (!prevRow) continue;
    const proj = { monthIndex: i };
    metrics.forEach(m => {
      const basePrev = num(prevRow[m.key]);
      proj[m.key] = basePrev * (1 + growthCombined[m.key]);
    });
    projections.push(proj);
  }

  // Cierre de año: real + proyectado
  const yearEnd = {}, yearPrevTotal = {};
  metrics.forEach(m => {
    if (m.avg) {
      const reals = completed.map(r => num(r[m.key]));
      const projs = projections.map(p => p[m.key]);
      const combined = [...reals, ...projs];
      yearEnd[m.key] = combined.length ? combined.reduce((a, v) => a + v, 0) / combined.length : 0;
      const allPrev = mgPrev.map(r => num(r[m.key])).filter(v => v > 0);
      yearPrevTotal[m.key] = allPrev.length ? allPrev.reduce((a, v) => a + v, 0) / allPrev.length : 0;
    } else {
      yearEnd[m.key] = completed.reduce((a, r) => a + num(r[m.key]), 0) + projections.reduce((a, p) => a + p[m.key], 0);
      yearPrevTotal[m.key] = mgPrev.reduce((a, r) => a + num(r[m.key]), 0);
    }
  });
  const projHTML = metrics.map(m => rowHTML(m, yearEnd[m.key], yearPrevTotal[m.key], `${state.year} (proy.)`, prevYear)).join('');

  // Insights clave derivados de todas las fuentes
  const insights = buildInsights({
    y, prevY, completed, projections, mgPrev,
    ytdCurr, ytdPrev, yearEnd, yearPrevTotal,
    growthYTD, growthRecent, growthCombined,
    prevYear, lastIdx,
  });

  // Render HTML
  const methodNote = `Proyección basada en la estacionalidad del año anterior ajustada por el crecimiento de ${state.year} (60% tendencia últimos ${recentCount} ${recentCount === 1 ? 'mes' : 'meses'} + 40% acumulado anual).`;

  container.innerHTML = `
    <div class="analisis-header">
      <h2>Análisis y proyecciones</h2>
      <p>Basado en ${completed.length} ${completed.length === 1 ? 'mes registrado' : 'meses registrados'} de ${state.year} vs ${prevYear}</p>
    </div>
    <div class="analysis-grid">
      <div class="card analysis-card">
        <h3>Último mes registrado <span class="badge">${MONTHS_LONG[lastIdx]} ${state.year}</span></h3>
        <div class="cmp-list">${monthCmpHTML}</div>
      </div>
      <div class="card analysis-card">
        <h3>Acumulado del año <span class="badge">Ene – ${MONTHS_LONG[lastIdx]}</span></h3>
        <div class="cmp-list">${ytdHTML}</div>
      </div>
      <div class="card analysis-card">
        <h3>Proyección cierre <span class="badge">${state.year}</span></h3>
        <div class="cmp-list">${projHTML}</div>
        <p class="analysis-note">${methodNote}</p>
      </div>
    </div>
    <div class="card projection-wrap">
      <h3>Proyección mensual de facturación</h3>
      <canvas id="chartProyeccion"></canvas>
      <div class="legend-note">
        <span class="dot-real">Real ${state.year}</span>
        <span class="dot-proj">Proyectado ${state.year}</span>
        <span class="dot-prev">${prevYear}</span>
      </div>
    </div>
    <div class="card">
      <h3>Insights clave</h3>
      <ul class="insights">${insights.map(i => `<li>${i}</li>`).join('')}</ul>
    </div>
  `;

  drawProjection({
    labels: MONTHS_SHORT,
    completed, projections, mgPrev,
    prevYear,
  });
}

function rowHTML(m, curr, prev, currLabel, prevLabel) {
  const delta = prev > 0 ? ((curr - prev) / prev) * 100 : null;
  const deltaClass = delta == null ? '' : (delta >= 0 ? 'up' : 'down');
  const deltaTxt = delta == null ? 's/d' : `${delta >= 0 ? '▲' : '▼'} ${fmtAR.format(Math.abs(delta))}%`;
  return `
    <div class="cmp-row">
      <div class="cmp-label">${m.label}</div>
      <div class="cmp-vals">
        <div><span class="cmp-year">${currLabel}</span><strong>${m.fmt.format(curr)}</strong></div>
        <div><span class="cmp-year">${prevLabel}</span><span class="cmp-prev">${m.fmt.format(prev)}</span></div>
      </div>
      <div class="cmp-delta ${deltaClass}">${deltaTxt}</div>
    </div>`;
}

function buildInsights(ctx) {
  const {
    y, prevY, completed, projections, mgPrev,
    ytdCurr, ytdPrev, yearEnd, yearPrevTotal,
    growthYTD, growthRecent, growthCombined,
    prevYear, lastIdx,
  } = ctx;
  const num = v => (v == null ? 0 : Number(v) || 0);
  const out = [];

  const pct = x => `${x >= 0 ? '+' : ''}${fmtAR.format(x * 100)}%`;
  const dir = (x, pos = 'up', neg = 'down') => `<span class="${x >= 0 ? pos : neg}">${pct(x)}</span>`;

  // 1. Tendencia general facturación
  const gComb = growthCombined['Facturación'];
  const gYTD = growthYTD['Facturación'];
  const gRec = growthRecent['Facturación'];
  out.push(`Facturación acumulada ${dir(gYTD)} vs ${prevYear} (${fmtMoney.format(ytdCurr['Facturación'])} contra ${fmtMoney.format(ytdPrev['Facturación'])}). ` +
           `Tendencia reciente ${dir(gRec)}. Proyección de cierre de año: <strong>${fmtMoney.format(yearEnd['Facturación'])}</strong> (${dir((yearEnd['Facturación'] - yearPrevTotal['Facturación']) / (yearPrevTotal['Facturación'] || 1))} vs ${prevYear}).`);

  // 2. Volumen (kg) y pedidos
  out.push(`Volumen vendido: <strong>${fmtAR.format(ytdCurr['Kg Vendidos'])} kg</strong> vs ${fmtAR.format(ytdPrev['Kg Vendidos'])} kg en ${prevYear} (${dir(growthYTD['Kg Vendidos'])}). ` +
           `Pedidos: <strong>${fmtInt.format(ytdCurr['Pedidos del mes'])}</strong> vs ${fmtInt.format(ytdPrev['Pedidos del mes'])} (${dir(growthYTD['Pedidos del mes'])}).`);

  // 3. Ticket promedio (facturación / pedido)
  const ticketCurr = ytdCurr['Pedidos del mes'] > 0 ? ytdCurr['Facturación'] / ytdCurr['Pedidos del mes'] : 0;
  const ticketPrev = ytdPrev['Pedidos del mes'] > 0 ? ytdPrev['Facturación'] / ytdPrev['Pedidos del mes'] : 0;
  const ticketDelta = ticketPrev > 0 ? (ticketCurr / ticketPrev) - 1 : 0;
  const kgPorPed = ytdCurr['Pedidos del mes'] > 0 ? ytdCurr['Kg Vendidos'] / ytdCurr['Pedidos del mes'] : 0;
  const kgPorPedPrev = ytdPrev['Pedidos del mes'] > 0 ? ytdPrev['Kg Vendidos'] / ytdPrev['Pedidos del mes'] : 0;
  const kgPedDelta = kgPorPedPrev > 0 ? (kgPorPed / kgPorPedPrev) - 1 : 0;
  out.push(`Ticket promedio por pedido: <strong>${fmtMoney.format(ticketCurr)}</strong> (${dir(ticketDelta)} vs ${fmtMoney.format(ticketPrev)}). ` +
           `Tamaño promedio por pedido: <strong>${fmtAR.format(kgPorPed)} kg</strong> (${dir(kgPedDelta)}).`);

  // 4. Base de clientes: activos, nuevos, reactivados, pérdidas
  const nuevos = completed.reduce((a, r) => a + num(r['Nuevos clientes']), 0);
  const reactiv = completed.reduce((a, r) => a + num(r['Clientes reactivados']), 0);
  const perdidos = completed.reduce((a, r) => a + num(r['Clientes que no compran mas']), 0);
  const nuevosPrev = mgPrev.filter(r => completed.some(c => c.monthIndex === r.monthIndex))
    .reduce((a, r) => a + num(r['Nuevos clientes']), 0);
  const activosProm = ytdCurr['Clientes activos'];
  const activosPromPrev = ytdPrev['Clientes activos'];
  const saldoNeto = nuevos + reactiv - perdidos;
  out.push(`Base de clientes activos (promedio mensual): <strong>${fmtInt.format(activosProm)}</strong> (${dir(growthYTD['Clientes activos'])} vs ${fmtInt.format(activosPromPrev)}). ` +
           `En lo que va del año se incorporaron <strong>${nuevos}</strong> nuevos${nuevosPrev > 0 ? ` (${dir((nuevos - nuevosPrev) / nuevosPrev)} vs ${nuevosPrev} en ${prevYear})` : ''}, ` +
           `${reactiv} reactivados y ${perdidos} dados de baja. Saldo neto: <strong class="${saldoNeto >= 0 ? 'up' : 'down'}">${saldoNeto >= 0 ? '+' : ''}${saldoNeto}</strong>.`);

  // 5. Mejor y peor mes del año en curso
  let best = completed[0], worst = completed[0];
  completed.forEach(r => {
    if (num(r['Facturación']) > num(best['Facturación'])) best = r;
    if (num(r['Facturación']) < num(worst['Facturación'])) worst = r;
  });
  out.push(`Mejor mes hasta ahora: <strong>${MONTHS_LONG[best.monthIndex]}</strong> con ${fmtMoney.format(num(best['Facturación']))}. ` +
           `Mes más bajo: ${MONTHS_LONG[worst.monthIndex]} con ${fmtMoney.format(num(worst['Facturación']))}.`);

  // 6. Mes con mayor proyección y estacionalidad
  if (projections.length) {
    const bestProj = projections.reduce((b, p) => b && num(b['Facturación']) > p['Facturación'] ? b : p, null);
    const worstProj = projections.reduce((w, p) => w && num(w['Facturación']) < p['Facturación'] ? w : p, null);
    out.push(`Meses con mejor proyección: <strong>${MONTHS_LONG[bestProj.monthIndex]}</strong> (${fmtMoney.format(bestProj['Facturación'])}) y ` +
             `más bajo proyectado: ${MONTHS_LONG[worstProj.monthIndex]} (${fmtMoney.format(worstProj['Facturación'])}).`);
  }

  // 7. Top producto y categoría (de Excel de productos)
  const products = y.products || [];
  if (products.length) {
    const prodAgg = products.map(p => ({
      producto: p.producto,
      categoria: p.categoria || '—',
      kg: (p.byMonth || []).reduce((a, m) => a + num(m?.kg), 0),
      uni: (p.byMonth || []).reduce((a, m) => a + num(m?.unidades), 0),
    }));
    const totalKg = prodAgg.reduce((a, p) => a + p.kg, 0) || 1;
    const topProd = prodAgg.slice().sort((a, b) => b.kg - a.kg)[0];
    const catMap = new Map();
    prodAgg.forEach(p => catMap.set(p.categoria, (catMap.get(p.categoria) || 0) + p.kg));
    const topCat = [...catMap.entries()].sort((a, b) => b[1] - a[1])[0];
    if (topProd && topProd.kg > 0) {
      out.push(`Producto líder del año: <strong>${topProd.producto}</strong> con ${fmtAR.format(topProd.kg)} kg (${fmtAR.format((topProd.kg / totalKg) * 100)}% del total). ` +
               (topCat ? `Categoría dominante: <strong>${topCat[0]}</strong> con ${fmtAR.format(topCat[1])} kg (${fmtAR.format((topCat[1] / totalKg) * 100)}%).` : ''));
    }
  }

  // 8. Concentración de clientes (Top 10)
  const byClient = y.byClient || [];
  if (byClient.length) {
    const rank = byClient.map(c => ({
      cliente: c.cliente,
      uni: (c.byMonth || []).reduce((a, m) => a + num(m?.unidades), 0),
      fact: (c.byMonth || []).reduce((a, m) => a + num(m?.facturacion), 0),
    })).sort((a, b) => b.uni - a.uni);
    const top10 = rank.slice(0, 10);
    const top10Uni = top10.reduce((a, c) => a + c.uni, 0);
    const totalUni = rank.reduce((a, c) => a + c.uni, 0) || 1;
    const conc = (top10Uni / totalUni) * 100;
    out.push(`Concentración: los <strong>top 10 clientes</strong> representan el <strong>${fmtAR.format(conc)}%</strong> de las unidades del año. ` +
             `Cliente principal: <strong>${top10[0].cliente}</strong> (${fmtAR.format((top10[0].uni / totalUni) * 100)}% del total).`);
  }

  // 9. Distribución por tipo de cliente (si hay datos)
  const byType = y.byClientType || prevY.byClientType || [];
  if (byType.length) {
    const typeMap = new Map();
    byType.forEach(c => typeMap.set(c.cliente, c.tipo || 'Sin clasificar'));
    const typeAgg = new Map();
    byClient.forEach(c => {
      const t = typeMap.get(c.cliente) || 'Sin clasificar';
      const u = (c.byMonth || []).reduce((a, m) => a + num(m?.unidades), 0);
      typeAgg.set(t, (typeAgg.get(t) || 0) + u);
    });
    const totU = [...typeAgg.values()].reduce((a, v) => a + v, 0) || 1;
    const tipos = [...typeAgg.entries()].sort((a, b) => b[1] - a[1]).slice(0, 3);
    if (tipos.length) {
      out.push(`Mix por tipo de cliente: ` + tipos.map(t => `<strong>${t[0]}</strong> ${fmtAR.format((t[1] / totU) * 100)}%`).join(' · ') + '.');
    }
  }

  // 10. Evaluación final / recomendación
  let verdict;
  if (gComb >= 0.15) verdict = `El año avanza con <span class="up">crecimiento fuerte</span>; el ritmo actual proyecta superar cómodamente a ${prevYear}.`;
  else if (gComb >= 0.05) verdict = `El año avanza con <span class="up">crecimiento moderado</span> sobre ${prevYear}; conviene sostener las acciones que lo impulsan.`;
  else if (gComb >= -0.05) verdict = `El año avanza <strong>en línea</strong> con ${prevYear}; cierre estimado similar al del año anterior.`;
  else if (gComb >= -0.15) verdict = `El año muestra una <span class="down">leve caída</span> frente a ${prevYear}; revisar clientes con baja actividad y mix de productos.`;
  else verdict = `El año presenta una <span class="down">caída importante</span> vs ${prevYear}; se recomienda priorizar plan de recuperación comercial.`;
  out.push(verdict);

  return out;
}

function drawProjection({ labels, completed, projections, mgPrev, prevYear }) {
  destroyChart('chartProyeccion');
  const num = v => (v == null ? 0 : Number(v) || 0);
  const real = labels.map((_, i) => {
    const r = completed.find(c => c.monthIndex === i);
    return r ? num(r['Facturación']) : null;
  });
  const proj = labels.map((_, i) => {
    const p = projections.find(p => p.monthIndex === i);
    return p ? p['Facturación'] : null;
  });
  const prev = labels.map((_, i) => {
    const r = mgPrev.find(r => r.monthIndex === i);
    return r ? num(r['Facturación']) : 0;
  });

  const canvas = document.getElementById('chartProyeccion');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  charts['chartProyeccion'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: `Real ${state.year}`, data: real, backgroundColor: '#7c8cffcc', borderColor: '#7c8cff', borderWidth: 1, borderRadius: 6, order: 2 },
        { label: `Proyectado ${state.year}`, data: proj, backgroundColor: 'rgba(124,140,255,.25)', borderColor: '#7c8cff', borderWidth: 1.5, borderDash: [6, 4], borderRadius: 6, order: 2 },
        { label: `${prevYear}`, data: prev, type: 'line', borderColor: '#ffb457', backgroundColor: 'rgba(255,180,87,.08)', tension: .3, pointRadius: 3, borderWidth: 2, fill: false, order: 1 },
      ],
    },
    options: chartOpts({ money: true }),
  });
}

function renderMensual() {
  const y = currentYearData();
  const mg = y.monthlyGeneral || [];
  const months = monthFilterArray();
  const hasClientFilter = state.search || state.type !== 'all';

  if (!hasClientFilter) {
    const filtered = mg.filter(r => months.includes(r.monthIndex));
    if (!filtered.length) return tableHTML('#tableMensual', [], []);
    const keys = Object.keys(filtered[0]).filter(k => k !== 'monthIndex');
    const rows = filtered.map(r => keys.map(k => r[k]));
    const cols = keys.map(k => ({ label: k, num: k !== 'mes' }));
    tableHTML('#tableMensual', cols, rows);
  } else {
    // Recalcular desde datos detallados por cliente
    const typeMap = new Map();
    (y.byClientType || []).forEach(c => typeMap.set(c.cliente, c.tipo || 'Sin clasificar'));
    const allClients = y.byClient || [];
    let clients = allClients.slice();
    if (state.type !== 'all') clients = clients.filter(c => (typeMap.get(c.cliente) || 'Sin clasificar') === state.type);
    if (state.search) clients = clients.filter(c => c.cliente.toLowerCase().includes(state.search));

    // Totales globales para proporciones
    const globalUniByMonth = months.map(i => allClients.reduce((a, c) => a + (c.byMonth[i]?.unidades || 0), 0));
    const mgByMonth = new Map(mg.map(r => [r.monthIndex, r]));

    const rows = months.map((i, idx) => {
      const uni = clients.reduce((a, c) => a + (c.byMonth[i]?.unidades || 0), 0);
      const ped = clients.reduce((a, c) => a + (c.byMonth[i]?.pedidos || 0), 0);
      const directFact = clients.reduce((a, c) => a + (c.byMonth[i]?.facturacion || 0), 0);
      const ratio = globalUniByMonth[idx] > 0 ? uni / globalUniByMonth[idx] : 0;
      const mgRow = mgByMonth.get(i);
      const fact = directFact > 0 ? directFact : (mgRow ? ratio * (Number(mgRow['Facturación']) || 0) : 0);
      const kg = mgRow ? ratio * (Number(mgRow['Kg Vendidos']) || 0) : 0;
      const clientesActivos = clients.filter(c => {
        const m = c.byMonth[i];
        return m && ((m.unidades || 0) > 0 || (m.pedidos || 0) > 0 || (m.facturacion || 0) > 0);
      }).length;
      return [MONTHS_LONG[i], kg, uni, ped, clientesActivos, fact];
    });
    const cols = [
      { label: 'Mes', num: false },
      { label: 'Kg Vendidos (est.)', num: true },
      { label: 'Unidades', num: true },
      { label: 'Pedidos', num: true },
      { label: 'Clientes activos', num: true },
      { label: directFactAvailable(clients, months) ? 'Facturación' : 'Facturación (est.)', num: true, money: true },
    ];
    tableHTML('#tableMensual', cols, rows);
  }
}

function directFactAvailable(clients, months) {
  return clients.some(c => months.some(i => (c.byMonth[i]?.facturacion || 0) > 0));
}

function renderClientes() {
  const y = currentYearData();
  const src = y.byClient || [];
  const mg = y.monthlyGeneral || [];
  const typeMap = new Map();
  (y.byClientType || []).forEach(c => typeMap.set(c.cliente, c.tipo));

  const months = monthFilterArray();

  // Verificar si hay facturación directa en byClient
  const hasDirectFact = src.some(c => months.some(i => (c.byMonth[i]?.facturacion || 0) > 0));

  // Totales globales para proporción si no hay facturación directa
  const globalUniByMonth = months.map(i => src.reduce((a, c) => a + (c.byMonth[i]?.unidades || 0), 0));
  const mgByMonth = new Map(mg.map(r => [r.monthIndex, r]));

  let list = src.map(c => {
    const u = months.reduce((a, i) => a + (c.byMonth[i]?.unidades || 0), 0);
    const p = months.reduce((a, i) => a + (c.byMonth[i]?.pedidos || 0), 0);
    let f = months.reduce((a, i) => a + (c.byMonth[i]?.facturacion || 0), 0);
    if (f === 0 && u > 0 && !hasDirectFact) {
      // Estimar facturación proporcionalmente
      f = months.reduce((a, i, idx) => {
        const cUni = c.byMonth[i]?.unidades || 0;
        const gUni = globalUniByMonth[idx];
        const mgRow = mgByMonth.get(i);
        const mgFact = mgRow ? Number(mgRow['Facturación']) || 0 : 0;
        return a + (gUni > 0 ? (cUni / gUni) * mgFact : 0);
      }, 0);
    }
    return { cliente: c.cliente, tipo: typeMap.get(c.cliente) || '—', unidades: u, pedidos: p, facturacion: f };
  });

  if (state.search) list = list.filter(c => c.cliente.toLowerCase().includes(state.search));
  if (state.type !== 'all') list = list.filter(c => c.tipo === state.type);

  // Orden
  const sort = state.sortBy.clientes;
  list.sort((a, b) => (b[sort.key === 'totalFacturacion' ? 'facturacion' : sort.key] ?? 0) - (a[sort.key === 'totalFacturacion' ? 'facturacion' : sort.key] ?? 0));
  if (sort.dir === 'asc') list.reverse();

  // Top N chart
  const metricKey = state.clientMetric;
  const metricLabels = { facturacion: 'Facturación', unidades: 'Unidades', pedidos: 'Pedidos' };
  const top = [...list].sort((a, b) => b[metricKey] - a[metricKey]).slice(0, state.topN);
  drawHBar('chartTopClients', top.map(c => c.cliente), [{
    label: metricLabels[metricKey] || metricKey, data: top.map(c => c[metricKey]),
    color: metricKey === 'facturacion' ? '#7c8cff' : metricKey === 'unidades' ? '#4bd6b0' : '#ffb457'
  }], { money: metricKey === 'facturacion' });

  const factLabel = hasDirectFact ? 'Facturación' : 'Facturación (est.)';
  const cols = [
    { label: 'Cliente', num: false, key: 'cliente' },
    { label: 'Tipo', num: false, key: 'tipo' },
    { label: 'Unidades', num: true, key: 'unidades' },
    { label: 'Pedidos', num: true, key: 'pedidos' },
    { label: factLabel, num: true, key: 'facturacion', money: true },
  ];
  const rows = list.map(c => [c.cliente, c.tipo, c.unidades, c.pedidos, c.facturacion]);
  tableHTML('#tableClientes', cols, rows);
}

function renderTipos() {
  const y = currentYearData();
  const months = monthFilterArray();
  const mg = y.monthlyGeneral || [];

  // Construir mapa de tipos: usar byClientType del año actual, o buscar en otros años
  const typeMap = new Map();
  if ((y.byClientType || []).length > 0) {
    (y.byClientType || []).forEach(c => typeMap.set(c.cliente, c.tipo || 'Sin clasificar'));
  } else {
    // Buscar tipos en otros años como referencia
    for (const yr of Object.values(state.data.years)) {
      (yr.byClientType || []).forEach(c => {
        if (!typeMap.has(c.cliente)) typeMap.set(c.cliente, c.tipo || 'Sin clasificar');
      });
    }
  }

  // Usar byClient como fuente de datos (tiene todos los clientes con métricas)
  let clients = (y.byClient || []).map(c => ({
    cliente: c.cliente,
    tipo: typeMap.get(c.cliente) || 'Sin clasificar',
    byMonth: c.byMonth,
  }));

  if (state.search) clients = clients.filter(c => c.cliente.toLowerCase().includes(state.search));

  // Agregar por tipo
  const agg = new Map();
  clients.forEach(c => {
    const key = c.tipo;
    const curr = agg.get(key) || { tipo: key, clientes: 0, unidades: 0, pedidos: 0, facturacion: 0 };
    curr.clientes += 1;
    months.forEach(i => {
      curr.unidades += c.byMonth[i]?.unidades || 0;
      curr.pedidos += c.byMonth[i]?.pedidos || 0;
      curr.facturacion += c.byMonth[i]?.facturacion || 0;
    });
    agg.set(key, curr);
  });
  let list = [...agg.values()].sort((a, b) => b.unidades - a.unidades);
  if (state.type !== 'all') list = list.filter(t => t.tipo === state.type);

  // Estimar facturación si no hay datos directos
  const hasDirectFact = list.some(t => t.facturacion > 0);
  if (!hasDirectFact) {
    const totalUni = list.reduce((a, t) => a + t.unidades, 0);
    const totalFact = mg.filter(r => months.includes(r.monthIndex)).reduce((a, r) => a + (Number(r['Facturación']) || 0), 0);
    list.forEach(t => {
      if (totalUni > 0) t.facturacion = (t.unidades / totalUni) * totalFact;
    });
  }

  drawDonut('chartTipos', list.map(t => t.tipo), list.map(t => t.unidades));

  const factLabel = hasDirectFact ? 'Facturación' : 'Facturación (est.)';
  const cols = [
    { label: 'Tipo', num: false },
    { label: 'Clientes', num: true },
    { label: 'Unidades', num: true },
    { label: 'Pedidos', num: true },
    { label: factLabel, num: true, money: true },
  ];
  const rows = list.map(t => [t.tipo, t.clientes, t.unidades, t.pedidos, t.facturacion]);
  tableHTML('#tableTipos', cols, rows);
}

function renderProductos() {
  const y = currentYearData();
  let src = y.products || [];
  const months = monthFilterArray();

  // Filtro de categoría
  if (state.category !== 'all') src = src.filter(p => (p.categoria || '—') === state.category);

  // Filtro de búsqueda: aplicar solo si matchea productos o categorías.
  // Si el texto busca un cliente (no matchea productos), no vaciar la lista.
  if (state.search) {
    const matched = src.filter(p =>
      p.producto.toLowerCase().includes(state.search) ||
      (p.categoria || '').toLowerCase().includes(state.search)
    );
    if (matched.length > 0) src = matched;
  }

  const list = src.map(p => {
    const total = months.reduce((a, i) => a + (p.byMonth[i]?.[state.productMetric] || 0), 0);
    const unidades = months.reduce((a, i) => a + (p.byMonth[i]?.unidades || 0), 0);
    const kg = months.reduce((a, i) => a + (p.byMonth[i]?.kg || 0), 0);
    return { categoria: p.categoria || '—', producto: p.producto, total, unidades, kg };
  }).sort((a, b) => b.total - a.total);

  // Actualizar dropdown de categorías según productos visibles
  const visibleCats = new Set(list.map(p => p.categoria));
  const cSel = $('#categorySelect');
  const allCats = new Set((y.products || []).map(p => p.categoria || '—'));
  cSel.innerHTML = '<option value="all">Todas</option>' +
    [...allCats].sort().map(c => {
      const disabled = !visibleCats.has(c) && state.category === 'all';
      return `<option value="${c}" ${disabled ? 'class="dimmed"' : ''}>${c}${disabled ? ' (sin resultados)' : ''}</option>`;
    }).join('');
  cSel.value = state.category;

  const prodMetricLabels = { kg: 'Kg', unidades: 'Unidades' };
  drawHBar('chartProductos', list.slice(0, 20).map(p => p.producto),
    [{ label: prodMetricLabels[state.productMetric] || state.productMetric, data: list.slice(0, 20).map(p => p.total), color: '#7c8cff' }]);

  const cols = [
    { label: 'Categoría', num: false },
    { label: 'Producto', num: false },
    { label: 'Unidades', num: true },
    { label: 'Kg', num: true },
  ];
  const rows = list.map(p => [p.categoria, p.producto, p.unidades, p.kg]);
  tableHTML('#tableProductos', cols, rows);
}

// === Helpers render ===
function tableHTML(sel, cols, rows) {
  const el = $(sel); if (!el) return;
  if (!cols.length) { el.innerHTML = '<tbody><tr><td>Sin datos</td></tr></tbody>'; return; }
  const thead = `<thead><tr>${cols.map(c => `<th class="${c.num ? 'num' : ''}">${c.label}</th>`).join('')}</tr></thead>`;
  const tbody = `<tbody>${rows.map(r => `<tr>${r.map((v, i) => {
    const c = cols[i];
    let txt;
    if (v === null || v === undefined || v === '') txt = '—';
    else if (c.num) txt = c.money ? fmtMoney.format(Number(v)||0) : fmtAR.format(Number(v)||0);
    else txt = String(v);
    return `<td class="${c.num ? 'num' : ''}">${txt}</td>`;
  }).join('')}</tr>`).join('')}</tbody>`;
  el.innerHTML = thead + tbody;
}

function destroyChart(id) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }

function chartOpts(extra = {}) {
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { labels: { color: '#e8ecf7' } },
      tooltip: {
        callbacks: {
          label: ctx => {
            const v = ctx.parsed.y ?? ctx.parsed.x ?? ctx.parsed;
            return `${ctx.dataset.label}: ${extra.money ? fmtMoney.format(v) : fmtAR.format(v)}`;
          }
        }
      }
    },
    scales: {
      x: { ticks: { color: '#9aa3c4' }, grid: { color: 'rgba(154,163,196,.08)' } },
      y: { ticks: { color: '#9aa3c4', callback: v => extra.money ? fmtMoney.format(v) : fmtAR.format(v) }, grid: { color: 'rgba(154,163,196,.08)' } }
    },
    ...extra.overrides,
  };
}

function drawBar(id, labels, datasets, extra = {}) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: datasets.map(d => ({
      label: d.label,
      data: d.data,
      backgroundColor: d.color + (d.muted ? '33' : 'cc'),
      borderColor: d.color + (d.muted ? '66' : ''),
      borderWidth: d.muted ? 0 : 1,
      borderRadius: 6,
      order: d.muted ? 2 : 1,
    })) },
    options: chartOpts(extra),
  });
}
function drawLine(id, labels, datasets, extra = {}) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets: datasets.map(d => ({ label: d.label, data: d.data, borderColor: d.color, backgroundColor: d.color + '33', fill: true, tension: .3, pointRadius: 3 })) },
    options: chartOpts(extra),
  });
}
function drawHBar(id, labels, datasets, extra = {}) {
  destroyChart(id);
  const canvas = document.getElementById(id);
  canvas.classList.add('hbar-canvas');
  const ctx = canvas.getContext('2d');
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: datasets.map(d => ({ label: d.label, data: d.data, backgroundColor: d.color + 'cc', borderColor: d.color, borderWidth: 1, borderRadius: 6 })) },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: {
        legend: { labels: { color: '#e8ecf7' } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = ctx.parsed.x ?? ctx.parsed;
              return `${ctx.dataset.label}: ${extra.money ? fmtMoney.format(v) : fmtAR.format(v)}`;
            }
          }
        }
      },
      scales: {
        x: {
          ticks: { color: '#9aa3c4', callback: v => extra.money ? fmtMoney.format(v) : fmtAR.format(v) },
          grid: { color: 'rgba(154,163,196,.08)' },
        },
        y: {
          ticks: { color: '#e8ecf7', font: { size: 12 }, autoSkip: false },
          grid: { display: false },
          afterFit: scale => { scale.width = Math.max(scale.width, 160); },
        },
      },
    },
  });
  // Ajuste alto dinámico para barras horizontales
  const h = Math.max(300, labels.length * 34);
  const wrap = canvas.closest('.hbar-wrap') || canvas.parentElement;
  wrap.style.height = h + 'px';
  canvas.style.height = '100%';
  charts[id].resize();
}
function drawDonut(id, labels, data) {
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  const palette = ['#7c8cff', '#4bd6b0', '#ffb457', '#ff6b8b', '#b488ff', '#57c7ff', '#f7d94c', '#63e2c6'];
  charts[id] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: labels.map((_, i) => palette[i % palette.length]), borderColor: '#171a2e', borderWidth: 2 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { color: '#e8ecf7' } },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${fmtAR.format(ctx.parsed)}` } },
      }
    },
  });
}

loadData();
