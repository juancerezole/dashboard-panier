const MONTHS_SHORT = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const MONTHS_LONG = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const fmtAR = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 2 });
const fmtMoney = new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', maximumFractionDigits: 0 });
const fmtInt = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });

const state = {
  data: null,
  year: null,
  month: 'all',
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
  render();

  $('#loader').hidden = true;
  $('#sourceInfo').textContent = `Fuentes: ${state.data.sources.map(s => s.file).join(' · ')}`;
}

function wireEvents() {
  $$('#tabs button').forEach(b => b.addEventListener('click', () => {
    state.tab = b.dataset.tab;
    $$('#tabs button').forEach(x => x.classList.toggle('active', x === b));
    $$('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `tab-${state.tab}`));
    render();
  }));

  $('#yearSelect').addEventListener('change', e => { state.year = e.target.value; refreshFilters(); render(); });
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
  if (state.month === 'all') return Array.from({length:12}, (_,i)=>i);
  return [Number(state.month)];
}

// === Render ===
function render() {
  const y = currentYearData();
  const months = monthFilterArray();
  renderSearchMatches(y, months);
  renderResumen();
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
      const prevFact = prevFiltered.reduce((a, r) => a + (Number(r['Facturación']) || 0), 0);
      if (prevFact > 0) {
        const delta = ((sum('Facturación') - prevFact) / prevFact) * 100;
        kpis[0].delta = `${delta >= 0 ? '▲' : '▼'} ${fmtAR.format(Math.abs(delta))}% vs ${prevYear}`;
        kpis[0].deltaClass = delta >= 0 ? 'up' : 'down';
      }
    }
    renderKpiGrid(kpis);

    const labels = months.map(i => MONTHS_SHORT[i]);
    const pick = (key) => months.map(i => {
      const row = mg.find(r => r.monthIndex === i);
      return row ? Number(row[key]) || 0 : 0;
    });
    drawBar('chartFacturacion', labels, [{ label: 'Facturación', data: pick('Facturación'), color: '#7c8cff' }], { money: true });
    drawBar('chartKg', labels, [{ label: 'Kg', data: pick('Kg Vendidos'), color: '#4bd6b0' }]);
    drawBar('chartPedidos', labels, [{ label: 'Pedidos', data: pick('Pedidos del mes'), color: '#ffb457' }]);
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
    if (state.data.years[prevYear]) {
      const prevY = state.data.years[prevYear];
      const prevTypeMap = new Map();
      (prevY.byClientType || []).forEach(c => prevTypeMap.set(c.cliente, c.tipo || 'Sin clasificar'));
      const prevAllClients = prevY.byClient || [];
      let prevClients = prevAllClients.slice();
      if (state.type !== 'all') prevClients = prevClients.filter(c => (prevTypeMap.get(c.cliente) || 'Sin clasificar') === state.type);
      if (state.search) prevClients = prevClients.filter(c => c.cliente.toLowerCase().includes(state.search));
      let prevFact = months.reduce((a, i) => a + prevClients.reduce((s, c) => s + (c.byMonth[i]?.facturacion || 0), 0), 0);
      // Si no hay facturación directa en año anterior, estimar también
      if (prevFact === 0 && prevClients.length > 0) {
        const prevMg = prevY.monthlyGeneral || [];
        prevFact = months.reduce((a, mi) => {
          const filtUni = prevClients.reduce((s, c) => s + (c.byMonth[mi]?.unidades || 0), 0);
          const totUni = prevAllClients.reduce((s, c) => s + (c.byMonth[mi]?.unidades || 0), 0);
          const mgRow = prevMg.find(r => r.monthIndex === mi);
          const mgFact = mgRow ? Number(mgRow['Facturación']) || 0 : 0;
          return a + (totUni > 0 ? (filtUni / totUni) * mgFact : 0);
        }, 0);
      }
      if (prevFact > 0) {
        const delta = ((totalFact - prevFact) / prevFact) * 100;
        kpis[0].delta = `${delta >= 0 ? '▲' : '▼'} ${fmtAR.format(Math.abs(delta))}% vs ${prevYear}`;
        kpis[0].deltaClass = delta >= 0 ? 'up' : 'down';
      }
    }
    renderKpiGrid(kpis);

    const labels = months.map(i => MONTHS_SHORT[i]);
    drawBar('chartFacturacion', labels, [{ label: 'Facturación', data: monthlyFact, color: '#7c8cff' }], { money: true });
    drawBar('chartKg', labels, [{ label: 'Kg', data: monthlyKg, color: '#4bd6b0' }]);
    drawBar('chartPedidos', labels, [{ label: 'Pedidos', data: monthlyAgg.map(m => m.pedidos), color: '#ffb457' }]);
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
  if (state.category !== 'all') src = src.filter(p => (p.categoria || '—') === state.category);
  if (state.search) src = src.filter(p => p.producto.toLowerCase().includes(state.search) || (p.categoria||'').toLowerCase().includes(state.search));

  const list = src.map(p => {
    const total = months.reduce((a, i) => a + (p.byMonth[i]?.[state.productMetric] || 0), 0);
    const unidades = months.reduce((a, i) => a + (p.byMonth[i]?.unidades || 0), 0);
    const kg = months.reduce((a, i) => a + (p.byMonth[i]?.kg || 0), 0);
    return { categoria: p.categoria || '—', producto: p.producto, total, unidades, kg };
  }).sort((a, b) => b.total - a.total);

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
    data: { labels, datasets: datasets.map(d => ({ label: d.label, data: d.data, backgroundColor: d.color + 'cc', borderColor: d.color, borderWidth: 1, borderRadius: 6 })) },
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
