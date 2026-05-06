// Lee los Google Sheets de una carpeta de Drive y genera dashboard/data.json
// con la misma estructura normalizada por año que usaba la versión local.
//
// Requiere:
//   - GOOGLE_DRIVE_FOLDER_ID       (env)  ID de la carpeta en Drive
//   - GOOGLE_SERVICE_ACCOUNT_KEY_FILE (env, opcional) ruta al JSON del service account
//     Por defecto: ./service-account.json
//
// La carpeta debe estar compartida con el email del service account (rol "Lector").

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

try { require('dotenv').config(); } catch (_) { /* dotenv es opcional */ }

const ROOT = __dirname;
const OUT_DIR = path.join(ROOT, 'dashboard');
const OUT_FILE = path.join(OUT_DIR, 'data.json');
const KEY_FILE = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_FILE
  || path.join(ROOT, 'service-account.json');
const FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;

const MONTHS = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

const MONTH_ALIASES = {
  'ene': 0, 'enero': 0,
  'feb': 1, 'febrero': 1,
  'mar': 2, 'marzo': 2,
  'abr': 3, 'abril': 3,
  'may': 4, 'mayo': 4,
  'jun': 5, 'junio': 5,
  'jul': 6, 'julio': 6,
  'ago': 7, 'agosto': 7,
  'sep': 8, 'set': 8, 'septiembre': 8,
  'oct': 9, 'octubre': 9,
  'nov': 10, 'noviembre': 10,
  'dic': 11, 'diciembre': 11,
};

const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
const lc = (s) => norm(s).toLowerCase();

function detectMonthInHeader(header) {
  const parts = lc(header).split(/[\s\n/]+/).filter(Boolean);
  for (const p of parts) {
    const idx = detectMonthIndex(p);
    if (idx >= 0) return idx;
  }
  return -1;
}

function extractMetric(header) {
  const parts = lc(header).split(/[\s\n/]+/).filter(Boolean);
  const metrics = parts.filter(p => detectMonthIndex(p) < 0);
  const joined = metrics.join(' ');
  if (/\bkg\b/.test(joined)) return 'kg';
  if (/\buni/.test(joined)) return 'uni';
  if (/\bped/.test(joined)) return 'ped';
  if (/\bfact/.test(joined)) return 'fact';
  return null;
}

function detectMonthIndex(headerPart) {
  const token = lc(headerPart).replace(/[^a-záéíóúñ]/gi, '');
  for (const k of Object.keys(MONTH_ALIASES)) {
    if (token.startsWith(k)) return MONTH_ALIASES[k];
  }
  return -1;
}

// Normaliza la matriz devuelta por la API de Sheets (filas de largo variable,
// celdas vacías al final omitidas) al mismo shape que producía xlsx: filas de
// igual largo con null en los huecos, y sin filas totalmente vacías.
function normalizeRows(values) {
  if (!values || !values.length) return [];
  const width = values.reduce((m, r) => Math.max(m, r.length), 0);
  return values
    .map(r => {
      const row = r.slice();
      while (row.length < width) row.push(null);
      return row.map(c => (c === '' || c === undefined ? null : c));
    })
    .filter(r => r.some(c => c !== null));
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const r = rows[i];
    const strings = r.filter(c => typeof c === 'string' && norm(c));
    if (strings.length >= 2) return i;
  }
  return 0;
}

function parseMonthlyGeneral(rows) {
  const hi = findHeaderRow(rows);
  const headers = rows[hi].map(norm);
  const out = [];
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i];
    const mes = norm(r[0]);
    if (!mes) continue;
    const monthIdx = detectMonthIndex(mes);
    if (monthIdx < 0) continue;
    const rec = { mes: MONTHS[monthIdx], monthIndex: monthIdx };
    headers.forEach((h, j) => {
      if (j === 0) return;
      const key = h.replace(/\s+/g, ' ');
      rec[key] = r[j];
    });
    out.push(rec);
  }
  out.sort((a, b) => a.monthIndex - b.monthIndex);
  return out;
}

function parseByClient(rows) {
  const hi = findHeaderRow(rows);
  const headers = rows[hi].map(norm);
  const clientes = [];
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i];
    const cliente = norm(r[0]);
    if (!cliente) continue;
    const byMonth = Array.from({ length: 12 }, () => ({ unidades: 0, pedidos: 0, facturacion: 0, hasData: false }));
    for (let j = 1; j < headers.length; j++) {
      const h = headers[j];
      if (!h) continue;
      const monthIdx = detectMonthInHeader(h);
      if (monthIdx < 0) continue;
      const val = Number(r[j]);
      if (!isFinite(val)) continue;
      const metric = extractMetric(h);
      if (metric === 'uni') { byMonth[monthIdx].unidades = val; byMonth[monthIdx].hasData = true; }
      else if (metric === 'ped') { byMonth[monthIdx].pedidos = val; byMonth[monthIdx].hasData = true; }
      else if (metric === 'fact') { byMonth[monthIdx].facturacion = val; byMonth[monthIdx].hasData = true; }
    }
    const totalUni = byMonth.reduce((a, m) => a + (m.unidades || 0), 0);
    const totalPed = byMonth.reduce((a, m) => a + (m.pedidos || 0), 0);
    const totalFact = byMonth.reduce((a, m) => a + (m.facturacion || 0), 0);
    clientes.push({ cliente, byMonth, totalUnidades: totalUni, totalPedidos: totalPed, totalFacturacion: totalFact });
  }
  return clientes;
}

function parseByClientType(rows) {
  const hi = findHeaderRow(rows);
  const headers = rows[hi].map(norm);
  const tipoCol = headers.findIndex(h => lc(h) === 'tipo');
  const clientCol = 0;
  const out = [];
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i];
    const cliente = norm(r[clientCol]);
    if (!cliente) continue;
    const tipo = tipoCol >= 0 ? norm(r[tipoCol]) : 'Sin clasificar';
    const byMonth = Array.from({ length: 12 }, () => ({ unidades: 0, pedidos: 0, facturacion: 0, hasData: false }));
    for (let j = 0; j < headers.length; j++) {
      if (j === clientCol || j === tipoCol) continue;
      const h = headers[j];
      const monthIdx = detectMonthInHeader(h);
      if (monthIdx < 0) continue;
      const val = Number(r[j]);
      if (!isFinite(val)) continue;
      const metric = extractMetric(h);
      if (metric === 'uni') { byMonth[monthIdx].unidades = val; byMonth[monthIdx].hasData = true; }
      else if (metric === 'ped') { byMonth[monthIdx].pedidos = val; byMonth[monthIdx].hasData = true; }
      else if (metric === 'fact') { byMonth[monthIdx].facturacion = val; byMonth[monthIdx].hasData = true; }
    }
    out.push({ cliente, tipo: tipo || 'Sin clasificar', byMonth });
  }
  return out;
}

function parseProducts(rows) {
  const hi = findHeaderRow(rows);
  const headers = rows[hi].map(norm);
  const out = [];
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i];
    const categoria = norm(r[0]);
    const producto = norm(r[1]);
    if (!categoria && !producto) continue;
    if (!producto) continue;
    const byMonth = Array.from({ length: 12 }, () => ({ unidades: 0, kg: 0, hasData: false }));
    for (let j = 2; j < headers.length; j++) {
      const h = headers[j];
      const monthIdx = detectMonthInHeader(h);
      if (monthIdx < 0) continue;
      const val = Number(r[j]);
      if (!isFinite(val)) continue;
      const metric = extractMetric(h);
      if (metric === 'uni') { byMonth[monthIdx].unidades = val; byMonth[monthIdx].hasData = true; }
      else if (metric === 'kg') { byMonth[monthIdx].kg = val; byMonth[monthIdx].hasData = true; }
    }
    out.push({ categoria: categoria || '—', producto, byMonth });
  }
  return out;
}

function parseSummary(rows) {
  const out = [];
  for (const r of rows) {
    const k = norm(r[0]);
    const v = r[1];
    if (!k) continue;
    if (v === null || v === undefined || v === '') continue;
    out.push({ indicador: k, valor: v });
  }
  return out;
}

function classifySheet(name) {
  const n = lc(name);
  if (n.includes('mensual general')) return 'monthlyGeneral';
  if (n.includes('tipo de cliente')) return 'byClientType';
  if (n.includes('x cliente') || n.includes('por cliente')) return 'byClient';
  if (n.includes('unidades') && n.includes('kg')) return 'products';
  if (n.includes('resumen')) return 'summary';
  return null;
}

function parseSpreadsheet(sheetsData) {
  const result = {};
  for (const [sheetName, values] of Object.entries(sheetsData)) {
    const kind = classifySheet(sheetName);
    if (!kind) continue;
    const rows = normalizeRows(values);
    if (!rows.length) continue;
    try {
      if (kind === 'monthlyGeneral') result.monthlyGeneral = parseMonthlyGeneral(rows);
      else if (kind === 'byClient') result.byClient = parseByClient(rows);
      else if (kind === 'byClientType') result.byClientType = parseByClientType(rows);
      else if (kind === 'products') result.products = parseProducts(rows);
      else if (kind === 'summary') result.summary = parseSummary(rows);
    } catch (e) {
      console.warn(`! Error en hoja "${sheetName}":`, e.message);
    }
  }
  return result;
}

function getAuth() {
  if (!fs.existsSync(KEY_FILE)) {
    throw new Error(
      `No se encontró el archivo de credenciales del service account en: ${KEY_FILE}\n` +
      `Configurá GOOGLE_SERVICE_ACCOUNT_KEY_FILE o copiá el JSON a ./service-account.json`
    );
  }
  return new google.auth.GoogleAuth({
    keyFile: KEY_FILE,
    scopes: [
      'https://www.googleapis.com/auth/drive.readonly',
      'https://www.googleapis.com/auth/spreadsheets.readonly',
    ],
  });
}

async function listSheetsInFolder(auth) {
  const drive = google.drive({ version: 'v3', auth });
  const files = [];
  let pageToken;
  do {
    const res = await drive.files.list({
      q: `'${FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
      fields: 'nextPageToken, files(id, name, modifiedTime)',
      pageSize: 100,
      pageToken,
      // Requerido si la carpeta está en Shared Drive:
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    files.push(...(res.data.files || []));
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return files;
}

async function fetchSpreadsheet(auth, fileId) {
  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: fileId,
    fields: 'sheets.properties.title',
  });
  const sheetNames = (meta.data.sheets || []).map(s => s.properties.title);
  if (!sheetNames.length) return {};
  const batch = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: fileId,
    ranges: sheetNames,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING',
  });
  const out = {};
  (batch.data.valueRanges || []).forEach((vr, i) => {
    out[sheetNames[i]] = vr.values || [];
  });
  return out;
}

async function main() {
  if (!FOLDER_ID) {
    console.error('Falta GOOGLE_DRIVE_FOLDER_ID. Definilo en .env o como variable de entorno.');
    process.exit(1);
  }

  const auth = getAuth();
  console.log(`→ Buscando Google Sheets en carpeta ${FOLDER_ID}...`);
  const files = await listSheetsInFolder(auth);
  if (!files.length) {
    console.error('No se encontraron Google Sheets en esa carpeta. ¿La compartiste con el service account?');
    process.exit(1);
  }

  const years = {};
  const sources = [];
  for (const file of files) {
    const m = file.name.match(/(20\d{2})/);
    const year = m ? m[1] : 'Sin año';
    const sheetsData = await fetchSpreadsheet(auth, file.id);
    const parsed = parseSpreadsheet(sheetsData);
    years[year] = parsed;
    sources.push({ year, file: file.name, fileId: file.id, modifiedTime: file.modifiedTime });
    console.log(`✓ ${file.name} → año ${year} (${Object.keys(parsed).join(', ') || 'sin datos reconocidos'})`);
  }

  if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });
  const out = { generatedAt: new Date().toISOString(), sources, years };
  fs.writeFileSync(OUT_FILE, JSON.stringify(out, null, 2), 'utf8');
  const jsFile = path.join(OUT_DIR, 'data.js');
  fs.writeFileSync(jsFile, 'window.__DATA__ = ' + JSON.stringify(out) + ';', 'utf8');
  console.log('→', OUT_FILE);
  console.log('→', jsFile);
}

main().catch(err => {
  console.error('✗', err.message);
  if (err.response?.data) console.error(err.response.data);
  process.exit(1);
});
