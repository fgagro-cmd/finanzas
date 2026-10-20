'use strict';
/* Finanzas — situación y evolución mensual del patrimonio (cuentas, bienes, deudas y cobros),
   índices del antiguo Control_financiero.xlsx y presupuesto de gastos personales.
   Todo se guarda en el propio móvil (IndexedDB) y funciona sin conexión.
   El código no contiene datos personales: se cargan desde un fichero (Ajustes → Cargar copia). */

const APP_VERSION = '1.1.1';

/* ---------- Utilidades ---------- */
const $ = (s, r = document) => r.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const nf = (n, min = 2, max = 2) => Number(n).toLocaleString('es-ES', { minimumFractionDigits: min, maximumFractionDigits: max, useGrouping: 'always' });
const eur = (n) => nf(n) + ' €';
const eur0 = (n) => nf(n, 0, 0) + ' €';
const pctf = (p) => nf(p * 100, 0, 1) + ' %';
const ratiof = (r) => nf(r, 2, 2);
const sgn = (n, f = eur0) => (n > 0.004 ? '+' : n < -0.004 ? '−' : '') + f(Math.abs(n));
const round = (n, d) => { const k = 10 ** d; return Math.round((n + Number.EPSILON * Math.sign(n)) * k) / k; };
const r2 = (n) => round(n, 2);
const sum = (xs, f = (x) => x) => xs.reduce((t, x) => t + (Number(f(x)) || 0), 0);
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2, 10));
const byName = (a, b) => a.nombre.localeCompare(b.nombre, 'es');
const EPS = 0.005;

function parseNum(s) {
  s = String(s ?? '').trim().replace(/\s/g, '').replace(/[€%]/g, '');
  if (!s) return NaN;
  const c = s.lastIndexOf(','), d = s.lastIndexOf('.');
  if (c > -1 && d > -1) s = c > d ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (c > -1) s = s.replace(',', '.');
  else if (d > -1 && /^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, ''); // 1.250.000 → miles
  return /^-?(\d+\.?\d*|\.\d+)$/.test(s) ? parseFloat(s) : NaN;
}
const fmtIn = (n, d = 2) => (n === '' || n == null || !isFinite(n) ? '' : String(round(n, d)).replace('.', ','));

function todayISO() { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); }
const fmtDate = (iso) => iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '';
const fmtDay = (iso) => iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) : '';
const monthOf = (iso) => iso.slice(0, 7);
function addM(k, n) {
  let [y, m] = k.split('-').map(Number);
  m += n - 1; y += Math.floor(m / 12); m = ((m % 12) + 12) % 12 + 1;
  return `${y}-${String(m).padStart(2, '0')}`;
}
function rangeM(a, b) { const r = []; for (let k = a; k <= b && r.length < 240; k = addM(k, 1)) r.push(k); return r; }
function lastDay(ym) { const [y, m] = ym.split('-').map(Number); return `${ym}-${String(new Date(y, m, 0).getDate()).padStart(2, '0')}`; }
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const MC = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
function monthLabel(ym) { const [y, m] = ym.split('-').map(Number); const s = MESES[m - 1] + ' ' + y; return s.charAt(0).toUpperCase() + s.slice(1); }
const monthShort = (ym) => MC[Number(ym.slice(5, 7)) - 1] + ' ' + ym.slice(2, 4);
function ago(iso) {
  if (!iso) return 'nunca';
  const min = Math.round((Date.now() - new Date(iso).getTime()) / 60000);
  if (min < 1) return 'ahora mismo';
  if (min < 60) return `hace ${min} min`;
  if (min < 1440) return `hace ${Math.round(min / 60)} h`;
  return new Date(iso).toLocaleDateString('es-ES');
}

/* ---------- Almacenamiento (IndexedDB) ---------- */
let dbp;
function idb() {
  if (!dbp) dbp = new Promise((res, rej) => {
    const r = indexedDB.open('finanzas', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('kv');
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  return dbp;
}
async function dbGet(k) {
  const db = await idb();
  return new Promise((res, rej) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); });
}
async function dbSet(k, v) {
  const db = await idb();
  return new Promise((res, rej) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); });
}

/* ---------- Estado ---------- */
let state;
let saveChain = Promise.resolve();
function persist() {
  saveChain = saveChain.then(() => dbSet('state', state)).catch((e) => toast('No se pudo guardar en el móvil: ' + (e && e.message || e), true));
  markDirty();
  return saveChain;
}
function emptyState() {
  return {
    v: 1, creado: todayISO(),
    cuentas: [], bienes: [], prestamos: [], deudas: [], cobros: [],
    auto: {}, cierres: [],
    presu: { cfg: { nomina: 0, start: monthOf(todayISO()), cats: [] }, months: {} },
    settings: { comisiones: true, lastBackup: null },
  };
}
function validState(s) {
  return s && ['cuentas', 'bienes', 'prestamos', 'deudas', 'cobros', 'cierres'].every((k) => Array.isArray(s[k])) && s.presu && s.presu.cfg && s.settings;
}
const stateInfo = (s) => `${s.cuentas.length} cuentas, ${s.cierres.length} cierres de mes, ${s.deudas.length + s.cobros.length} deudas y cobros`;
const isEmpty = () => !state.cuentas.length && !state.bienes.length && !state.cierres.length;

function commit(mutator) {
  const c = structuredClone(state);
  try { mutator(c); } catch (e) { toast(e.message, true); return false; }
  state = c; persist(); renderAll();
  return true;
}

/* ---------- Cálculos (mismas fórmulas que Control_financiero.xlsx) ---------- */
const GRUPOS = { personal: 'Personal', empresa: 'Empresa', ventas: 'Ventas' };
const TIPOS = { banco: 'Banco', efectivo: 'Efectivo' };
const pagado = (x) => r2(sum(x.pagos || [], (p) => p.imp));
const pendiente = (x) => r2(Math.max(0, x.importe - pagado(x)));
const abierta = (x) => pendiente(x) > EPS;

function autoCobros(st = state) {
  const a = st.auto || {};
  const xs = [
    { k: 'ventas', nombre: 'Clientes de Ventas', sub: 'Deudas de clientes (sin tus fincas)', v: a.ventas ? a.ventas.pend : 0, src: 'ventas' },
  ];
  if (st.settings.comisiones) xs.push({ k: 'comisiones', nombre: 'Comisiones pendientes', sub: 'De la app de Ventas', v: a.ventas ? a.ventas.com : 0, src: 'ventas' });
  xs.push({ k: 'agro', nombre: 'Trabajos externos', sub: 'AgroServicios · sin cobrar', v: a.agro ? a.agro.pend : 0, src: 'agro' });
  xs.push({ k: 'clientes', nombre: 'Asesoramiento', sub: 'App Asesor · meses vencidos sin cobrar', v: a.clientes ? a.clientes.pend : 0, src: 'clientes' });
  return xs;
}
/* Las 8 cifras de partida de un mes. Todo lo demás se deriva de ellas. */
function inputs(st = state) {
  const a = st.auto || {};
  const banc = st.cuentas.filter((c) => c.tipo === 'banco'), efec = st.cuentas.filter((c) => c.tipo !== 'banco');
  const stockV = a.ventas ? a.ventas.stock : 0, stockF = a.stock ? a.stock.valor : 0;
  const cobAuto = sum(autoCobros(st), (x) => x.v), cobMan = sum(st.cobros, pendiente);
  const grupo = (g) => sum(st.cuentas.filter((c) => c.grupo === g), (c) => c.saldo);
  return {
    bancos: r2(sum(banc, (c) => c.saldo)),
    efectivo: r2(sum(efec, (c) => c.saldo)),
    stock: r2(stockV + stockF),
    maquinaria: r2(sum(st.bienes.filter((b) => b.maquinaria), (b) => b.valor)),
    otros: r2(sum(st.bienes.filter((b) => !b.maquinaria), (b) => b.valor)),
    prestamos: r2(sum(st.prestamos, (p) => p.capital)),
    deudas: r2(sum(st.deudas, pendiente)),
    cobros: r2(cobAuto + cobMan),
    // desglose (no entra en las fórmulas)
    x: { stockV: r2(stockV), stockF: r2(stockF), cobAuto: r2(cobAuto), cobMan: r2(cobMan), personal: r2(grupo('personal')), empresa: r2(grupo('empresa') + stockF), ventas: r2(grupo('ventas') + stockV) },
  };
}
function derive(i) {
  const activos = i.bancos + i.efectivo + i.stock + i.maquinaria + i.otros;
  const deuda = i.prestamos + i.deudas;
  const neto = activos - deuda + i.cobros;
  const liquidez = i.bancos + i.efectivo;
  const div = (a, b) => (Math.abs(b) > EPS ? a / b : null);
  return {
    ...i, activos: r2(activos), deuda: r2(deuda), neto: r2(neto), liquidez: r2(liquidez),
    ratio: div(liquidez, deuda), endeud: div(deuda, activos), indep: div(neto, activos), pstock: div(i.stock, neto),
    balance: r2(liquidez + i.cobros - i.deudas),
  };
}
const actual = () => derive(inputs());
const cierresAsc = () => [...state.cierres].sort((a, b) => a.mes.localeCompare(b.mes));
function cierreAnterior(mes) { const xs = cierresAsc().filter((c) => c.mes < mes); return xs[xs.length - 1] || null; }
const ultimoCierre = () => { const xs = cierresAsc(); return xs[xs.length - 1] || null; };

/* Semáforos con los umbrales anotados en el Excel */
const INDICES = {
  ratio: { nombre: 'Ratio liquidez / deuda', fmt: ratiof, obj: 'up', tramos: [[0.2, 'r', 'Riesgo alto'], [0.5, 'a', 'Aceptable'], [1, 'g', 'Sólido'], [Infinity, 'g', 'Excelente']], ayuda: 'Liquidez total ÷ deuda total. Menos de 0,2 riesgo alto · 0,2–0,5 aceptable · 0,5 sólido · 1 excelente.' },
  endeud: { nombre: 'Endeudamiento', fmt: ratiof, obj: 'down', tramos: [[0.3, 'g', 'Muy sano'], [0.5, 'a', 'Aceptable'], [Infinity, 'r', 'Sobreendeudamiento']], ayuda: 'Deuda total ÷ activos. Menos de 0,3 muy sano · 0,3–0,5 aceptable · desde 0,5 sobreendeudamiento.' },
  indep: { nombre: 'Independencia financiera', fmt: ratiof, obj: 'up', tramos: [[0.3, 'r', 'Muy endeudado'], [0.5, 'a', 'Endeudado medio'], [0.7, 'g', 'Sano'], [0.999, 'g', 'Muy independiente'], [Infinity, 'g', 'Sin deudas']], ayuda: 'Capital neto ÷ activos. Menos de 0,3 muy endeudado · 0,3–0,5 endeudado medio · 0,5–0,7 sano · más de 0,7 muy independiente · 1 sin deudas.' },
  pstock: { nombre: '% Stock', fmt: pctf, obj: null, tramos: null, ayuda: 'Valor del stock (ventas + fincas) ÷ capital neto. Sin objetivo: sirve para vigilar cuánto patrimonio está inmovilizado en producto.' },
  balance: { nombre: 'Balance de situación', fmt: eur0, obj: 'up', tramos: [[-EPS, 'r', 'Negativo'], [Infinity, 'g', 'Positivo']], ayuda: 'Liquidez + cobros pendientes − deudas pendientes (sin préstamos). Lo que quedaría si cobras y pagas todo lo pendiente hoy.' },
};
function semaforo(k, v) {
  const d = INDICES[k];
  if (v == null || !isFinite(v)) return ['n', 'Sin datos'];
  if (!d.tramos) return ['n', 'Informativo'];
  for (const [lim, c, txt] of d.tramos) if (v < lim) return [c, txt];
  return ['n', ''];
}

/* ---------- Datos de las otras apps (mismo origen fgagro-cmd.github.io) ----------
   Se leen directamente del almacenamiento del navegador de este móvil. No se escribe nada en ellas. */
const FUENTES = {
  ventas: { nombre: 'Ventas', url: '../ventas/' },
  stock: { nombre: 'Stock Fincas', url: '../stock-fincas/' },
  agro: { nombre: 'AgroServicios', url: '../agroservicios/' },
  clientes: { nombre: 'Asesor', url: '../asesor/' }, // antes la app Clientes; la clave 'clientes' se mantiene para no perder el dato ya guardado
};
const leido = {}; // fuente → true si se ha leído en esta sesión
function readOtherIDB(name) {
  return new Promise((res) => {
    let r;
    try { r = indexedDB.open(name); } catch (_) { return res(null); }
    // Si la base no existe, no la creamos: se aborta y se considera «sin datos».
    r.onupgradeneeded = () => { try { r.transaction.abort(); } catch (_) { /* nada */ } };
    r.onerror = () => res(null);
    r.onblocked = () => res(null);
    r.onsuccess = () => {
      const db = r.result;
      try {
        if (!db.objectStoreNames.contains('kv')) { db.close(); return res(null); }
        const q = db.transaction('kv').objectStore('kv').get('state');
        q.onsuccess = () => { db.close(); res(q.result || null); };
        q.onerror = () => { db.close(); res(null); };
      } catch (_) { try { db.close(); } catch (__) { /* nada */ } res(null); }
    };
  });
}
function calcVentas(s) {
  if (!s || !Array.isArray(s.compras) || !Array.isArray(s.pedidos)) return null;
  const stock = new Map();
  const get = (id) => { let x = stock.get(id); if (!x) { x = { q: 0, v: 0 }; stock.set(id, x); } return x; };
  for (const c of s.compras) { const x = get(c.prod); x.q += c.cant; x.v += c.cant * c.precio; }
  const propio = new Set((s.clientes || []).filter((c) => c.propio).map((c) => c.id));
  const pag = new Map();
  for (const c of s.cobros || []) for (const [pid, v] of Object.entries(c.aplic || {})) pag.set(pid, (pag.get(pid) || 0) + v);
  let pend = 0;
  for (const p of s.pedidos) {
    let imp = 0;
    for (const l of p.lineas) {
      imp += r2(l.cant * l.precio);
      if (l.prod) { const x = get(l.prod); x.q -= l.cant; x.v -= l.cant * l.coste; }
    }
    const pe = r2(Math.max(0, r2(imp) - r2(pag.get(p.id) || 0)));
    if (pe > EPS && !propio.has(p.cli)) pend += pe;
  }
  let valor = 0;
  for (const x of stock.values()) if (round(x.q, 6) > 1e-9) valor += round(x.v, 6);
  const com = sum((s.comisiones || []).filter((m) => !m.cobrado), (m) => r2(m.modo === 'fijo' ? m.fijo : m.cant * m.precio * m.pct));
  return { stock: r2(valor), pend: r2(pend), com: r2(com) };
}
function calcStock(s) {
  if (!s || !Array.isArray(s.productos)) return null;
  const m = new Map();
  for (const p of s.productos) m.set(p.id, { q: p.q0 || 0, v: p.v0 || 0 });
  for (const e of s.entradas || []) { const x = m.get(e.producto); if (x) { x.q += e.cantidad; x.v += e.importe; } }
  for (const o of s.movs || []) if (o.producto) { const x = m.get(o.producto); if (x) { x.q -= o.cantidad; x.v -= o.importe; } }
  let valor = 0;
  for (const x of m.values()) if (round(x.q, 4) > 0.0005) valor += x.v;
  return { valor: r2(valor) };
}
function calcAgro(s) {
  if (!s || !Array.isArray(s.trabajos)) return null;
  return { pend: r2(sum(s.trabajos.filter((t) => !t.cobrado), (t) => r2(t.horas * t.precioH))) };
}
function calcClientes(db, hoyYM = monthOf(todayISO())) {
  if (!db || !Array.isArray(db.clientes) || !db.años) return null;
  let pend = 0;
  const cli = new Map(db.clientes.map((c) => [c.id, c]));
  for (const [y, t] of Object.entries(db.años)) {
    for (const [id, r] of Object.entries(t || {})) {
      const c = cli.get(id); if (!c || !r) continue;
      for (let i = 0; i < 12; i++) {
        const k = `${y}-${String(i + 1).padStart(2, '0')}`;
        if (k > hoyYM) break; // se cobra por adelantado: el mes en curso ya está vencido
        if ((r.cob || [])[i]) continue;
        if ((c.alta && k < c.alta) || (c.baja && k > c.baja)) continue;
        pend += Number((r.imp || [])[i]) || 0;
      }
    }
  }
  return { pend: r2(pend) };
}
// Cobros de asesoría: ahora se anotan en Asesor (IndexedDB «asesor», mismos campos clientes y años que la app Clientes).
// Si Asesor aún no está instalada o no tiene tablas de planes y cobros, se usa la app Clientes como respaldo.
async function leerAsesoria() {
  try {
    const a = await readOtherIDB('asesor');
    if (a && Array.isArray(a.clientes) && a.años && Object.keys(a.años).length) {
      const r = calcClientes(a);
      if (r) return r;
    }
  } catch (_) { /* se prueba el respaldo */ }
  try { const raw = localStorage.getItem('clientes_v1'); return raw ? calcClientes(JSON.parse(raw)) : null; } catch (_) { return null; }
}
async function leerFuentes() {
  const res = {};
  try { res.ventas = calcVentas(await readOtherIDB('ventas-fgagro')); } catch (_) { res.ventas = null; }
  try { res.stock = calcStock(await readOtherIDB('stock-fincas')); } catch (_) { res.stock = null; }
  try { res.agro = calcAgro(await readOtherIDB('agroservicios')); } catch (_) { res.agro = null; }
  try { res.clientes = await leerAsesoria(); } catch (_) { res.clientes = null; }
  let cambio = false;
  const nuevo = { ...(state.auto || {}) };
  for (const [k, v] of Object.entries(res)) {
    if (!v) continue;
    leido[k] = true;
    const prev = nuevo[k];
    const same = prev && Object.keys(v).every((f) => Math.abs((prev[f] || 0) - v[f]) < EPS);
    if (!same) { nuevo[k] = { ...v, fecha: new Date().toISOString() }; cambio = true; }
  }
  if (cambio) { state = { ...state, auto: nuevo }; persist(); }
  return cambio;
}
function estadoFuente(k) {
  const a = (state.auto || {})[k];
  if (leido[k]) return { txt: 'al día', cls: 'auto' };
  if (a) return { txt: 'dato del ' + fmtDay(a.fecha.slice(0, 10)), cls: 'old' };
  return { txt: 'sin datos', cls: 'old' };
}

/* ---------- Cierres de mes ---------- */
function fotoActual() {
  const i = inputs();
  return {
    i: { bancos: i.bancos, efectivo: i.efectivo, stock: i.stock, maquinaria: i.maquinaria, otros: i.otros, prestamos: i.prestamos, deudas: i.deudas, cobros: i.cobros },
    det: {
      cuentas: [...state.cuentas].map((c) => [c.nombre, r2(c.saldo)]),
      stockVentas: i.x.stockV, stockFincas: i.x.stockF, cobrosApps: i.x.cobAuto, cobrosOtros: i.x.cobMan,
    },
  };
}
function cerrarMes(c, mes, origen) {
  const f = fotoActual();
  const prev = c.cierres.find((x) => x.mes === mes);
  const nuevo = { mes, fecha: lastDay(mes), ...f, comentario: prev ? prev.comentario : '', objetivo: prev ? prev.objetivo : '', origen, hecho: new Date().toISOString(), revisar: origen === 'auto' };
  c.cierres = c.cierres.filter((x) => x.mes !== mes).concat([nuevo]);
}
/* Al abrir la app ya en el mes siguiente se guarda la foto del mes que terminó. */
function autoCierre() {
  if (isEmpty()) return;
  const prev = addM(monthOf(todayISO()), -1);
  const last = ultimoCierre();
  const toca = last ? last.mes < prev : monthOf(state.creado) <= prev;
  if (!toca) return;
  commit((c) => cerrarMes(c, prev, 'auto'));
  toast(`Cierre de ${monthLabel(prev).toLowerCase()} guardado. Revísalo y añade un comentario.`);
}

/* ---------- UI general ---------- */
const ui = { tab: 'inicio', dseg: 'deudas', pseg: 'resumen', pmes: monthOf(todayISO()), pmodo: 'mes', pfiltro: null };
let toastTimer;
function toast(msg, err = false) {
  const t = $('#toast'); t.textContent = msg; t.className = 'show' + (err ? ' err' : '');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.className = ''; }, err ? 4200 : 2600);
}
const sheetRoot = $('#sheetRoot');
let sheetOpen = false;
function openSheet(title, bodyHTML) {
  if (!sheetOpen) history.pushState({ sheet: true }, '');
  sheetOpen = true;
  sheetRoot.innerHTML = `<div class="scrim"></div><section class="sheet" role="dialog" aria-modal="true" aria-label="${esc(title)}"><header><h2>${esc(title)}</h2><button class="x" aria-label="Cerrar">✕</button></header><div class="sheet-body">${bodyHTML}</div></section>`;
  $('.scrim', sheetRoot).onclick = closeSheet;
  $('.x', sheetRoot).onclick = closeSheet;
  return $('.sheet-body', sheetRoot);
}
function closeSheet() {
  if (!sheetOpen) return;
  sheetOpen = false; sheetRoot.innerHTML = '';
  if (history.state && history.state.sheet) history.back();
}
window.addEventListener('popstate', () => { if (sheetOpen) { sheetOpen = false; sheetRoot.innerHTML = ''; } });
const reopen = (fn) => { closeSheet(); setTimeout(fn, 60); };
function armDelete(btn, action, label = 'Borrar') {
  let armed = false, timer;
  btn.onclick = () => {
    if (!armed) { armed = true; btn.classList.add('armed'); btn.textContent = 'Pulsa otra vez para confirmar'; timer = setTimeout(() => { armed = false; btn.classList.remove('armed'); btn.textContent = label; }, 3500); }
    else { clearTimeout(timer); action(); }
  };
}
const field = (id, label, value, { mode = 'decimal', ph = '', type = 'text' } = {}) =>
  `<label class="fld"><span>${label}</span><input id="${id}" type="${type}" ${type === 'text' ? `inputmode="${mode}"` : ''} value="${esc(value)}" placeholder="${esc(ph)}" autocomplete="off"></label>`;
const segBtns = (id, opts, sel) => `<div class="seg" id="${id}">${opts.map(([v, t]) => `<button type="button" data-v="${v}" aria-pressed="${v === sel}">${t}</button>`).join('')}</div>`;
function bindSeg(root, id, cb) { const el = $('#' + id, root); el.onclick = (e) => { const b = e.target.closest('[data-v]'); if (!b) return; el.querySelectorAll('button').forEach((x) => x.setAttribute('aria-pressed', x === b)); cb(b.dataset.v); }; }
const trendTxt = (d, obj) => {
  if (d == null || !isFinite(d) || Math.abs(d) < EPS) return '';
  const good = obj === 'up' ? d > 0 : obj === 'down' ? d < 0 : null;
  return `<b class="${good == null ? '' : good ? 'up' : 'down'}">${d > 0 ? '▲' : '▼'}</b>`;
};

/* ---------- Gráficas (SVG propio, sin librerías) ---------- */
function niceStep(v) { if (v <= 0) return 1; const e = 10 ** Math.floor(Math.log10(v)); const f = v / e; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * e; }
const kfmt = (v) => (Math.abs(v) >= 1000 ? nf(v / 1000, 0, 0) + 'k' : nf(v, 0, 0));
function chartSVG(series, labels, { bars = false } = {}) {
  const W = 340, H = 170, L = 36, R = 10, T = 12, B = 22;
  const all = series.flatMap((s) => s.vals).filter((v) => v != null);
  let lo = Math.min(0, ...all), hi = Math.max(...all, 1);
  if (!bars && lo >= 0) lo = Math.min(...all) * 0.9;
  const st = niceStep((hi - lo) / 4 || 1);
  lo = Math.floor(lo / st) * st; hi = Math.ceil(hi / st) * st; if (hi === lo) hi = lo + st;
  const n = labels.length, pw = W - L - R, ph = H - T - B;
  const x = (i) => L + (bars ? (i + 0.5) * pw / n : (n === 1 ? pw / 2 : i * pw / (n - 1)));
  const y = (v) => T + ph - (v - lo) / (hi - lo) * ph;
  let g = '<g class="grid">';
  for (let v = lo; v <= hi + st / 2; v += st) g += `<line x1="${L}" x2="${W - R}" y1="${y(v).toFixed(1)}" y2="${y(v).toFixed(1)}"/>`;
  g += '</g><g class="axis">';
  for (let v = lo; v <= hi + st / 2; v += st) g += `<text x="${L - 5}" y="${(y(v) + 3).toFixed(1)}" text-anchor="end">${kfmt(v)}</text>`;
  const every = Math.max(1, Math.ceil(n / 6));
  labels.forEach((l, i) => { if (i === n - 1 || (i % every === 0 && n - 1 - i >= Math.max(1, every * 0.75))) g += `<text x="${x(i).toFixed(1)}" y="${H - 6}" text-anchor="middle">${esc(l)}</text>`; });
  g += '</g>';
  let marks = '';
  series.forEach((s, si) => {
    const cls = si === 0 ? '1' : '2';
    if (bars) {
      const bw = Math.min(26, pw / n - 4);
      s.vals.forEach((v, i) => {
        if (v == null) return;
        const y0 = y(Math.max(0, lo)), y1 = y(v), top = Math.min(y0, y1), h = Math.max(1, Math.abs(y0 - y1));
        const r = Math.min(4, h / 2, bw / 2);
        // extremo de datos redondeado, base recta
        marks += v >= 0
          ? `<path class="f${cls}" d="M${x(i) - bw / 2},${y0} V${top + r} Q${x(i) - bw / 2},${top} ${x(i) - bw / 2 + r},${top} H${x(i) + bw / 2 - r} Q${x(i) + bw / 2},${top} ${x(i) + bw / 2},${top + r} V${y0} Z"/>`
          : `<rect class="f${cls}" x="${x(i) - bw / 2}" y="${top}" width="${bw}" height="${h}" rx="2"/>`;
      });
    } else {
      const pts = s.vals.map((v, i) => (v == null ? null : [x(i), y(v)])).filter(Boolean);
      marks += `<polyline class="line l${cls}" points="${pts.map((p) => p.map((c) => c.toFixed(1)).join(',')).join(' ')}"/>`;
      if (pts.length) { const p = pts[pts.length - 1]; marks += `<circle class="dot f${cls}" cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="4.5"/>`; }
    }
  });
  return `<svg viewBox="0 0 ${W} ${H}" role="img">${g}${marks}<g class="hov"></g><rect class="hit" x="${L}" y="${T}" width="${pw}" height="${ph}" fill="transparent"/></svg>`;
}
function bindChart(box, series, labels, opts = {}) {
  const svg = $('svg', box), read = $('.read', box), hov = $('.hov', box);
  const W = 340, L = 36, R = 10, T = 12, B = 22, H = 170, pw = W - L - R, n = labels.length;
  const fmt = opts.fmt || eur0;
  const show = (i) => {
    read.innerHTML = `${esc(opts.full ? opts.full[i] : labels[i])}: ` + series.map((s) => `${series.length > 1 ? esc(s.name) + ' ' : ''}<b>${s.vals[i] == null ? '—' : fmt(s.vals[i])}</b>`).join(' · ');
    const xi = L + (opts.bars ? (i + 0.5) * pw / n : (n === 1 ? pw / 2 : i * pw / (n - 1)));
    hov.innerHTML = `<line class="cross" x1="${xi}" x2="${xi}" y1="${T}" y2="${H - B}"/>`;
  };
  const pick = (ev) => {
    const r = svg.getBoundingClientRect();
    const px = (ev.clientX - r.left) / r.width * W;
    const i = opts.bars ? Math.floor((px - L) / (pw / n)) : Math.round((px - L) / (pw / Math.max(1, n - 1)));
    show(Math.max(0, Math.min(n - 1, i)));
  };
  svg.addEventListener('pointerdown', pick);
  svg.addEventListener('pointermove', (e) => { if (e.pointerType === 'mouse' || e.buttons) pick(e); });
  show(n - 1);
}

/* ---------- Pestaña: Dashboard ---------- */
function headInicio() {
  if (isEmpty()) return '';
  const a = actual(), u = ultimoCierre(), du = u ? derive(u.i) : null;
  const dv = du ? a.neto - du.neto : null;
  return `<div class="kpis">
    <div class="k hero"><span>Capital neto</span><strong>${eur0(a.neto)}</strong></div>
    ${du ? `<div class="k line" style="border-top:0;padding-top:0;margin-top:-4px"><span>Desde el cierre de ${monthLabel(u.mes).toLowerCase()}</span><strong style="font-size:16px">${sgn(dv)}</strong></div>` : ''}
    <div class="k"><span>Activos</span><strong>${eur0(a.activos)}</strong></div>
    <div class="k"><span>Deuda total</span><strong>${eur0(a.deuda)}</strong></div>
    <div class="k"><span>Liquidez</span><strong>${eur0(a.liquidez)}</strong></div>
    <div class="k"><span>Cobros pendientes</span><strong>${eur0(a.cobros)}</strong></div>
  </div>`;
}
function renderInicio() {
  if (isEmpty()) {
    $('#view').innerHTML = `<div class="welcome"><h3 style="margin-top:0">Bienvenido a Finanzas</h3><p>Para empezar, carga el fichero con tus datos (histórico del Excel, cuentas, bienes y presupuesto).</p><div class="actions"><label class="btn primary" style="display:grid;place-items:center;cursor:pointer">Cargar fichero de datos<input id="wLoad" type="file" accept="application/json,.json" hidden></label><button class="btn" id="wNew">Empezar desde cero</button></div></div>`;
    $('#wLoad').onchange = (e) => restoreFromFile(e.target);
    $('#wNew').onclick = () => { ui.tab = 'cuentas'; renderAll(); };
    return;
  }
  const a = actual(), u = ultimoCierre(), du = u ? derive(u.i) : null;
  const xs = cierresAsc();
  const revisar = xs.filter((c) => c.revisar);
  const mesAct = monthOf(todayISO());
  const cerrado = xs.some((c) => c.mes === mesAct);
  const card = (k, v, prev, extra = '') => {
    const d = INDICES[k], [cls, txt] = semaforo(k, v);
    const dif = prev != null && v != null ? v - prev : null;
    return `<button class="ic" data-idx="${k}"><div class="lbl">${d.nombre}</div><div class="val">${v == null ? '—' : d.fmt(v)}</div><span class="st ${cls}">${txt}</span>${dif != null && Math.abs(dif) > (k === 'balance' ? 0.5 : 0.0005) ? ` <span class="trend">${trendTxt(dif, d.obj)} ${k === 'balance' ? sgn(dif) : k === 'pstock' ? sgn(dif * 100, (z) => nf(z, 0, 1) + ' pt') : sgn(dif, (z) => nf(z, 2, 2))}</span>` : ''}${extra}</button>`;
  };
  // Series para las gráficas: cierres + «hoy» si el mes en curso no está cerrado
  const pts = xs.map((c) => ({ lab: monthShort(c.mes), full: 'Cierre de ' + monthLabel(c.mes).toLowerCase(), d: derive(c.i) }));
  if (!cerrado) pts.push({ lab: 'Hoy', full: 'Hoy', d: a });
  const labels = pts.map((p) => p.lab), full = pts.map((p) => p.full);
  const sNeto = [{ name: 'Capital neto', vals: pts.map((p) => p.d.neto) }];
  const sAD = [{ name: 'Activos', vals: pts.map((p) => p.d.activos) }, { name: 'Deuda', vals: pts.map((p) => p.d.deuda) }];
  const sLiq = [{ name: 'Liquidez', vals: pts.map((p) => p.d.liquidez) }];
  $('#view').innerHTML = `
    ${revisar.map((c) => `<button class="banner" data-mes="${c.mes}"><span>Cierre de ${monthLabel(c.mes).toLowerCase()} hecho automáticamente: revísalo y añade el comentario</span><span>›</span></button>`).join('')}
    <div class="sec"><h3>Índices ahora</h3><span class="small" style="font-weight:500">${du ? 'vs. cierre de ' + monthShort(u.mes) : ''}</span></div>
    <div class="idx" id="idx">
      ${card('ratio', a.ratio, du && du.ratio)}
      ${card('endeud', a.endeud, du && du.endeud)}
      ${card('indep', a.indep, du && du.indep)}
      ${card('pstock', a.pstock, du && du.pstock)}
      ${card('balance', a.balance, du && du.balance)}
      <button class="ic" data-idx="var"><div class="lbl">Variación capital neto</div><div class="val">${du ? sgn(a.neto - du.neto) : '—'}</div><span class="st ${!du ? 'n' : a.neto >= du.neto ? 'g' : 'r'}">${du && Math.abs(du.neto) > EPS ? sgn((a.neto - du.neto) / du.neto * 100, (z) => nf(z, 1, 1) + ' %') : 'Sin cierre previo'}</span></button>
    </div>
    <div class="chart" id="ch1"><h4>Capital neto</h4><div class="read"></div>${chartSVG(sNeto, labels)}</div>
    <div class="chart" id="ch2"><h4>Activos y deuda</h4><div class="legend"><span><i class="k1"></i>Activos</span><span><i class="k2"></i>Deuda total</span></div><div class="read"></div>${chartSVG(sAD, labels)}</div>
    <div class="chart" id="ch3"><h4>Liquidez (bancos + efectivo)</h4><div class="read"></div>${chartSVG(sLiq, labels, { bars: true })}</div>
    <div class="sec"><h3>Cierres de mes</h3><span></span></div>
    ${cerrado ? '' : `<div class="actions" style="margin-bottom:10px"><button class="btn" id="closeNow">Cerrar ${monthLabel(mesAct).toLowerCase()} ahora</button></div>`}
    <ul class="list hist" id="hist">${[...xs].reverse().map((c) => {
      const d = derive(c.i), p = cierreAnterior(c.mes), dp = p ? derive(p.i) : null;
      const v = dp ? d.neto - dp.neto : null;
      return `<li><button class="row" data-mes="${c.mes}"><div class="main"><div class="name">${monthLabel(c.mes)} ${c.revisar ? '<span class="tag old">revisar</span>' : ''}</div><div class="sub">${esc((c.comentario || '').slice(0, 70)) || (c.origen === 'excel' ? 'Del Excel' : 'Sin comentario')}${(c.comentario || '').length > 70 ? '…' : ''}</div></div><div class="side"><div class="big">${eur0(d.neto)}</div><div class="sub ${v != null && v < 0 ? 'neg' : ''}">${v == null ? '' : sgn(v)}</div></div></button></li>`;
    }).join('')}</ul>`;
  $('#view').querySelectorAll('.banner').forEach((b) => { b.onclick = () => openCierre(b.dataset.mes); });
  $('#hist').onclick = (e) => { const b = e.target.closest('[data-mes]'); if (b) openCierre(b.dataset.mes); };
  $('#idx').onclick = (e) => { const b = e.target.closest('[data-idx]'); if (b) openIndice(b.dataset.idx); };
  const cn = $('#closeNow'); if (cn) cn.onclick = () => openCerrarAhora();
  bindChart($('#ch1'), sNeto, labels, { full });
  bindChart($('#ch2'), sAD, labels, { full });
  bindChart($('#ch3'), sLiq, labels, { full, bars: true });
}
function openIndice(k) {
  if (k === 'var') return openSheet('Variación capital neto', '<p>Diferencia entre el capital neto de ahora y el del último cierre de mes.</p><p class="small">Capital neto = activos − deuda total + cobros pendientes (igual que en el Excel).</p>');
  const d = INDICES[k];
  const xs = cierresAsc().slice(-12);
  openSheet(d.nombre, `<p>${esc(d.ayuda)}</p>${d.obj ? `<p class="small">Objetivo: ${d.obj === 'up' ? 'aumentar' : 'reducir'}.</p>` : ''}
    <h3>Últimos cierres</h3><div class="box">${xs.map((c) => { const v = derive(c.i)[k]; const [cls, txt] = semaforo(k, v); return `<div><span>${monthLabel(c.mes)}</span><b>${v == null ? '—' : d.fmt(v)} <span class="st ${cls}" style="margin-left:6px">${txt}</span></b></div>`; }).join('')}</div>`);
}
function openCerrarAhora() {
  const mes = monthOf(todayISO());
  const a = actual();
  const body = openSheet('Cerrar ' + monthLabel(mes).toLowerCase(), `
    <p class="small" style="margin-top:0">Se guarda la foto de hoy como cierre de ${monthLabel(mes).toLowerCase()} (${fmtDate(lastDay(mes))}). Si cambias algo después, puedes rehacerla desde el histórico. Si no lo cierras tú, se cierra solo al abrir la app el mes que viene.</p>
    <div class="box"><div><span>Activos</span><b>${eur(a.activos)}</b></div><div><span>Deuda total</span><b>${eur(a.deuda)}</b></div><div><span>Cobros pendientes</span><b>${eur(a.cobros)}</b></div><div class="tot"><span>Capital neto</span><b>${eur(a.neto)}</b></div></div>
    <label class="fld"><span>Comentario del mes</span><textarea id="cCom" placeholder="Qué ha pasado este mes (ingresos, pagos, compras…)"></textarea></label>
    <label class="fld"><span>Objetivo para el mes que viene</span><input id="cObj" type="text" autocomplete="off" placeholder="Ej.: reducir deuda, ratio 0,25…"></label>
    <div class="actions"><button class="btn primary" id="cSave">Cerrar el mes</button></div>`);
  $('#cSave', body).onclick = () => {
    const com = $('#cCom', body).value.trim(), obj = $('#cObj', body).value.trim();
    if (commit((c) => { cerrarMes(c, mes, 'manual'); const x = c.cierres.find((z) => z.mes === mes); x.comentario = com; x.objetivo = obj; x.revisar = false; })) { closeSheet(); toast('Mes cerrado'); }
  };
}
const CAMPOS = [['bancos', 'Total bancos'], ['efectivo', 'Total efectivo'], ['stock', 'Total stock'], ['maquinaria', 'Valor maquinaria'], ['otros', 'Otros activos'], ['prestamos', 'Préstamos'], ['deudas', 'Deudas pendientes'], ['cobros', 'Cobros pendientes']];
function openCierre(mes) {
  const c = state.cierres.find((x) => x.mes === mes); if (!c) return;
  const d = derive(c.i), p = cierreAnterior(mes), dp = p ? derive(p.i) : null;
  const vr = (k, f = eur) => (dp ? ` <span class="small">(${sgn(d[k] - dp[k], f)})</span>` : '');
  const idxRow = (k) => { const v = d[k]; const [cls, txt] = semaforo(k, v); return `<div><span>${INDICES[k].nombre}</span><b>${v == null ? '—' : INDICES[k].fmt(v)} <span class="st ${cls}" style="margin-left:4px">${txt}</span></b></div>`; };
  const esUlt = ultimoCierre() && ultimoCierre().mes === mes;
  const body = openSheet('Cierre de ' + monthLabel(mes).toLowerCase(), `
    <p class="small" style="margin-top:0">A ${fmtDate(c.fecha)} · ${c.origen === 'excel' ? 'traído del Excel' : c.origen === 'auto' ? 'guardado automáticamente' : 'cerrado a mano'}</p>
    <div class="box">
      <div><span>Activos</span><b>${eur(d.activos)}${vr('activos')}</b></div>
      <div><span>Deuda total</span><b>${eur(d.deuda)}${vr('deuda')}</b></div>
      <div><span>Cobros pendientes</span><b>${eur(d.cobros)}</b></div>
      <div class="tot"><span>Capital neto</span><b>${eur(d.neto)}${vr('neto')}</b></div>
      <div><span>Liquidez</span><b>${eur(d.liquidez)}</b></div>
    </div>
    <div class="box">${idxRow('ratio')}${idxRow('endeud')}${idxRow('indep')}${idxRow('pstock')}${idxRow('balance')}</div>
    ${c.det && c.det.cuentas && c.det.cuentas.length ? `<details class="more" style="margin:0 0 12px"><summary>Saldo de cada cuenta</summary><div class="box" style="margin-top:6px">${c.det.cuentas.map(([n, v]) => `<div><span>${esc(n)}</span><b>${eur(v)}</b></div>`).join('')}${c.det.stockVentas != null ? `<div><span>Stock ventas</span><b>${eur(c.det.stockVentas)}</b></div><div><span>Stock fincas</span><b>${eur(c.det.stockFincas)}</b></div>` : ''}</div></details>` : ''}
    <label class="fld"><span>Comentario</span><textarea id="cCom">${esc(c.comentario || '')}</textarea></label>
    <label class="fld"><span>Objetivo para el mes siguiente</span><input id="cObj" type="text" value="${esc(c.objetivo || '')}" autocomplete="off"></label>
    <details class="more" style="margin:0 0 12px"><summary>Corregir cifras del cierre</summary>
      <div class="grid2" style="margin-top:8px">${CAMPOS.map(([k, t]) => field('cf_' + k, t + ' €', fmtIn(c.i[k]))).join('')}</div>
    </details>
    <div class="actions"><button class="btn primary" id="cSave">Guardar</button>
    ${esUlt ? '<button class="btn" id="cRedo">Rehacer con los datos de ahora</button>' : ''}</div>
    <div class="sep"></div><div class="actions"><button class="btn danger" id="cDel">Borrar este cierre</button></div>`);
  $('#cSave', body).onclick = () => {
    const vals = {};
    for (const [k, t] of CAMPOS) { const n = parseNum($('#cf_' + k, body).value); if (!isFinite(n)) return toast('Revisa «' + t + '»', true); vals[k] = r2(n); }
    if (commit((s) => { const x = s.cierres.find((z) => z.mes === mes); x.comentario = $('#cCom', body).value.trim(); x.objetivo = $('#cObj', body).value.trim(); x.i = vals; x.revisar = false; })) { closeSheet(); toast('Cierre guardado'); }
  };
  const rd = $('#cRedo', body);
  if (rd) armDelete(rd, () => { if (commit((s) => { const com = $('#cCom', body).value.trim(), obj = $('#cObj', body).value.trim(); cerrarMes(s, mes, c.origen === 'excel' ? 'manual' : c.origen); const x = s.cierres.find((z) => z.mes === mes); x.comentario = com; x.objetivo = obj; x.revisar = false; })) { closeSheet(); toast('Cierre rehecho con los datos de ahora'); } }, 'Rehacer con los datos de ahora');
  armDelete($('#cDel', body), () => { if (commit((s) => { s.cierres = s.cierres.filter((z) => z.mes !== mes); })) { closeSheet(); toast('Cierre borrado'); } }, 'Borrar este cierre');
}

/* ---------- Pestaña: Cuentas ---------- */
function headCuentas() {
  const i = inputs(), d = derive(i);
  return `<div class="kpis">
    <div class="k hero"><span>Liquidez total</span><strong>${eur0(d.liquidez)}</strong></div>
    <div class="k"><span>Bancos</span><strong>${eur0(i.bancos)}</strong></div>
    <div class="k"><span>Efectivo</span><strong>${eur0(i.efectivo)}</strong></div>
    <div class="k line"><span>Total activos</span><strong>${eur0(d.activos)}</strong></div>
  </div>`;
}
function renderCuentas() {
  const i = inputs();
  const ord = { personal: 0, empresa: 1, ventas: 2 };
  const cs = [...state.cuentas].sort((a, b) => ord[a.grupo] - ord[b.grupo] || byName(a, b));
  const maq = state.bienes.filter((b) => b.maquinaria).sort((a, b) => b.valor - a.valor || byName(a, b));
  const otros = state.bienes.filter((b) => !b.maquinaria).sort((a, b) => b.valor - a.valor || byName(a, b));
  const bienRow = (b) => `<li><button class="row" data-id="${b.id}"><div class="main"><div class="name">${esc(b.nombre)}</div><div class="sub">${b.anio ? 'Compra ' + b.anio + (b.compra ? ' · ' + eur0(b.compra) : '') + ' · ' : ''}${b.fecha ? 'valor a ' + fmtDate(b.fecha) : 'valor actual estimado'}</div></div><div class="side"><div class="big">${eur0(b.valor)}</div></div></button></li>`;
  const auto = (k, nombre, v, src) => { const e = estadoFuente(src); return `<li><button class="row auto" data-src="${src}"><div class="main"><div class="name">${nombre}</div><div class="sub">De ${FUENTES[src].nombre} · <span class="tag ${e.cls}">${e.txt}</span></div></div><div class="side"><div class="big">${eur(v)}</div></div></button></li>`; };
  $('#view').innerHTML = `
    <div class="summary"><span>Personal <b>${eur0(i.x.personal)}</b></span><span>Empresa <b>${eur0(i.x.empresa)}</b></span><span>Ventas <b>${eur0(i.x.ventas)}</b></span></div>
    <div class="sec"><h3>Bancos y efectivo</h3><span>${eur(i.bancos + i.efectivo)}</span></div>
    ${cs.length ? `<ul class="list" id="ctas">${cs.map((c) => `<li><button class="row" data-id="${c.id}"><div class="main"><div class="name">${esc(c.nombre)}</div><div class="sub">${GRUPOS[c.grupo]} · ${TIPOS[c.tipo]}${c.fecha ? ' · ' + fmtDay(c.fecha) : ''}</div></div><div class="side"><div class="big ${c.saldo < 0 ? 'neg' : ''}">${eur(c.saldo)}</div></div></button></li>`).join('')}</ul>` : '<p class="empty">Añade tus cuentas con el botón +</p>'}
    <div class="sec"><h3>Stock</h3><span>${eur(i.stock)}</span></div>
    <ul class="list" id="autoList">${auto('sv', 'Stock de ventas', i.x.stockV, 'ventas')}${auto('sf', 'Stock de tus fincas', i.x.stockF, 'stock')}</ul>
    <div class="sec"><h3>Bienes</h3><span>${eur(i.maquinaria + i.otros)}</span></div>
    <div id="bienes">
      <div class="sec" style="margin-top:6px"><h3 style="text-transform:none;letter-spacing:0;font-size:14.5px;color:var(--ink)">Maquinaria · ${maq.length} ${maq.length === 1 ? 'máquina' : 'máquinas'}</h3><span>${eur(i.maquinaria)}</span></div>
      ${maq.length ? `<ul class="list">${maq.map(bienRow).join('')}</ul>` : '<p class="empty" style="padding:14px">Sin maquinaria</p>'}
      <div class="sec"><h3 style="text-transform:none;letter-spacing:0;font-size:14.5px;color:var(--ink)">Otros activos</h3><span>${eur(i.otros)}</span></div>
      ${otros.length ? `<ul class="list">${otros.map(bienRow).join('')}</ul>` : '<p class="empty" style="padding:14px">Sin otros bienes</p>'}
    </div>`;
  const ct = $('#ctas'); if (ct) ct.onclick = (e) => { const b = e.target.closest('[data-id]'); if (b) openSaldo(b.dataset.id); };
  const bn = $('#bienes'); if (bn) bn.onclick = (e) => { const b = e.target.closest('[data-id]'); if (b) openBien(b.dataset.id); };
  $('#autoList').onclick = (e) => { const b = e.target.closest('[data-src]'); if (b) openFuente(b.dataset.src); };
}
function openAnadirCuentas() {
  const body = openSheet('Añadir', `<div class="actions"><button class="btn primary" id="aC">Cuenta bancaria o caja de efectivo</button><button class="btn" id="aB">Bien (casa, vehículo, maquinaria…)</button></div>`);
  $('#aC', body).onclick = () => reopen(() => openCuenta(null));
  $('#aB', body).onclick = () => reopen(() => openBien(null));
}
/* Hoja rápida para poner el saldo de hoy */
function openSaldo(id) {
  const c = state.cuentas.find((x) => x.id === id); if (!c) return;
  const body = openSheet(c.nombre, `
    <p class="small" style="margin-top:0">${GRUPOS[c.grupo]} · ${TIPOS[c.tipo]}${c.fecha ? ' · último cambio ' + fmtDate(c.fecha) : ''}</p>
    ${field('sV', 'Saldo de hoy €', fmtIn(c.saldo))}
    <div class="actions"><button class="btn primary" id="sSave">Guardar saldo</button><button class="btn" id="sEdit">Editar cuenta</button></div>`);
  const i = $('#sV', body);
  const save = () => {
    const n = parseNum(i.value); if (!isFinite(n)) return toast('Escribe un saldo válido', true);
    if (commit((s) => { const x = s.cuentas.find((z) => z.id === id); x.saldo = r2(n); x.fecha = todayISO(); })) { closeSheet(); toast(c.nombre + ': ' + eur(n)); }
  };
  $('#sSave', body).onclick = save;
  i.onkeydown = (e) => { if (e.key === 'Enter') save(); };
  $('#sEdit', body).onclick = () => reopen(() => openCuenta(id));
  setTimeout(() => { i.focus(); i.select(); }, 60);
}
function openCuenta(id) {
  const isNew = !id;
  const c = isNew ? { nombre: '', grupo: 'personal', tipo: 'banco', saldo: 0 } : state.cuentas.find((x) => x.id === id);
  let grupo = c.grupo, tipo = c.tipo;
  const body = openSheet(isNew ? 'Cuenta nueva' : 'Editar cuenta', `
    <label class="fld"><span>Nombre</span><input id="kN" type="text" value="${esc(c.nombre)}" autocomplete="off" placeholder="Ej.: Cajamar (personal)"></label>
    <div class="fld"><span>Grupo</span>${segBtns('kG', Object.entries(GRUPOS), grupo)}</div>
    <div class="fld"><span>Tipo</span>${segBtns('kT', Object.entries(TIPOS), tipo)}</div>
    ${isNew ? field('kS', 'Saldo de hoy €', '') : ''}
    <p class="hint">Efectivo suma en «Total efectivo» y Banco en «Total bancos», como en el Excel. La tarjeta de crédito va en Deudas.</p>
    <div class="actions"><button class="btn primary" id="kSave">Guardar</button></div>
    ${isNew ? '' : '<div class="sep"></div><div class="actions"><button class="btn danger" id="kDel">Borrar cuenta</button></div>'}`);
  bindSeg(body, 'kG', (v) => { grupo = v; });
  bindSeg(body, 'kT', (v) => { tipo = v; });
  $('#kSave', body).onclick = () => {
    const nombre = $('#kN', body).value.trim(); if (!nombre) return toast('Pon un nombre', true);
    let saldo = 0; if (isNew) { saldo = parseNum($('#kS', body).value || '0'); if (!isFinite(saldo)) return toast('Saldo no válido', true); }
    if (commit((s) => { if (isNew) s.cuentas.push({ id: uid(), nombre, grupo, tipo, saldo: r2(saldo), fecha: todayISO() }); else Object.assign(s.cuentas.find((z) => z.id === id), { nombre, grupo, tipo }); })) { closeSheet(); toast('Cuenta guardada'); }
  };
  const del = $('#kDel', body);
  if (del) armDelete(del, () => { if (commit((s) => { s.cuentas = s.cuentas.filter((z) => z.id !== id); })) { closeSheet(); toast('Cuenta borrada'); } }, 'Borrar cuenta');
}
function openBien(id) {
  const isNew = !id;
  const b = isNew ? { nombre: '', valor: '', maquinaria: false } : state.bienes.find((x) => x.id === id);
  let maq = !!b.maquinaria;
  const body = openSheet(isNew ? 'Bien nuevo' : b.nombre, `
    ${field('bV', 'Valor actual estimado €', fmtIn(b.valor))}
    <label class="fld"><span>Nombre</span><input id="bN" type="text" value="${esc(b.nombre)}" autocomplete="off"></label>
    <button type="button" class="switch" id="bM" role="switch" aria-checked="${maq}"><span>Es maquinaria (suma en «Valor maquinaria»)</span><i></i></button>
    <p class="hint">Cada máquina lleva su valor actual, que actualizas tú si se revaloriza o se deprecia. La suma de todas es el «Valor maquinaria» del cierre.</p>
    <div class="actions"><button class="btn primary" id="bSave">Guardar</button></div>
    ${isNew ? '' : '<div class="sep"></div><div class="actions"><button class="btn danger" id="bDel">Borrar</button></div>'}`);
  $('#bM', body).onclick = (e) => { maq = !maq; e.currentTarget.setAttribute('aria-checked', maq); };
  const save = () => {
    const nombre = $('#bN', body).value.trim(), valor = parseNum($('#bV', body).value);
    if (!nombre) return toast('Pon un nombre', true);
    if (!isFinite(valor)) return toast('Escribe un valor', true);
    if (commit((s) => { if (isNew) s.bienes.push({ id: uid(), nombre, valor: r2(valor), maquinaria: maq, fecha: todayISO() }); else { const x = s.bienes.find((z) => z.id === id); if (Math.abs(x.valor - valor) > EPS) x.fecha = todayISO(); Object.assign(x, { nombre, valor: r2(valor), maquinaria: maq }); } })) { closeSheet(); toast('Guardado'); }
  };
  $('#bSave', body).onclick = save;
  const del = $('#bDel', body);
  if (del) armDelete(del, () => { if (commit((s) => { s.bienes = s.bienes.filter((z) => z.id !== id); })) { closeSheet(); toast('Borrado'); } });
  if (!isNew) setTimeout(() => { const i = $('#bV', body); i.focus(); i.select(); }, 60);
}
function openFuente(src) {
  const f = FUENTES[src], a = (state.auto || {})[src], e = estadoFuente(src);
  const filas = {
    ventas: a ? [['Stock de ventas (a precio medio)', a.stock], ['Deudas de clientes', a.pend], ['Comisiones pendientes', a.com]] : [],
    stock: a ? [['Valor del stock de tus fincas', a.valor]] : [],
    agro: a ? [['Trabajos externos sin cobrar', a.pend]] : [],
    clientes: a ? [['Meses vencidos sin cobrar', a.pend]] : [],
  }[src];
  const body = openSheet(f.nombre, `
    <p class="small" style="margin-top:0">La app de Finanzas lee estas cifras directamente de la app ${f.nombre} instalada en este mismo móvil, sin conexión y sin tocar nada en ella. Estado: <span class="tag ${e.cls}">${e.txt}</span>${a ? ' · último cambio ' + ago(a.fecha) : ''}.</p>
    ${filas.length ? `<div class="box">${filas.map(([t, v]) => `<div><span>${t}</span><b>${eur(v)}</b></div>`).join('')}</div>` : '<p>Todavía no hay datos. Abre la app ' + f.nombre + ' en este móvil (con Chrome) al menos una vez.</p>'}
    ${leido[src] ? '' : `<p class="hint">Si no está al día: abre <a href="${f.url}">${f.nombre}</a> en este móvil y vuelve aquí. En otro dispositivo se usa el último dato guardado.</p>`}
    <div class="actions"><button class="btn primary" id="fRead">Leer ahora</button></div>`);
  $('#fRead', body).onclick = async () => { await leerFuentes(); closeSheet(); renderAll(); toast(leido[src] ? 'Datos de ' + f.nombre + ' al día' : 'No se encuentran datos de ' + f.nombre + ' en este móvil', !leido[src]); };
}

/* ---------- Pestaña: Deudas y cobros ---------- */
function headDeudas() {
  const i = inputs(), d = derive(i);
  return `<div class="kpis">
    <div class="k hero"><span>Deuda total</span><strong>${eur0(d.deuda)}</strong></div>
    <div class="k"><span>Préstamos</span><strong>${eur0(i.prestamos)}</strong></div>
    <div class="k"><span>Deudas pendientes</span><strong>${eur0(i.deudas)}</strong></div>
    <div class="k line"><span>Cobros pendientes</span><strong>${eur0(i.cobros)}</strong></div>
  </div>`;
}
function itemRow(x, tipo) {
  const pg = pagado(x), pe = pendiente(x);
  return `<li><button class="row" data-id="${x.id}" data-t="${tipo}" style="display:block">
    <div style="display:flex;justify-content:space-between;gap:12px;align-items:baseline">
      <div class="main"><div class="name">${esc(x.quien)}</div><div class="sub">${esc(x.concepto || '')}${x.concepto ? ' · ' : ''}${fmtDate(x.fecha)}${pg > EPS ? ` · ${tipo === 'deudas' ? 'pagado' : 'cobrado'} ${eur0(pg)} de ${eur0(x.importe)}` : ''}</div></div>
      <div class="side"><div class="big">${eur(pe > EPS ? pe : x.importe)}</div></div>
    </div>${pg > EPS && pe > EPS ? `<div class="prog"><i style="width:${Math.min(100, pg / x.importe * 100).toFixed(1)}%"></i></div>` : ''}
  </button></li>`;
}
function listaItems(tipo) {
  const xs = state[tipo];
  const ab = xs.filter(abierta).sort((a, b) => a.fecha.localeCompare(b.fecha));
  const ce = xs.filter((x) => !abierta(x)).sort((a, b) => b.fecha.localeCompare(a.fecha));
  return `${ab.length ? `<ul class="list" data-list="${tipo}">${ab.map((x) => itemRow(x, tipo)).join('')}</ul>` : `<p class="empty" style="padding:18px">Nada pendiente</p>`}
    ${ce.length ? `<details class="more"><summary>${tipo === 'deudas' ? 'Pagadas' : 'Cobrados'} (${ce.length})</summary><ul class="list" data-list="${tipo}" style="margin-top:8px">${ce.slice(0, 50).map((x) => itemRow(x, tipo).replace('class="row"', 'class="row dim"')).join('')}</ul></details>` : ''}`;
}
function renderDeudas() {
  const i = inputs();
  let html = segBtns('dSeg', [['deudas', 'Deudas'], ['cobros', 'Cobros pendientes']], ui.dseg);
  if (ui.dseg === 'deudas') {
    html += `<div class="sec"><h3>Préstamos</h3><span>${eur(i.prestamos)}</span></div>
      <ul class="list" id="prest">${[...state.prestamos].sort(byName).map((p) => `<li><button class="row" data-id="${p.id}"><div class="main"><div class="name">${esc(p.nombre)}</div><div class="sub">Capital pendiente${p.fecha ? ' a ' + fmtDate(p.fecha) : ''}</div></div><div class="side"><div class="big">${eur(p.capital)}</div></div></button></li>`).join('')}
      <li><button class="row auto" data-id="" style="justify-content:center;color:var(--muted)">+ Añadir préstamo</button></li></ul>
      <div class="sec"><h3>Deudas pendientes</h3><span>${eur(i.deudas)}</span></div>
      ${listaItems('deudas')}`;
  } else {
    const ac = autoCobros();
    html += `<div class="sec"><h3>De tus apps</h3><span>${eur(i.x.cobAuto)}</span></div>
      <ul class="list" id="autoCob">${ac.map((x) => { const e = estadoFuente(x.src); return `<li><button class="row auto" data-src="${x.src}"><div class="main"><div class="name">${x.nombre}</div><div class="sub">${x.sub} · <span class="tag ${e.cls}">${e.txt}</span></div></div><div class="side"><div class="big">${eur(x.v)}</div></div></button></li>`; }).join('')}</ul>
      <div class="sec"><h3>Otros cobros</h3><span>${eur(i.x.cobMan)}</span></div>
      ${listaItems('cobros')}`;
  }
  $('#view').innerHTML = html;
  bindSeg($('#view'), 'dSeg', (v) => { ui.dseg = v; renderAll(); });
  $('#view').querySelectorAll('[data-list]').forEach((ul) => { ul.onclick = (e) => { const b = e.target.closest('[data-id]'); if (b) openItem(b.dataset.t, b.dataset.id); }; });
  const pr = $('#prest'); if (pr) pr.onclick = (e) => { const b = e.target.closest('[data-id]'); if (b) openPrestamo(b.dataset.id || null); };
  const au = $('#autoCob'); if (au) au.onclick = (e) => { const b = e.target.closest('[data-src]'); if (b) openFuente(b.dataset.src); };
}
function openPrestamo(id) {
  const isNew = !id;
  const p = isNew ? { nombre: '', capital: '' } : state.prestamos.find((x) => x.id === id);
  const body = openSheet(isNew ? 'Préstamo nuevo' : p.nombre, `
    ${field('pC', 'Capital pendiente hoy €', fmtIn(p.capital))}
    <label class="fld"><span>Nombre</span><input id="pN" type="text" value="${esc(p.nombre)}" autocomplete="off" placeholder="Ej.: Préstamo Caixa"></label>
    <p class="hint">Lo actualizas tú cuando pagas cuotas o amortizas.</p>
    <div class="actions"><button class="btn primary" id="pSave">Guardar</button></div>
    ${isNew ? '' : '<div class="sep"></div><div class="actions"><button class="btn danger" id="pDel">Borrar préstamo</button></div>'}`);
  const save = () => {
    const nombre = $('#pN', body).value.trim(), cap = parseNum($('#pC', body).value);
    if (!nombre) return toast('Pon un nombre', true);
    if (!(cap >= 0)) return toast('Escribe el capital pendiente', true);
    if (commit((s) => { if (isNew) s.prestamos.push({ id: uid(), nombre, capital: r2(cap), fecha: todayISO() }); else Object.assign(s.prestamos.find((z) => z.id === id), { nombre, capital: r2(cap), fecha: todayISO() }); })) { closeSheet(); toast('Préstamo guardado'); }
  };
  $('#pSave', body).onclick = save;
  const del = $('#pDel', body);
  if (del) armDelete(del, () => { if (commit((s) => { s.prestamos = s.prestamos.filter((z) => z.id !== id); })) { closeSheet(); toast('Borrado'); } }, 'Borrar préstamo');
  setTimeout(() => { const i = $('#pC', body); i.focus(); i.select(); }, 60);
}
function openItem(tipo, id) {
  const isNew = !id, esD = tipo === 'deudas';
  const x = isNew ? { quien: '', concepto: '', importe: '', fecha: todayISO(), pagos: [], nota: '' } : state[tipo].find((z) => z.id === id);
  if (!x) return;
  const pg = pagado(x), pe = isNew ? 0 : pendiente(x);
  const pagos = [...(x.pagos || [])].sort((a, b) => b.fecha.localeCompare(a.fecha));
  const body = openSheet(isNew ? (esD ? 'Deuda nueva' : 'Cobro nuevo') : x.quien, `
    ${isNew ? '' : `<div class="box"><div><span>Importe</span><b>${eur(x.importe)}</b></div><div><span>${esD ? 'Pagado' : 'Cobrado'}</span><b>${eur(pg)}</b></div><div class="tot"><span>Pendiente</span><b>${eur(pe)}</b></div></div>
    ${pe > EPS ? `<div class="actions"><button class="btn primary" id="iPay">${esD ? 'Registrar pago' : 'Registrar cobro'}</button></div>` : ''}
    ${pagos.length ? `<h3>${esD ? 'Pagos' : 'Cobros'}</h3><ul class="list" id="iPagos">${pagos.map((p) => `<li><button class="row" data-p="${p.id}"><div class="main"><div class="name">${eur(p.imp)}</div><div class="sub">${fmtDate(p.fecha)}</div></div><div class="side small">Quitar</div></button></li>`).join('')}</ul>` : ''}
    <h3>Datos</h3>`}
    <label class="fld"><span>${esD ? 'A quién se le debe' : 'Quién te debe'}</span><input id="iQ" type="text" value="${esc(x.quien)}" autocomplete="off"></label>
    <label class="fld"><span>Concepto</span><input id="iC" type="text" value="${esc(x.concepto || '')}" autocomplete="off"></label>
    <div class="grid2">${field('iI', 'Importe total €', fmtIn(x.importe))}<label class="fld"><span>Fecha</span><input id="iF" type="date" value="${esc(x.fecha)}"></label></div>
    <label class="fld"><span>Nota</span><input id="iN" type="text" value="${esc(x.nota || '')}" autocomplete="off"></label>
    <div class="actions"><button class="btn ${isNew ? 'primary' : ''}" id="iSave">Guardar</button></div>
    ${isNew ? '' : '<div class="sep"></div><div class="actions"><button class="btn danger" id="iDel">Borrar</button></div>'}`);
  $('#iSave', body).onclick = () => {
    const quien = $('#iQ', body).value.trim(), imp = parseNum($('#iI', body).value), fecha = $('#iF', body).value || todayISO();
    if (!quien) return toast(esD ? 'Indica a quién se le debe' : 'Indica quién te debe', true);
    if (!(imp > 0)) return toast('Escribe el importe', true);
    if (!isNew && imp < pg - EPS) return toast('El importe no puede ser menor que lo ya ' + (esD ? 'pagado' : 'cobrado'), true);
    const data = { quien, concepto: $('#iC', body).value.trim(), importe: r2(imp), fecha, nota: $('#iN', body).value.trim() };
    if (commit((s) => { if (isNew) s[tipo].push({ id: uid(), ...data, pagos: [], ts: Date.now() }); else Object.assign(s[tipo].find((z) => z.id === id), data); })) { closeSheet(); toast('Guardado'); }
  };
  if (!isNew) {
    const pay = $('#iPay', body); if (pay) pay.onclick = () => reopen(() => openPago(tipo, id));
    const pl = $('#iPagos', body);
    if (pl) pl.onclick = (e) => { const b = e.target.closest('[data-p]'); if (!b) return; if (b.dataset.armed) { commit((s) => { const z = s[tipo].find((q) => q.id === id); z.pagos = z.pagos.filter((p) => p.id !== b.dataset.p); }); reopen(() => openItem(tipo, id)); } else { b.dataset.armed = '1'; b.querySelector('.side').textContent = 'Pulsa otra vez'; setTimeout(() => { delete b.dataset.armed; const sd = b.querySelector('.side'); if (sd) sd.textContent = 'Quitar'; }, 3000); } };
    armDelete($('#iDel', body), () => { if (commit((s) => { s[tipo] = s[tipo].filter((z) => z.id !== id); })) { closeSheet(); toast('Borrado'); } });
  }
}
function openPago(tipo, id) {
  const x = state[tipo].find((z) => z.id === id), esD = tipo === 'deudas', pe = pendiente(x);
  const body = openSheet(esD ? 'Registrar pago' : 'Registrar cobro', `
    <div class="box"><div><span>${esc(x.quien)}${x.concepto ? ' · ' + esc(x.concepto) : ''}</span><b>pendiente ${eur(pe)}</b></div></div>
    <div class="grid2">${field('yI', 'Importe €', fmtIn(pe))}<label class="fld"><span>Fecha</span><input id="yF" type="date" value="${todayISO()}"></label></div>
    <div class="actions"><button class="btn primary" id="ySave">Guardar</button></div>`);
  const i = $('#yI', body);
  $('#ySave', body).onclick = () => {
    const n = parseNum(i.value); if (!(n > 0)) return toast('Escribe el importe', true);
    if (n > pe + EPS) return toast('Es más de lo pendiente (' + eur(pe) + ')', true);
    if (commit((s) => { s[tipo].find((z) => z.id === id).pagos.push({ id: uid(), fecha: $('#yF', body).value || todayISO(), imp: r2(n) }); })) { closeSheet(); toast(n >= pe - EPS ? (esD ? 'Deuda saldada' : 'Cobro completado') : 'Pago parcial anotado'); }
  };
  setTimeout(() => { i.focus(); i.select(); }, 60);
}

/* ---------- Pestaña: Presupuesto (antes artefacto «Presupuesto personal») ---------- */
const PTIPOS = ['Fijo', 'Variable', 'Colchón', 'Ahorro'];
const P = () => state.presu;
const pcat = (id) => P().cfg.cats.find((c) => c.id === id);
const pcatName = (id) => (pcat(id) ? pcat(id).name : 'Sin categoría');
function pcalc() {
  const m = ui.pmes, cfg = P().cfg;
  const from = m < cfg.start ? m : cfg.start;
  const keys = ui.pmodo === 'mes' ? [m] : rangeM(from, m);
  const n = keys.length;
  const items = keys.flatMap((k) => P().months[k] || []);
  const rows = cfg.cats.map((c) => ({ c, spent: sum(items.filter((x) => x.c === c.id), (x) => x.a), budget: c.amt * n }));
  const known = new Set(cfg.cats.map((c) => c.id));
  const orphan = sum(items.filter((x) => !known.has(x.c)), (x) => x.a);
  if (orphan > 0) rows.push({ c: { id: '__none', name: 'Sin categoría', type: 'Variable', amt: 0 }, spent: orphan, budget: 0 });
  return { rows, spent: sum(rows, (r) => r.spent), budget: sum(rows, (r) => r.budget), n };
}
const ptone = (r) => { if (r.c.type === 'Ahorro') return ''; const p = r.budget ? r.spent / r.budget : (r.spent > 0 ? 2 : 0); if (p > 1) return 'bad'; if (p >= 0.9 && r.c.type !== 'Fijo') return 'warn'; return ''; };
function headPresu() {
  const k = monthOf(todayISO()), cats = P().cfg.cats;
  const items = P().months[k] || [];
  const spent = sum(items, (x) => x.a), budget = sum(cats, (c) => c.amt);
  const pct = budget ? Math.min(1, spent / budget) : 0;
  return `<div class="kpis">
    <div class="k hero"><span>Gastado en ${MESES[Number(k.slice(5, 7)) - 1]}</span><strong>${eur0(spent)}</strong></div>
    <div class="hbar"><i style="width:${(pct * 100).toFixed(1)}%"></i></div>
    <div class="k"><span>Presupuesto</span><strong>${eur0(budget)}</strong></div>
    <div class="k"><span>${spent > budget ? 'Te pasas' : 'Quedan'}</span><strong>${eur0(Math.abs(budget - spent))}</strong></div>
  </div>`;
}
function monthNav() { return `<div class="monthnav"><button id="pPrev" aria-label="Mes anterior">‹</button><b>${monthLabel(ui.pmes)}</b><button id="pNext" aria-label="Mes siguiente">›</button></div>`; }
function renderPresu() {
  let html = segBtns('pSeg', [['resumen', 'Resumen'], ['movs', 'Gastos'], ['ajustes', 'Presupuesto']], ui.pseg);
  if (ui.pseg === 'resumen') {
    const { rows, spent, budget, n } = pcalc();
    const left = budget - spent;
    const cls = spent > budget ? 'bad' : (budget && spent / budget >= 0.9 ? 'warn' : '');
    let pace = '';
    if (ui.pmodo === 'mes' && ui.pmes === monthOf(todayISO())) {
      const vr = rows.filter((r) => r.c.type === 'Variable'), vb = sum(vr, (r) => r.budget), vs = sum(vr, (r) => r.spent);
      if (vb > 0) { const d = new Date(), dim = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate(); const dp = Math.round(d.getDate() / dim * 100), sp = Math.round(vs / vb * 100); pace = `<p class="small" style="margin:8px 0 0;color:${sp > dp + 10 ? 'var(--amber)' : 'var(--ok)'}">Gastos variables: llevas el ${sp} % y va el ${dp} % del mes.</p>`; }
    }
    html += monthNav() + segBtns('pModo', [['mes', 'Mes'], ['acum', 'Acumulado']], ui.pmodo) + `
      <div class="box" style="display:block"><div class="small">${ui.pmodo === 'mes' ? 'Gastado este mes' : `Gastado en el acumulado (${n} ${n === 1 ? 'mes' : 'meses'})`}</div>
        <div style="font-size:26px;font-weight:700;font-variant-numeric:tabular-nums">${eur(spent)}</div>
        <div class="bar ${cls}"><i style="width:${budget ? Math.min(100, spent / budget * 100).toFixed(1) : 0}%"></i></div>
        <div style="display:flex;justify-content:space-between;margin-top:6px"><span class="small">Presupuesto ${eur(budget)}</span><b class="${left < 0 ? 'neg' : ''}" style="font-size:14px">${left < 0 ? 'Te pasas ' + eur(-left) : 'Quedan ' + eur(left)}</b></div>${pace}</div>
      <ul class="list" id="pCats">${rows.map((r) => {
        const p = r.budget ? Math.min(100, r.spent / r.budget * 100) : (r.spent > 0 ? 100 : 0);
        let stx, sc = '';
        if (r.c.type === 'Ahorro') { stx = r.spent >= r.budget ? 'Cumplido ✓' : 'Faltan ' + eur(r.budget - r.spent); sc = r.spent >= r.budget ? 'color:var(--ok)' : ''; }
        else if (r.spent > r.budget) { stx = 'Te pasas ' + eur(r.spent - r.budget); sc = 'color:var(--danger);font-weight:600'; }
        else stx = 'Quedan ' + eur(r.budget - r.spent);
        return `<li><button class="row" data-c="${esc(r.c.id)}" style="display:block"><div style="display:flex;justify-content:space-between;gap:10px"><span class="name">${esc(r.c.name)}</span><span style="font-variant-numeric:tabular-nums;white-space:nowrap">${eur(r.spent)} <span class="small">/ ${eur(r.budget)}</span></span></div><div class="bar ${ptone(r)}"><i style="width:${p.toFixed(1)}%"></i></div><div class="sub" style="${sc}">${stx}</div></button></li>`;
      }).join('')}</ul>
      <p class="hint" style="margin-top:10px">Toca una categoría para ver sus gastos. Los pagos anuales (como el seguro) cuadran mejor en el acumulado.</p>`;
  } else if (ui.pseg === 'movs') {
    let list = [...(P().months[ui.pmes] || [])];
    if (ui.pfiltro) list = list.filter((x) => x.c === ui.pfiltro);
    list.sort((a, b) => b.d.localeCompare(a.d));
    html += monthNav() + (ui.pfiltro ? `<div class="chips"><button class="chip" aria-pressed="true" id="pClr">${esc(pcatName(ui.pfiltro))} ✕</button></div>` : '') +
      `<div class="summary"><span>${list.length} ${list.length === 1 ? 'gasto' : 'gastos'}</span><span>Total <b>${eur(sum(list, (x) => x.a))}</b></span></div>`;
    if (!list.length) html += '<p class="empty">Aún no hay gastos en este mes.<br>Pulsa + para apuntar el primero.</p>';
    let last = '';
    html += '<div id="pMovs">';
    list.forEach((x) => {
      if (x.d !== last) { html += `<div class="day">${new Date(x.d + 'T12:00').toLocaleDateString('es-ES', { weekday: 'short', day: 'numeric', month: 'short' })}</div>`; last = x.d; }
      html += `<button class="row" data-id="${x.id}" style="margin-bottom:6px"><div class="main"><div class="name">${esc(x.n || pcatName(x.c))}</div><div class="sub">${x.n ? esc(pcatName(x.c)) : '&nbsp;'}</div></div><div class="side"><div class="big">${eur(x.a)}</div></div></button>`;
    });
    html += '</div>';
  } else {
    const cfg = P().cfg, tot = sum(cfg.cats, (c) => c.amt), diff = cfg.nomina - tot;
    html += `<div class="box" style="display:block">${field('pNom', 'Nómina mensual objetivo €', fmtIn(cfg.nomina))}<div><span>Total presupuestado</span><b>${eur(tot)}</b></div><div><span>Nómina − total</span><b class="${diff < 0 ? 'neg' : ''}">${eur(diff)}</b></div></div>
      <h3>Categorías y presupuesto mensual</h3>
      <div id="pCfg">${cfg.cats.map((c) => `<div class="crow">
        <input type="text" data-k="name" data-id="${esc(c.id)}" value="${esc(c.name)}" aria-label="Nombre">
        <input type="text" inputmode="decimal" data-k="amt" data-id="${esc(c.id)}" value="${esc(fmtIn(c.amt))}" aria-label="Importe mensual">
        <button class="xdel" data-del="${esc(c.id)}" aria-label="Eliminar categoría">✕</button>
        <select class="tp" data-k="type" data-id="${esc(c.id)}" aria-label="Tipo">${PTIPOS.map((t) => `<option ${t === c.type ? 'selected' : ''}>${t}</option>`).join('')}</select></div>`).join('')}</div>
      <div class="actions" style="margin-top:10px"><button class="btn" id="pAdd">+ Añadir categoría</button></div>
      <label class="fld" style="margin-top:14px"><span>El acumulado empieza en</span><input type="month" id="pStart" value="${esc(cfg.start)}" class="inp"></label>
      <p class="hint" style="margin-top:0">El presupuesto va por libre: los gastos no cambian el saldo de las cuentas.</p>`;
  }
  $('#view').innerHTML = html;
  bindSeg($('#view'), 'pSeg', (v) => { ui.pseg = v; if (v !== 'movs') ui.pfiltro = null; renderAll(); });
  const pv = $('#pPrev'); if (pv) { pv.onclick = () => { ui.pmes = addM(ui.pmes, -1); renderAll(); }; $('#pNext').onclick = () => { ui.pmes = addM(ui.pmes, 1); renderAll(); }; }
  const pm = $('#pModo'); if (pm) bindSeg($('#view'), 'pModo', (v) => { ui.pmodo = v; renderAll(); });
  const pc = $('#pCats'); if (pc) pc.onclick = (e) => { const b = e.target.closest('[data-c]'); if (b) { ui.pfiltro = b.dataset.c === '__none' ? null : b.dataset.c; ui.pseg = 'movs'; renderAll(); window.scrollTo(0, 0); } };
  const clr = $('#pClr'); if (clr) clr.onclick = () => { ui.pfiltro = null; renderAll(); };
  const mv = $('#pMovs'); if (mv) mv.onclick = (e) => { const b = e.target.closest('[data-id]'); if (b) openGasto(b.dataset.id); };
  if (ui.pseg === 'ajustes') bindPresuCfg();
}
function bindPresuCfg() {
  const v = $('#view');
  $('#pNom', v).onchange = (e) => { const n = parseNum(e.target.value); if (isFinite(n)) commit((s) => { s.presu.cfg.nomina = r2(n); }); };
  $('#pStart', v).onchange = (e) => { if (/^\d{4}-\d{2}$/.test(e.target.value)) commit((s) => { s.presu.cfg.start = e.target.value; }); };
  $('#pCfg', v).onchange = (e) => {
    const t = e.target, k = t.dataset.k; if (!k) return;
    commit((s) => {
      const c = s.presu.cfg.cats.find((x) => x.id === t.dataset.id); if (!c) return;
      if (k === 'name') c.name = t.value.trim() || c.name;
      else if (k === 'amt') { const n = parseNum(t.value); if (isFinite(n) && n >= 0) c.amt = r2(n); }
      else if (k === 'type') c.type = t.value;
    });
  };
  $('#pCfg', v).onclick = (e) => {
    const b = e.target.closest('[data-del]'); if (!b) return;
    const id = b.dataset.del;
    if (Object.values(P().months).some((a) => a.some((x) => x.c === id))) return toast('Tiene gastos apuntados: no se puede borrar', true);
    if (!b.dataset.armed) { b.dataset.armed = '1'; b.textContent = '✓?'; toast('Pulsa otra vez para borrar'); setTimeout(() => { delete b.dataset.armed; b.textContent = '✕'; }, 3000); return; }
    commit((s) => { s.presu.cfg.cats = s.presu.cfg.cats.filter((c) => c.id !== id); });
  };
  $('#pAdd', v).onclick = () => commit((s) => { s.presu.cfg.cats.push({ id: 'c' + uid().slice(0, 8), name: 'Nueva categoría', type: 'Variable', amt: 0 }); });
}
function openGasto(id) {
  let g = null;
  if (id) for (const k in P().months) { const f = P().months[k].find((x) => x.id === id); if (f) { g = f; break; } }
  const isNew = !g;
  const hoy = todayISO();
  const x = g || { d: ui.pmes === monthOf(hoy) ? hoy : ui.pmes + '-01', c: ui.pfiltro && pcat(ui.pfiltro) ? ui.pfiltro : null, a: '', n: '' };
  let cat = x.c;
  if (!P().cfg.cats.length) return toast('Primero crea las categorías en «Presupuesto»', true);
  const body = openSheet(isNew ? 'Nuevo gasto' : 'Editar gasto', `
    ${field('gA', 'Importe €', isNew ? '' : fmtIn(x.a))}
    <div class="chips wrap" id="gC" style="margin-bottom:14px">${P().cfg.cats.map((c) => `<button type="button" class="chip" data-c="${esc(c.id)}" aria-pressed="${c.id === cat}">${esc(c.name)}</button>`).join('')}</div>
    <div class="grid2"><label class="fld"><span>Fecha</span><input id="gD" type="date" value="${esc(x.d)}"></label><label class="fld"><span>Concepto (opcional)</span><input id="gN" type="text" value="${esc(x.n || '')}" autocomplete="off"></label></div>
    <div class="actions"><button class="btn primary" id="gSave">Guardar</button></div>
    ${isNew ? '' : '<div class="sep"></div><div class="actions"><button class="btn danger" id="gDel">Eliminar</button></div>'}`);
  $('#gC', body).onclick = (e) => { const b = e.target.closest('[data-c]'); if (!b) return; cat = b.dataset.c; $('#gC', body).querySelectorAll('.chip').forEach((c) => c.setAttribute('aria-pressed', c === b)); };
  $('#gSave', body).onclick = () => {
    const a = parseNum($('#gA', body).value);
    if (!(a > 0)) return toast('Pon un importe', true);
    if (!cat) return toast('Elige una categoría', true);
    const d = $('#gD', body).value || hoy;
    const m = { id: g ? g.id : uid(), d, c: cat, a: r2(a), n: $('#gN', body).value.trim() };
    if (commit((s) => { const ms = s.presu.months; for (const k in ms) ms[k] = ms[k].filter((z) => z.id !== m.id); const nk = monthOf(d); ms[nk] = (ms[nk] || []).concat([m]); })) { ui.pmes = monthOf(d); closeSheet(); toast('Guardado'); }
  };
  const del = $('#gDel', body);
  if (del) armDelete(del, () => { if (commit((s) => { const ms = s.presu.months; for (const k in ms) ms[k] = ms[k].filter((z) => z.id !== g.id); })) { closeSheet(); toast('Eliminado'); } }, 'Eliminar');
  setTimeout(() => { if (isNew) $('#gA', body).focus(); }, 60);
}

/* ---------- Copia automática en Google Sheets ----------
   La app envía el estado completo al receptor (Apps Script). La hoja es un espejo de un solo sentido. */
let sync = { url: '', clave: '', enabled: false, dirty: false, lastOk: null, lastError: '' };
let syncRev = 0, syncTimer = null, syncRetry = null, syncRunning = false;
const syncOn = () => !!(sync.enabled && sync.url && sync.clave);
async function loadSync() { try { const s = await dbGet('sync'); if (s) sync = { ...sync, ...s }; } catch (_) { /* nada */ } }
function saveSync() { dbSet('sync', sync).catch(() => {}); }
function markDirty() {
  if (!syncOn()) return;
  sync.dirty = true; syncRev++; saveSync(); renderSyncStatus();
  clearTimeout(syncTimer); syncTimer = setTimeout(() => runSync(), 2500);
}
async function callScript(accion, extra = {}) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), 30000);
  try {
    const res = await fetch(sync.url, { method: 'POST', body: JSON.stringify({ accion, clave: sync.clave, ...extra }), signal: ctrl.signal });
    const txt = await res.text();
    try { return JSON.parse(txt); } catch (_) { throw new Error('La dirección no responde como el receptor (revisa que esté publicado para «Cualquier usuario»)'); }
  } catch (e) {
    if (e && e.name === 'AbortError') throw new Error('Tiempo de espera agotado');
    if (e instanceof TypeError) throw new Error('Sin conexión con Google');
    throw e;
  } finally { clearTimeout(t); }
}
function syncPayload() {
  const EUR = '#,##0.00', R = '0.00', PC = '0.0%';
  const a = actual();
  const res = [
    ['Fecha', todayISO()], ['Total bancos', a.bancos], ['Total efectivo', a.efectivo], ['Total stock', a.stock], ['Valor maquinaria', a.maquinaria], ['Otros activos', a.otros],
    ['Total activos', a.activos], ['Préstamos', a.prestamos], ['Deudas pendientes', a.deudas], ['Cobros pendientes', a.cobros], ['Total deuda', a.deuda],
    ['Capital neto', a.neto], ['Liquidez total', a.liquidez], ['Ratio liquidez/deuda', a.ratio == null ? '' : round(a.ratio, 4)], ['Endeudamiento', a.endeud == null ? '' : round(a.endeud, 4)],
    ['Índice de independencia financiera', a.indep == null ? '' : round(a.indep, 4)], ['% Stock', a.pstock == null ? '' : round(a.pstock, 4)], ['Balance de situación', a.balance],
  ];
  const hist = [];
  cierresAsc().forEach((c) => {
    const d = derive(c.i), p = cierreAnterior(c.mes), dp = p ? derive(p.i) : null;
    const dv = (k) => (dp ? r2(d[k] - dp[k]) : ''), dpc = (k) => (dp && Math.abs(dp[k]) > EPS ? round((d[k] - dp[k]) / dp[k], 4) : '');
    const nn = (v) => (v == null ? '' : round(v, 4));
    hist.push([c.fecha, d.bancos, d.efectivo, d.stock, d.maquinaria, d.otros, d.activos, dv('activos'), dpc('activos'), d.prestamos, d.deudas, d.cobros, d.deuda, dv('deuda'), d.neto, dv('neto'), dpc('neto'), d.liquidez, nn(d.ratio), nn(d.endeud), nn(d.indep), nn(d.pstock), d.balance, c.comentario || '', c.objetivo || '', { excel: 'Excel', auto: 'Automático', manual: 'A mano' }[c.origen] || '']);
  });
  const ctas = [...state.cuentas].sort(byName).map((c) => [c.nombre, GRUPOS[c.grupo], TIPOS[c.tipo], c.saldo, c.fecha || '']);
  const bienes = [...state.bienes].sort(byName).map((b) => [b.nombre, b.maquinaria ? 'Maquinaria' : 'Otros activos', b.valor, b.fecha || '']);
  const prest = [...state.prestamos].sort(byName).map((p) => [p.nombre, p.capital, p.fecha || '']);
  const items = (xs) => [...xs].sort((x, y) => x.fecha.localeCompare(y.fecha)).map((x) => [x.fecha, x.quien, x.concepto || '', x.importe, pagado(x), pendiente(x), abierta(x) ? 'Pendiente' : 'Saldada', x.nota || '']);
  const cobAuto = autoCobros().map((x) => [todayISO(), x.nombre + ' (automático)', x.sub, x.v, 0, x.v, 'Pendiente', '']);
  const gastos = [];
  Object.keys(P().months).sort().forEach((k) => [...P().months[k]].sort((x, y) => x.d.localeCompare(y.d)).forEach((x) => gastos.push([x.d, x.n || '', pcatName(x.c), x.a])));
  const tablas = [
    { nombre: 'Resumen', cabecera: ['Concepto', 'Valor'], filas: res, formatos: [null, EUR], fechas: [], formatosFila: { 0: 'dd/mm/yyyy', 13: R, 14: R, 15: R, 16: PC } },
    { nombre: 'Histórico', cabecera: ['Mes (cierre)', 'Total bancos', 'Total efectivo', 'Total stock', 'Valor maquinaria', 'Otros activos', 'Total activos', 'Variación absoluta', 'Variación activos %', 'Préstamos', 'Deudas pendientes', 'Cobros pendientes', 'Total deuda', 'Variación deuda', 'Capital neto', 'Variación capital neto', '% Variación C. Neto', 'Liquidez total', 'Ratio liquidez/deuda', 'Endeudamiento', 'Índice indep. financiera', '% Stock', 'Balance situación', 'Comentario', 'Objetivo', 'Cierre'], filas: hist, formatos: ['dd/mm/yyyy', EUR, EUR, EUR, EUR, EUR, EUR, EUR, PC, EUR, EUR, EUR, EUR, EUR, EUR, EUR, PC, EUR, R, R, R, PC, EUR, null, null, null], fechas: [0] },
    { nombre: 'Cuentas', cabecera: ['Cuenta', 'Grupo', 'Tipo', 'Saldo (€)', 'Actualizado'], filas: ctas, formatos: [null, null, null, EUR, 'dd/mm/yyyy'], fechas: [4] },
    { nombre: 'Bienes', cabecera: ['Bien', 'Cuenta como', 'Valor (€)', 'Actualizado'], filas: bienes, formatos: [null, null, EUR, 'dd/mm/yyyy'], fechas: [3] },
    { nombre: 'Préstamos', cabecera: ['Préstamo', 'Capital pendiente (€)', 'Actualizado'], filas: prest, formatos: [null, EUR, 'dd/mm/yyyy'], fechas: [2] },
    { nombre: 'Deudas', cabecera: ['Fecha', 'A quién', 'Concepto', 'Importe (€)', 'Pagado (€)', 'Pendiente (€)', 'Estado', 'Nota'], filas: items(state.deudas), formatos: ['dd/mm/yyyy', null, null, EUR, EUR, EUR, null, null], fechas: [0] },
    { nombre: 'Cobros', cabecera: ['Fecha', 'Quién', 'Concepto', 'Importe (€)', 'Cobrado (€)', 'Pendiente (€)', 'Estado', 'Nota'], filas: cobAuto.concat(items(state.cobros)), formatos: ['dd/mm/yyyy', null, null, EUR, EUR, EUR, null, null], fechas: [0] },
    { nombre: 'Gastos', cabecera: ['Fecha', 'Concepto', 'Categoría', 'Importe (€)'], filas: gastos, formatos: ['dd/mm/yyyy', null, null, EUR], fechas: [0] },
  ];
  return { tablas, estado: JSON.stringify(state), version: APP_VERSION };
}
async function runSync() {
  if (!syncOn() || syncRunning) return;
  if (!navigator.onLine) { renderSyncStatus(); return; }
  syncRunning = true; const rev = syncRev; renderSyncStatus();
  try {
    const r = await callScript('sincronizar', syncPayload());
    if (!r.ok) throw new Error(r.error === 'clave' ? 'La clave no coincide con la del script' : (r.error || 'Error del receptor'));
    sync.lastOk = new Date().toISOString(); sync.lastError = '';
    if (syncRev === rev) sync.dirty = false;
  } catch (e) {
    sync.lastError = e.message || String(e); sync.dirty = true;
    clearTimeout(syncRetry); syncRetry = setTimeout(() => runSync(), 60000);
  } finally {
    syncRunning = false; saveSync(); renderSyncStatus();
    if (sync.dirty && !sync.lastError) { clearTimeout(syncTimer); syncTimer = setTimeout(() => runSync(), 800); }
  }
}
function syncStatusText() {
  if (!syncOn()) return '';
  if (syncRunning) return '☁ Enviando a Google Sheets…';
  if (sync.lastError) return '☁ No se pudo enviar: ' + sync.lastError;
  if (sync.dirty) return navigator.onLine ? '☁ Pendiente de enviar…' : '☁ Pendiente de enviar (sin conexión)';
  return '☁ Hoja al día · ' + ago(sync.lastOk);
}
function renderSyncStatus() {
  const el = $('#syncStatus'); if (el) { el.textContent = syncStatusText(); el.hidden = !syncOn(); }
  const box = $('#syncLine'); if (box) box.textContent = syncStatusText();
}
function drawSync(box) {
  if (syncOn()) {
    box.innerHTML = `
      <p class="small" id="syncLine">${esc(syncStatusText())}</p>
      <p class="small">Cada cambio se envía solo a tu hoja de Google (Resumen, Histórico, Cuentas, Bienes, Préstamos, Deudas, Cobros y Gastos). Sin conexión se queda pendiente y se envía al volver.</p>
      <div class="actions">
        <button class="btn primary" id="syNow">Sincronizar ahora</button>
        <button class="btn" id="syRes">Restaurar desde la hoja</button>
        <button class="btn danger" id="syOff">Desactivar</button>
      </div>`;
    $('#syNow', box).onclick = () => { if (!navigator.onLine) return toast('Sin conexión: se enviará al volver', true); sync.lastError = ''; markDirty(); clearTimeout(syncTimer); runSync(); };
    $('#syRes', box).onclick = async () => {
      try {
        const est = await callScript('estado');
        if (!est.ok || !est.estado) throw new Error(est.error === 'sin copia' ? 'La hoja todavía no tiene copia' : (est.error || 'No se pudo leer la copia'));
        const s = JSON.parse(est.estado);
        if (!validState(s)) throw new Error('La copia de la hoja no es válida');
        if (!confirm(`Se reemplazarán los datos de este móvil por los de la hoja (${stateInfo(s)}, copia ${ago(est.fecha)}). ¿Continuar?`)) return;
        state = s; await persist(); closeSheet(); renderAll(); toast('Restaurado desde la hoja');
      } catch (e) { toast(e.message || 'No se pudo restaurar', true); }
    };
    $('#syOff', box).onclick = () => { sync.enabled = false; sync.dirty = false; sync.lastError = ''; saveSync(); renderSyncStatus(); drawSync(box); toast('Copia automática desactivada'); };
  } else {
    box.innerHTML = `
      <p class="small">Envía automáticamente todo a una hoja de Google Sheets de tu Drive (copia de seguridad y datos para consultar). Se configura una vez con el script que se explica en el README.</p>
      <label class="fld"><span>Dirección del script (termina en /exec)</span><input id="syU" type="text" inputmode="url" value="${esc(sync.url)}" placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off"></label>
      <label class="fld"><span>Clave</span><input id="syK" type="password" value="${esc(sync.clave)}" autocomplete="off"></label>
      <div class="actions"><button class="btn primary" id="syOn">Probar y activar</button></div>`;
    $('#syOn', box).onclick = () => activateSync($('#syU', box).value.trim(), $('#syK', box).value.trim(), box);
  }
}
async function activateSync(url, clave, box) {
  url = url.replace(/\s+/g, '').replace(/[?#].*$/, '').replace(/\/dev\/?$/, '/exec');
  if (!/^https:\/\/script\.google\.com\/(a\/macros\/[^/]+|macros)\/s\/[^/\s]+\/exec\/?$/.test(url)) return toast('La dirección debe ser la de «Aplicación web» del script (empieza por https://script.google.com/ y termina en /exec)', true);
  clave = clave.trim();
  if (!clave) return toast('Escribe la clave', true);
  const prev = { ...sync };
  sync.url = url; sync.clave = clave;
  const finish = () => { sync.enabled = true; sync.dirty = true; sync.lastError = ''; syncRev++; saveSync(); renderSyncStatus(); drawSync(box); runSync(); toast('Conectado con Google Sheets'); };
  try {
    const ping = await callScript('ping');
    if (!ping.ok) throw new Error(ping.error === 'clave' ? 'La clave no coincide con la del script' : (ping.error || 'El receptor no responde'));
    const est = await callScript('estado');
    if (est.ok && est.estado) {
      let info = null; try { info = JSON.parse(est.estado); } catch (_) { /* nada */ }
      const okInfo = info && validState(info);
      box.innerHTML = `
        <p><b>La hoja ya contiene una copia</b> (${okInfo ? stateInfo(info) : 'datos'}, ${ago(est.fecha)}).</p>
        <p class="small">Si este móvil es nuevo, usa la copia de la hoja. Si lo de este móvil es lo correcto, sobrescribe la hoja.</p>
        <div class="actions">
          <button class="btn primary" id="syUse" ${okInfo ? '' : 'disabled'}>Usar la copia de la hoja en este móvil</button>
          <button class="btn" id="syOver">Sobrescribir la hoja con los datos de este móvil</button>
          <button class="btn" id="syCan">Cancelar</button>
        </div>`;
      $('#syUse', box).onclick = async () => { state = info; await persist(); renderAll(); finish(); };
      $('#syOver', box).onclick = finish;
      $('#syCan', box).onclick = () => { sync = prev; drawSync(box); };
      return;
    }
    finish();
  } catch (e) { sync = prev; toast(e.message || 'No se pudo conectar', true); }
}

/* ---------- Ajustes y copia de seguridad ---------- */
function download(name, text, type) {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
async function shareOrDownloadBackup() {
  const name = `finanzas-copia-${todayISO()}.json`;
  const text = JSON.stringify({ app: 'finanzas', version: APP_VERSION, exportedAt: new Date().toISOString(), state }, null, 1);
  const file = new File([text], name, { type: 'application/json' });
  try {
    if (navigator.canShare && navigator.canShare({ files: [file] })) await navigator.share({ files: [file], title: 'Copia Finanzas' });
    else download(name, text, 'application/json');
  } catch (e) { if (e && e.name === 'AbortError') return; download(name, text, 'application/json'); }
  state.settings.lastBackup = new Date().toISOString(); persist(); toast('Copia generada');
}
async function restoreFromFile(input) {
  const file = input.files[0]; if (!file) return;
  try {
    const j = JSON.parse(await file.text());
    if (j && j.app === 'finanzas' && j.parcial === 'maquinaria' && Array.isArray(j.bienes)) {
      // Fichero parcial: sustituye solo la maquinaria, sin tocar el resto de datos
      const xs = j.bienes.filter((b) => b && b.nombre && isFinite(b.valor));
      const tot = sum(xs, (b) => b.valor);
      if (!confirm(`Se sustituirá la maquinaria actual (${eur(inputs().maquinaria)}) por ${xs.length} máquinas (${eur(tot)}). El resto de datos no cambia. ¿Continuar?`)) { input.value = ''; return; }
      if (commit((c) => { c.bienes = c.bienes.filter((b) => !b.maquinaria).concat(xs.map((b) => ({ id: uid(), nombre: String(b.nombre), valor: r2(b.valor), maquinaria: true, anio: b.anio || null, compra: b.compra != null ? r2(b.compra) : null, fecha: b.fecha || todayISO() }))); })) { closeSheet(); ui.tab = 'cuentas'; renderAll(); toast(`${xs.length} máquinas cargadas · ${eur0(tot)}`); }
      input.value = ''; return;
    }
    const s = j && j.state;
    if (!j || j.app !== 'finanzas' || !validState(s)) throw new Error('El fichero no es una copia de Finanzas');
    if (!isEmpty() && !confirm(`Se reemplazarán todos los datos actuales por los del fichero (${stateInfo(s)}). ¿Continuar?`)) { input.value = ''; return; }
    state = s; await persist(); closeSheet(); await leerFuentes(); autoCierre(); renderAll(); toast('Datos cargados: ' + stateInfo(s));
  } catch (err) { toast(err.message || 'No se pudo leer el fichero', true); }
  input.value = '';
}
function openSettings() {
  const fu = Object.keys(FUENTES).map((k) => { const e = estadoFuente(k); return `<div><span>${FUENTES[k].nombre}</span><b><span class="tag ${e.cls}">${e.txt}</span></b></div>`; }).join('');
  const body = openSheet('Ajustes', `
    <h3 style="margin-top:4px">Datos de tus otras apps</h3>
    <p class="small">Stock y cobros pendientes se leen solos de Ventas, Stock Fincas, AgroServicios y Asesor instaladas en este móvil.</p>
    <div class="box">${fu}</div>
    <button type="button" class="switch" id="stCom" role="switch" aria-checked="${!!state.settings.comisiones}"><span>Contar las comisiones pendientes de Ventas como cobros</span><i></i></button>
    <div class="actions"><button class="btn" id="stRead">Leer ahora</button></div>
    <div class="sep"></div>
    <h3>Copia automática en Google Sheets</h3>
    <div id="syncBox"></div>
    <div class="sep"></div>
    <h3>Copia de seguridad en fichero</h3>
    <p class="small">Los datos viven en este móvil. Guarda de vez en cuando un fichero con todo (por ejemplo en Drive).${state.settings.lastBackup ? ` Última copia: ${new Date(state.settings.lastBackup).toLocaleString('es-ES')}.` : ' Aún no has hecho ninguna.'}</p>
    <div class="actions">
      <button class="btn primary" id="sBk">Guardar copia / compartir</button>
      <label class="btn" style="display:grid;place-items:center;cursor:pointer">Cargar datos desde fichero<input id="sRs" type="file" accept="application/json,.json" hidden></label>
    </div>
    <div class="sep"></div>
    <p class="small">Finanzas v${APP_VERSION} · sustituye a Control_financiero.xlsx.</p>`);
  drawSync($('#syncBox', body));
  $('#stCom', body).onclick = (e) => { const v = !state.settings.comisiones; e.currentTarget.setAttribute('aria-checked', v); commit((s) => { s.settings.comisiones = v; }); };
  $('#stRead', body).onclick = async () => { await leerFuentes(); closeSheet(); renderAll(); const n = Object.keys(FUENTES).filter((k) => leido[k]).length; toast(`Leídas ${n} de 4 apps en este móvil`, n < 4); };
  $('#sBk', body).onclick = shareOrDownloadBackup;
  $('#sRs', body).onchange = (e) => restoreFromFile(e.target);
}

/* ---------- Render general ---------- */
const TABS = {
  inicio: { title: 'Finanzas', fab: null, head: headInicio, render: renderInicio },
  cuentas: { title: 'Cuentas', fab: 'Añadir', head: headCuentas, render: renderCuentas },
  deudas: { title: 'Deudas y cobros', fab: () => (ui.dseg === 'deudas' ? 'Nueva deuda' : 'Nuevo cobro'), head: headDeudas, render: renderDeudas },
  presu: { title: 'Presupuesto', fab: () => (ui.pseg === 'ajustes' ? null : 'Nuevo gasto'), head: headPresu, render: renderPresu },
};
function renderAll() {
  const t = TABS[ui.tab];
  $('#topTitle').textContent = t.title;
  $('#head').innerHTML = t.head && !isEmpty() ? t.head() : '';
  document.querySelectorAll('.tab').forEach((b) => { if (b.dataset.tab === ui.tab) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); });
  $('#inicioBadge').hidden = !state.cierres.some((c) => c.revisar);
  const fab = typeof t.fab === 'function' ? t.fab() : t.fab;
  $('#fab').hidden = !fab; if (fab) $('#fabLabel').textContent = fab;
  renderSyncStatus();
  t.render();
}

async function init() {
  document.querySelectorAll('.tab').forEach((b) => b.onclick = () => { ui.tab = b.dataset.tab; window.scrollTo(0, 0); renderAll(); });
  $('#btnSettings').onclick = openSettings;
  $('#fab').onclick = () => {
    if (ui.tab === 'cuentas') openAnadirCuentas();
    else if (ui.tab === 'deudas') openItem(ui.dseg, null);
    else if (ui.tab === 'presu') openGasto(null);
  };
  await loadSync();
  $('#syncStatus').onclick = openSettings;
  window.addEventListener('online', () => { if (sync.dirty) runSync(); else renderSyncStatus(); });
  window.addEventListener('offline', renderSyncStatus);
  document.addEventListener('visibilitychange', async () => {
    if (document.visibilityState !== 'visible' || !state) return;
    if (sync.dirty) runSync();
    await leerFuentes(); autoCierre();
    if (!sheetOpen) renderAll();
  });
  setInterval(renderSyncStatus, 60000);
  try {
    let s = await dbGet('state');
    if (!s) { s = emptyState(); await dbSet('state', s); }
    state = s;
  } catch (e) {
    state = emptyState();
    toast('Este navegador no permite guardar datos: lo que anotes se perderá al cerrar', true);
  }
  renderAll();
  await leerFuentes();
  autoCierre();
  renderAll();
  if (syncOn() && sync.dirty) setTimeout(runSync, 1500);
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
}
if (typeof window !== 'undefined' && !window.__FINANZAS_TEST__) init();
