#!/usr/bin/env node
/*
 * Previsión financiera mensual · FG Agro (versión 1.1, 27/10/2026)
 *
 * Cruza Finanzas con las cuentas de las fincas y el resto del negocio y saca la previsión de
 * caja a 12 meses, el colchón de 3 meses y los datos para las recomendaciones.
 *
 * Uso:  node prevision.js <carpeta-datos> [AAAA-MM-DD] > resultado.json
 *   <carpeta-datos> tiene un .json por rango, tal como lo devuelve Google Sheets (get_values):
 *     {"range": "...", "values": [[...], ...]}
 *   Nombres de fichero (ver LEEME): fin-resumen, fin-cuentas, fin-prestamos, fin-deudas, fin-cobros,
 *   fin-historico, fincas-tabla, fincas-presupuestos, fincas-parametros, asesor-clientes,
 *   ventas-ventas, ventas-comisiones, agro-trabajos.
 *   o bien los .xlsx exportados de Drive: fincas.xlsx (Stock Fincas), ventas.xlsx (ventas FG Agro), y si se quiere
 *   finanzas.xlsx, asesor.xlsx y agro.xlsx. Si hay .json, manda el .json.
 *   La fecha es la de la ejecución (por defecto, hoy). La previsión empieza en el mes de esa fecha.
 *
 * Reglas (decisiones de Francis, 27/10/2026):
 *  - Horizonte: 12 meses. Colchón: 3 meses de salidas (personal + pagos de fincas + cuotas).
 *  - Personal: presupuesto fijo de 2.000 €/mes (negocio y personal no se separan).
 *  - Fincas, solo pagos reales (gastos + tareas que no son «Con/Sin maquinaria»):
 *      · pagos corrientes = los del mismo mes de hace un año (sin arrendamiento ni recolección contratada);
 *      · cosecha + subvenciones = presupuesto de la campaña (o lo real de la anterior si no hay) menos lo
 *        ya cobrado, repartido con el calendario medio de cobro de las dos últimas campañas cerradas;
 *      · arrendamiento = regla de «Parámetros» (la de Cuentas fincas) sobre el presupuesto, menos lo pagado,
 *        al mismo ritmo que la cosecha; recolección contratada = presupuesto menos lo pagado, igual.
 *  - Resto del negocio (como la Vista única): asesoría = cuotas activas de cada mes; productos = margen
 *    (venta − coste) del mismo mes del año pasado; comisiones = las del mismo mes del año pasado × lo que
 *    se suele cobrar; AgroServicios (sin amortización) = media mensual de este año.
 *  - Pendientes: cobros automáticos de asesoría y clientes de Ventas y todas las deudas, en 3 meses;
 *    comisiones pendientes y cobros a mano NO entran (son «liquidez que puedes activar»).
 *  - Préstamos: cuota de Finanzas. Mensual = cada mes; anual = la cuota entera en el mes del «Próximo cobro», cada año
 *    (si esa fecha ya ha pasado en el mes de partida y es anterior a la foto de Finanzas, se da por pagada). Hasta la fecha fin.
 *  - Impuestos (solo fincas, módulos): pago fraccionado 131 = 2 % de los ingresos agrarios del trimestre
 *    (prudente: sin restar retenciones); se paga en abril, julio, octubre y enero.
 */
'use strict';
const fs = require('fs');
const path = require('path');

const PERSONAL_MES = 2000;
const MESES_COLCHON = 3;
const HORIZONTE = 12;
const MESES_PENDIENTES = 3;
const PF_PORC = 0.02;
const avisos = [];
const MESES_NOMBRE = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

/* ---------- utilidades ---------- */
function num(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  let s = String(v).trim().replace(/\s|€/g, '');
  if (!s) return 0;
  const pc = s.endsWith('%');
  s = s.replace('%', '');
  if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  const n = parseFloat(s);
  return isFinite(n) ? (pc ? n / 100 : n) : 0;
}
const r2 = (n) => Math.round(n * 100) / 100;
const r0 = (n) => Math.round(n);
function fecha(v) { // dd/mm/aaaa, aaaa-mm-dd o fecha de Excel → 'aaaa-mm-dd'
  if (v instanceof Date) return new Date(v.getTime() + 12 * 3600e3).toISOString().slice(0, 10); // Excel guarda la medianoche de Madrid en UTC
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : '';
}
const ym = (iso) => iso.slice(0, 7);
function addM(k, n) { let [y, m] = k.split('-').map(Number); m += n; while (m > 12) { m -= 12; y++; } while (m < 1) { m += 12; y--; } return `${y}-${String(m).padStart(2, '0')}`; }
function diffM(a, b) { const [ya, ma] = a.split('-').map(Number), [yb, mb] = b.split('-').map(Number); return (yb - ya) * 12 + (mb - ma); }
const etiqueta = (k) => MESES_NOMBRE[Number(k.slice(5)) - 1] + ' ' + k.slice(0, 4);
const sum = (xs, f = (x) => x) => xs.reduce((t, x) => t + (Number(f(x)) || 0), 0);

// Cada fuente sale de un .json de get_values (nombre.json) o, si no está, de la pestaña del .xlsx exportado de Drive.
const XLSX_DE = {
  'fincas-tabla': ['fincas.xlsx', 'Tabla principal'], 'fincas-presupuestos': ['fincas.xlsx', 'Presupuestos'], 'fincas-parametros': ['fincas.xlsx', 'Parámetros'],
  'ventas-ventas': ['ventas.xlsx', 'Ventas'], 'ventas-comisiones': ['ventas.xlsx', 'Comisiones'],
  'fin-resumen': ['finanzas.xlsx', 'Resumen'], 'fin-cuentas': ['finanzas.xlsx', 'Cuentas'], 'fin-prestamos': ['finanzas.xlsx', 'Préstamos'], 'fin-deudas': ['finanzas.xlsx', 'Deudas'],
  'fin-cobros': ['finanzas.xlsx', 'Cobros'], 'fin-historico': ['finanzas.xlsx', 'Histórico'], 'asesor-clientes': ['asesor.xlsx', 'Clientes'], 'agro-trabajos': ['agro.xlsx', 'Trabajos externos'],
};
const libros = {};
function filasXlsx(fichero, pestana) {
  const f = path.join(dir, fichero);
  if (!fs.existsSync(f)) return null;
  let XLSX;
  try { XLSX = require('xlsx'); } catch (e) { avisos.push('Falta la librería xlsx (npm install xlsx) para leer ' + fichero + '.'); return null; }
  if (!libros[f]) libros[f] = XLSX.readFile(f, { cellDates: true });
  const sh = libros[f].Sheets[pestana];
  if (!sh) { avisos.push(`No encuentro la pestaña «${pestana}» en ${fichero}.`); return null; }
  return XLSX.utils.sheet_to_json(sh, { header: 1, raw: true, defval: '' });
}
function leer(dirDatos, nombre, obligatorio = true) {
  let vals = null;
  const f = path.join(dirDatos, nombre + '.json');
  if (fs.existsSync(f)) { const j = JSON.parse(fs.readFileSync(f, 'utf8')); vals = Array.isArray(j) ? j : (j.values || []); }
  else if (XLSX_DE[nombre]) vals = filasXlsx(...XLSX_DE[nombre]);
  if (!vals) { if (obligatorio) avisos.push(`Faltan los datos de ${nombre}: esa parte va a 0.`); return []; }
  if (!vals.length) return [];
  const cab = vals[0].map((c) => String(c).trim());
  return vals.slice(1).filter((r) => r && r.some((x) => String(x ?? '').trim())).map((r) => { const o = {}; cab.forEach((c, i) => { o[c] = r[i] == null ? '' : r[i]; }); return o; });
}

/* ---------- entrada ---------- */
const dir = process.argv[2];
if (!dir) { console.error('Uso: node prevision.js <carpeta-datos> [AAAA-MM-DD]'); process.exit(1); }
const hoy = process.argv[3] || new Date().toISOString().slice(0, 10);
const mes0 = ym(hoy);
const meses = Array.from({ length: HORIZONTE }, (_, i) => addM(mes0, i));

const resumen = leer(dir, 'fin-resumen');
const cuentas = leer(dir, 'fin-cuentas');
const prestamos = leer(dir, 'fin-prestamos');
const deudas = leer(dir, 'fin-deudas');
const cobros = leer(dir, 'fin-cobros');
const historico = leer(dir, 'fin-historico');
const tabla = leer(dir, 'fincas-tabla');
const presup = leer(dir, 'fincas-presupuestos');
const params = leer(dir, 'fincas-parametros');
const clientes = leer(dir, 'asesor-clientes');
const ventas = leer(dir, 'ventas-ventas');
const comisiones = leer(dir, 'ventas-comisiones');
const trabajos = leer(dir, 'agro-trabajos');

const R = {};
resumen.forEach((r) => { R[String(r.Concepto).trim()] = r.Valor; });

/* ---------- punto de partida ---------- */
const liqCuentas = sum(cuentas, (c) => num(c['Saldo (€)']));
const liquidez = R['Liquidez total'] != null ? num(R['Liquidez total']) : liqCuentas;
const fechaFoto = fecha(R.Fecha) || '';
const cuentasViejas = cuentas.filter((c) => { const f = fecha(c.Actualizado); return f && diffM(ym(f), mes0) >= 2; }).map((c) => c.Cuenta);

/* ---------- fincas ---------- */
const lineas = tabla.map((r) => {
  const f = fecha(r.Fecha);
  return { f, m: f ? ym(f) : '', camp: String(r['Campaña'] || '').trim(), cat: String(r.Categoria || r['Categoría'] || '').trim(), tipo: String(r.Tipo || '').trim(), tipo2: String(r['Tipo 2'] || '').trim(), finca: String(r.Finca || '').trim(), imp: num(r.Importe) };
}).filter((l) => l.m);
const esPago = (l) => l.cat === 'Gasto' || (l.cat === 'Tarea' && l.tipo2 !== 'Con maquinaria' && l.tipo2 !== 'Sin maquinaria');
const esArr = (l) => l.cat === 'Gasto' && l.tipo === 'Arrendamiento';
const esRecCont = (l) => l.cat === 'Tarea' && l.tipo === 'Recolección' && esPago(l);
const esIng = (l) => l.cat === 'Ingreso';

// Presupuesto por campaña y finca
const P = {}; // P[camp][fincaNombre][concepto] = valor
const fincaId = {};
presup.forEach((r) => {
  const c = String(r['Campaña']).trim(), fn = String(r.Finca).trim(), k = String(r.Concepto).trim();
  ((P[c] = P[c] || {})[fn] = P[c][fn] || {})[k] = num(r.Valor);
  fincaId[fn] = String(r['Finca ID']).trim();
});
const PAR = {};
params.forEach((r) => { PAR[String(r.Clave).trim()] = String(r.Valor).trim(); });
const nombreDeId = {};
params.forEach((r) => { const m = String(r.Clave).match(/^arr_pct\.(F\d+)$/); if (m) nombreDeId[m[1]] = String(r['Descripción'] || '').split('·')[0].trim(); });
Object.entries(fincaId).forEach(([n, id]) => { nombreDeId[id] = n; });

const ingresoPresu = (b) => (b.kg || 0) * (b.precio || 0) + (b.subvenciones || 0) + (b.otros_ingresos || 0);
function arrendamientoRegla(fn, b) {
  const id = fincaId[fn] || Object.keys(nombreDeId).find((k) => nombreDeId[k] === fn);
  const pct = num(PAR['arr_pct.' + id]) / 100, base = PAR['arr_base.' + id] || 'ninguno';
  if (!pct || base === 'ninguno') return 0;
  const cosecha = (b.kg || 0) * (b.precio || 0);
  const insumos = (b.fertilizantes || 0) + (b.fitosanitarios || 0) + (b.agua || 0) + (b.resto || 0);
  const moCont = (b.recoleccion || 0) + (b.poda || 0) + (b.tratamientos || 0) + (b.desbroce || 0) + (b.otras_tareas || 0);
  let v = 0;
  if (base === 'cosecha-recoleccion') v = cosecha - (b.recoleccion || 0);
  else if (base === 'cosecha-recoleccion-insumos-poda') v = cosecha - (b.recoleccion || 0) - insumos - (b.poda || 0);
  else if (base === 'ingresos-insumos-mocontratada') v = ingresoPresu(b) - insumos - moCont;
  return Math.max(0, pct * v);
}

// Calendario de cobro de la cosecha: reparto por «desfase» (meses desde enero del año de la campaña)
const campCerradas = (() => {
  const cs = [...new Set(lineas.filter(esIng).map((l) => l.camp))].filter((c) => /^\d{4}$/.test(c)).map(Number).sort();
  // cerrada = su desfase 20 (septiembre del año siguiente) ya ha pasado
  return cs.filter((c) => diffM(`${c}-01`, mes0) > 20).slice(-2);
})();
const perfil = new Array(36).fill(0);
campCerradas.forEach((c) => {
  const ls = lineas.filter((l) => esIng(l) && Number(l.camp) === c);
  const tot = sum(ls, (l) => l.imp);
  if (tot <= 0) return;
  ls.forEach((l) => { const d = diffM(`${c}-01`, l.m); if (d >= 0 && d < 36) perfil[d] += l.imp / tot / campCerradas.length; });
});
if (!campCerradas.length) avisos.push('No hay campañas cerradas para saber cuándo se cobra la cosecha: se reparte por igual.');

function repartoCampania(c, pendiente) { // reparte «pendiente» de la campaña c en los meses del horizonte
  const out = {};
  if (pendiente <= 0) return out;
  const dHoy = diffM(`${c}-01`, mes0);
  let masa = 0;
  for (let d = Math.max(0, dHoy); d < 36; d++) masa += perfil[d];
  if (masa < 0.02) { // ya casi no queda nada por cobrar según el calendario: lo que falte, en 3 meses
    for (let i = 0; i < 3; i++) out[meses[i]] = (out[meses[i]] || 0) + pendiente / 3;
    return out;
  }
  meses.forEach((m) => { const d = diffM(`${c}-01`, m); if (d >= 0 && d < 36 && perfil[d]) out[m] = pendiente * perfil[d] / masa; });
  return out;
}

const campActual = Number(mes0.slice(0, 4));
// Campañas abiertas: hasta septiembre del año siguiente (desfase 20); las cerradas ya no se prevén.
const campanias = [campActual - 1, campActual, campActual + 1].filter((c) => diffM(`${c}-01`, mes0) <= 20);
const fincasCamp = {}; // detalle por campaña
const flujoCosecha = {}, flujoArr = {}, flujoRec = {};
campanias.forEach((c) => {
  const cobrado = sum(lineas.filter((l) => esIng(l) && Number(l.camp) === c), (l) => l.imp);
  const arrPag = sum(lineas.filter((l) => esArr(l) && Number(l.camp) === c), (l) => l.imp);
  const recPag = sum(lineas.filter((l) => esRecCont(l) && Number(l.camp) === c), (l) => l.imp);
  let fuente = 'presupuesto';
  let presu = P[String(c)];
  if (!presu) { presu = P[String(c - 1)]; fuente = presu ? `presupuesto ${c - 1} (no hay de ${c})` : 'sin presupuesto'; }
  let ingEsp = 0, arrEsp = 0, recEsp = 0;
  const fincas = [];
  if (presu) Object.entries(presu).forEach(([fn, b]) => {
    const i = ingresoPresu(b), a = arrendamientoRegla(fn, b);
    ingEsp += i; arrEsp += a; recEsp += b.recoleccion || 0;
    fincas.push({ finca: fn, ingreso: r0(i), arrendamiento: r0(a), recoleccion: r0(b.recoleccion || 0) });
  });
  // fincas con ingresos en la campaña anterior pero sin presupuesto: lo real de la anterior
  const conPresu = new Set(Object.keys(presu || {}));
  const sinPresu = [...new Set(lineas.filter((l) => esIng(l) && Number(l.camp) === c - 1 && !conPresu.has(l.finca)).map((l) => l.finca))];
  sinPresu.forEach((fn) => {
    const v = sum(lineas.filter((l) => esIng(l) && Number(l.camp) === c - 1 && l.finca === fn), (l) => l.imp);
    ingEsp += v; fincas.push({ finca: fn, ingreso: r0(v), arrendamiento: 0, recoleccion: 0, nota: `sin presupuesto ${c}: lo real de ${c - 1}` });
  });
  const pend = Math.max(0, ingEsp - cobrado), pendArr = Math.max(0, arrEsp - arrPag), pendRec = Math.max(0, recEsp - recPag);
  const add = (dst, src) => Object.entries(src).forEach(([m, v]) => { dst[m] = (dst[m] || 0) + v; });
  add(flujoCosecha, repartoCampania(c, pend));
  add(flujoArr, repartoCampania(c, pendArr));
  add(flujoRec, repartoCampania(c, pendRec));
  fincasCamp[c] = { fuente, ingresoEsperado: r0(ingEsp), cobrado: r0(cobrado), porCobrar: r0(pend), arrendamientoEsperado: r0(arrEsp), arrendamientoPagado: r0(arrPag), arrendamientoPorPagar: r0(pendArr), recoleccionEsperada: r0(recEsp), recoleccionPagada: r0(recPag), recoleccionPorPagar: r0(pendRec), fincas };
});

// Pagos corrientes: los del mismo mes de hace un año
const pagosMes = {};
lineas.filter((l) => esPago(l) && !esArr(l) && !esRecCont(l)).forEach((l) => { pagosMes[l.m] = (pagosMes[l.m] || 0) + l.imp; });
const pagosCorrientes = (m) => pagosMes[addM(m, -12)] || 0;

/* ---------- resto del negocio ---------- */
// Asesoría: cuotas activas del mes
function cuotasMes(m) {
  return sum(clientes, (c) => {
    const p = num(c['Precio/mes (€)']); if (!p) return 0;
    const ini = String(c['Primer mes'] || '').trim().slice(0, 7), baja = String(c['Baja (último mes)'] || '').trim().slice(0, 7);
    if (ini && m < ini) return 0;
    if (baja && m > baja) return 0;
    return p;
  });
}
// Productos: margen del mismo mes del año pasado
const margenMes = {};
ventas.forEach((v) => { const f = fecha(v.Fecha); if (!f) return; margenMes[ym(f)] = (margenMes[ym(f)] || 0) + num(v['Importe (€)']) - num(v['Importe coste (€)']); });
// Comisiones: las del mismo mes del año pasado × proporción que se suele cobrar
const comMes = {};
let comTot = 0, comCob = 0;
comisiones.forEach((c) => { const f = fecha(c.Fecha); const v = num(c['Comisión (€)']); if (!f) return; comMes[ym(f)] = (comMes[ym(f)] || 0) + v; comTot += v; if (/^s/i.test(String(c.Cobrada))) comCob += v; });
const comRatio = comTot ? comCob / comTot : 0;
// AgroServicios: beneficio sin amortización, media mensual de este año
const anio = mes0.slice(0, 4);
const agroAnio = sum(trabajos.filter((t) => fecha(t.Fecha).startsWith(anio)), (t) => num(t['Importe (€)']) - num(t['Otros gastos (€)']));
const mesesTranscurridos = Math.max(1, Number(mes0.slice(5)) - 1);
const agroMes = agroAnio / mesesTranscurridos;

/* ---------- préstamos ---------- */
// Caída media normal del capital (sin amortizaciones extraordinarias), por si falta la cuota
const capHist = historico.map((h) => ({ m: ym(fecha(h['Mes (cierre)'])), v: num(h['Préstamos']) })).filter((x) => x.m).sort((a, b) => a.m.localeCompare(b.m));
const caidas = [];
for (let i = 1; i < capHist.length; i++) { const d = capHist[i - 1].v - capHist[i].v, n = Math.max(1, diffM(capHist[i - 1].m, capHist[i].m)); if (d > 0 && d / n < 2000) caidas.push(d / n); }
const caidaMedia = caidas.length ? sum(caidas) / caidas.length : 0;
const capTotal = sum(prestamos, (p) => num(p['Capital pendiente (€)']));
const loans = prestamos.map((p) => {
  const cap = num(p['Capital pendiente (€)']);
  const tin = String(p['Interés TIN (%)'] ?? '').trim() === '' ? null : num(p['Interés TIN (%)']);
  const cuotaTxt = p['Cuota (€)'] ?? p['Cuota mensual (€)'] ?? ''; // «Cuota mensual (€)» hasta Finanzas 1.2.0
  const cuota = String(cuotaTxt).trim() === '' ? null : num(cuotaTxt);
  const fin = fecha(p['Fecha fin']) || '';
  const anual = /^anual/i.test(String(p['Periodicidad'] ?? '').trim());
  const prox = anual ? (fecha(p['Próximo cobro']) || '') : '';
  if (anual && !prox) avisos.push(`El préstamo «${p['Préstamo']}» es anual pero no tiene fecha de próximo cobro: lo reparto como si fuera mensual (cuota / 12).`);
  const est = cuota == null;
  const cuotaUsada = cuota != null ? cuota : (capTotal ? caidaMedia * cap / capTotal * (anual ? 12 : 1) : 0);
  let interesMes = tin != null ? cap * (tin / 100) / 12 : null;
  const mesesRest = fin ? Math.max(0, diffM(mes0, ym(fin)) + 1) : null;
  return { nombre: p['Préstamo'], periodicidad: anual ? 'anual' : 'mensual', proximoCobro: prox, capital: r2(cap), tin, ahorroAnualPor1000: tin != null ? r2(10 * tin) : null, cuota, cuotaUsada: r2(cuotaUsada), cuotaEstimada: est, fin, mesesRestantes: mesesRest, interesAnual: tin != null ? r0(cap * tin / 100) : null, interesMes: interesMes != null ? r2(interesMes) : null };
});
if (loans.some((l) => l.cuotaEstimada || l.tin == null || !l.fin)) avisos.push('Faltan datos de algún préstamo (interés, cuota o fecha de fin) en Finanzas → Deudas y cobros → Préstamos. Hasta que estén, la cuota se estima con lo que baja el capital cada mes y no se puede valorar si conviene amortizar.');
function cuotaPrestamo(l, m) {
  if (l.fin && m > ym(l.fin)) return 0;
  if (l.periodicidad !== 'anual') return l.cuotaUsada;
  if (!l.proximoCobro) return l.cuotaUsada / 12;
  if (m.slice(5) !== l.proximoCobro.slice(5, 7)) return 0;
  const dia = `${m}-${l.proximoCobro.slice(8, 10)}`; // fecha de cobro de ese año
  if (m === mes0 && fechaFoto && dia <= fechaFoto) return 0; // ya pagado: está en la liquidez de partida
  return l.cuotaUsada;
}
const cuotaMes = (m) => sum(loans, (l) => cuotaPrestamo(l, m));

/* ---------- pendientes ---------- */
const pendDe = (r) => num(r['Pendiente (€)']);
const deudasAbiertas = deudas.filter((d) => pendDe(d) > 0.005).map((d) => ({ quien: d['A quién'], concepto: d.Concepto || '', pendiente: r2(pendDe(d)), desde: fecha(d.Fecha) }));
const deudaPend = sum(deudasAbiertas, (d) => d.pendiente);
const cobrosAbiertos = cobros.filter((c) => pendDe(c) > 0.005).map((c) => ({ quien: c['Quién'], concepto: c.Concepto || '', pendiente: r2(pendDe(c)), auto: /automático/i.test(String(c['Quién'])) }));
const cobBase = sum(cobrosAbiertos.filter((c) => c.auto && !/comisiones/i.test(c.quien)), (c) => c.pendiente);
const cobFuera = cobrosAbiertos.filter((c) => !c.auto || /comisiones/i.test(c.quien));
const asesoriaPendientes = clientes.filter((c) => num(c['Pendiente 2026 (€)'] ?? c['Pendiente (€)']) > 0).map((c) => ({ cliente: c.Cliente, pendiente: r2(num(c['Pendiente 2026 (€)'] ?? c['Pendiente (€)'])), precioMes: num(c['Precio/mes (€)']) })).sort((a, b) => b.pendiente - a.pendiente);

/* ---------- impuestos de las fincas (módulos) ---------- */
const ingFincasReal = {};
lineas.filter(esIng).forEach((l) => { ingFincasReal[l.m] = (ingFincasReal[l.m] || 0) + l.imp; });
const ingFincasMes = (m) => (m < mes0 ? (ingFincasReal[m] || 0) : (flujoCosecha[m] || 0));
function pagoFraccionado(m) { // enero, abril, julio, octubre: 2 % del trimestre anterior
  const mm = Number(m.slice(5));
  if (![1, 4, 7, 10].includes(mm)) return 0;
  const tri = [addM(m, -3), addM(m, -2), addM(m, -1)];
  return PF_PORC * sum(tri, ingFincasMes);
}

/* ---------- previsión mes a mes ---------- */
let saldo = liquidez;
const filas = meses.map((m, i) => {
  const e = {
    fincas: flujoCosecha[m] || 0,
    asesoria: cuotasMes(m),
    productos: margenMes[addM(m, -12)] || 0,
    comisiones: (comMes[addM(m, -12)] || 0) * comRatio,
    agroservicios: agroMes,
    cobrosPendientes: i < MESES_PENDIENTES ? cobBase / MESES_PENDIENTES : 0,
  };
  const s = {
    personal: PERSONAL_MES,
    fincasCorrientes: pagosCorrientes(m),
    arrendamientos: flujoArr[m] || 0,
    recoleccion: flujoRec[m] || 0,
    prestamos: cuotaMes(m),
    deudasPendientes: i < MESES_PENDIENTES ? deudaPend / MESES_PENDIENTES : 0,
    impuestosFincas: pagoFraccionado(m),
  };
  const ent = sum(Object.values(e)), sal = sum(Object.values(s));
  saldo += ent - sal;
  const o = { mes: m, etiqueta: etiqueta(m), entradas: {}, salidas: {}, totalEntradas: r0(ent), totalSalidas: r0(sal), neto: r0(ent - sal), saldoFinal: r0(saldo) };
  Object.entries(e).forEach(([k, v]) => { o.entradas[k] = r0(v); });
  Object.entries(s).forEach(([k, v]) => { o.salidas[k] = r0(v); });
  return o;
});
const recurrenteMedio = sum(filas, (f) => f.salidas.personal + f.salidas.fincasCorrientes + f.salidas.prestamos) / filas.length;
const colchon = r0(MESES_COLCHON * recurrenteMedio);
filas.forEach((f) => { f.holgura = f.saldoFinal - colchon; f.bajoColchon = f.saldoFinal < colchon; f.negativo = f.saldoFinal < 0; });
const peor = filas.reduce((a, b) => (b.saldoFinal < a.saldoFinal ? b : a), filas[0]);
const excedenteSeguro = Math.max(0, peor.saldoFinal - colchon); // lo que se puede usar hoy sin bajar del colchón en ningún mes

/* ---------- datos para las recomendaciones ---------- */
const conInteres = loans.filter((l) => l.tin != null).sort((a, b) => b.tin - a.tin);
const amortizacion = conInteres.length && excedenteSeguro > 0 ? (() => {
  const l = conInteres[0], imp = Math.min(excedenteSeguro, l.capital);
  return { prestamo: l.nombre, tin: l.tin, importePosible: r0(imp), ahorroInteresesPrimerAnio: r0(imp * l.tin / 100), nota: 'Ahorro aproximado del primer año (interés simple sobre lo amortizado). Revisar comisión de amortización anticipada del contrato.' };
})() : null;

const ultimoCierre = historico.length ? historico[historico.length - 1] : null;
const stockTotal = num(R['Total stock']);

const salida = {
  version: '1.0',
  generado: new Date().toISOString(),
  fechaEjecucion: hoy,
  horizonte: { desde: meses[0], hasta: meses[meses.length - 1] },
  partida: {
    liquidez: r0(liquidez), fechaFoto, cuentasSinActualizar: cuentasViejas,
    ultimoCierre: ultimoCierre ? { mes: fecha(ultimoCierre['Mes (cierre)']), liquidez: r0(num(ultimoCierre['Liquidez total'])), capitalNeto: r0(num(ultimoCierre['Capital neto'])), ratioLiquidezDeuda: num(ultimoCierre['Ratio liquidez/deuda']), comentario: ultimoCierre.Comentario || '' } : null,
    capitalNeto: r0(num(R['Capital neto'])), totalDeuda: r0(num(R['Total deuda'])), stock: r0(stockTotal),
  },
  colchon: { meses: MESES_COLCHON, importe: colchon, salidaRecurrenteMedia: r0(recurrenteMedio) },
  resumen: {
    saldoFinal12m: filas[filas.length - 1].saldoFinal,
    peorMes: { mes: peor.etiqueta, saldo: peor.saldoFinal, holgura: r0(peor.holgura) },
    mesesBajoColchon: filas.filter((f) => f.bajoColchon).map((f) => f.etiqueta),
    mesesNegativos: filas.filter((f) => f.negativo).map((f) => f.etiqueta),
    excedenteSeguro: r0(excedenteSeguro),
    entradas12m: r0(sum(filas, (f) => f.totalEntradas)), salidas12m: r0(sum(filas, (f) => f.totalSalidas)),
  },
  meses: filas,
  fincas: { campaniasPerfil: campCerradas, perfilCobroPorDesfase: perfil.map((x) => Math.round(x * 1000) / 10), campanias: fincasCamp },
  negocio: { cuotasAsesoriaMes: r0(cuotasMes(mes0)), comisionesRatioCobro: Math.round(comRatio * 100), agroservicios: { beneficioAnio: r0(agroAnio), mediaMes: r0(agroMes) } },
  prestamos: { lista: loans, capitalTotal: r0(capTotal), caidaMediaCapitalMes: r0(caidaMedia), amortizacion },
  liquidez: {
    deudasPendientes: deudasAbiertas, totalDeudas: r0(deudaPend),
    cobrosEnPrevision: r0(cobBase), cobrosFueraDePrevision: cobFuera, totalFueraDePrevision: r0(sum(cobFuera, (c) => c.pendiente)),
    asesoriaPendientePorCliente: asesoriaPendientes.slice(0, 10), stock: r0(stockTotal),
  },
  fiscalFincas: {
    regimen: 'IRPF estimación objetiva (módulos) solo para la actividad agraria (decisión de Francis: el resto no se tiene en cuenta)',
    pagosFraccionados: filas.filter((f) => f.salidas.impuestosFincas > 0).map((f) => ({ mes: f.etiqueta, importe: f.salidas.impuestosFincas })),
    ingresosAgrariosAnioEnCurso: r0(sum(Object.entries(ingFincasReal).filter(([m]) => m.startsWith(anio)), ([, v]) => v)),
  },
  supuestos: [
    `Personal: ${PERSONAL_MES} €/mes fijos.`,
    'Fincas: pagos corrientes = los del mismo mes de hace un año; cosecha, arrendamiento y recolección contratada = presupuesto de la campaña menos lo ya anotado, con el calendario de cobro de las dos últimas campañas cerradas.',
    'Arrendamientos con la regla de «Parámetros» (la misma de Cuentas fincas) sobre el presupuesto.',
    'Asesoría: cuotas activas cada mes. Productos: margen del mismo mes del año pasado. Comisiones: las del año pasado × la parte que se suele cobrar. AgroServicios: media mensual de este año.',
    `Deudas pendientes y cobros automáticos de asesoría y clientes de Ventas: repartidos en ${MESES_PENDIENTES} meses. Comisiones pendientes y cobros a mano: fuera de la previsión.`,
    'Pago fraccionado (131) de las fincas: 2 % de los ingresos agrarios del trimestre, sin restar retenciones (prudente).',
    'Préstamos: cuota de Finanzas; si falta, la bajada media del capital (sin intereses).',
  ],
  avisos,
};
process.stdout.write(JSON.stringify(salida, null, 1));
