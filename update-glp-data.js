/* Automatizaciones Digitales — refresco de datos del GLP (corre en GitHub Actions, a diario).
 * Escribe glp-data.json en la raíz del repo. La página /glp lo lee desde raw.githubusercontent
 * (CORS abierto), así que se actualiza SIN redeploy. La TRM va aparte, en vivo del lado del cliente.
 *
 * Fuentes:
 *  - Molécula regulada del productor, por fuente: Ecopetrol PME-VPRECIOS .xls  [SIN LLAVE]
 *  - GLP Capachos: precio LIBRE (no regulado) publicado en la misma hoja      [SIN LLAVE]
 *  - Transporte del producto importado hasta Barranca: nota al pie de la hoja [SIN LLAVE]
 *  - Mont Belvieu propano/butano y Brent: EIA API v2  [OPCIONAL: EIA_API_KEY de los Secrets]
 *
 * OJO — "Precio Importado": Ecopetrol dejó de publicar esa columna. La última hoja que la trae
 * es "DIC 15 2018- ENE 14 2019". No es un fallo del parser: la fuente cambió. Se reporta el
 * último valor publicado como referencia HISTÓRICA y se deja explícito que el nivel vivo del
 * importado entra por paridad de importación (Mont Belvieu × TRM), no por esta hoja.
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ECO_XLS = 'https://www.ecopetrol.com.co/wps/wcm/connect/d29a5ea7-3017-41b6-a3c7-91d9f09db257/PME-VPRECIOS+GLP+2019+-+2026.xls?MOD=AJPERES&CVID=pZcMRdF';
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AutomatizacionesDigitales-GLP/1.0';

const num = v => (typeof v === 'number' && isFinite(v)) ? v : null;
const r2  = v => (v == null ? null : +v.toFixed(2));
/* "1.408,55" / "125,12" → 1408.55 / 125.12 (las notas al pie vienen en formato colombiano) */
const coma = s => { const v = parseFloat(String(s).trim().replace(/\./g, '').replace(',', '.')); return isFinite(v) ? v : null; };
const filas = hoja => XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '' });
const buscaFila = (rows, ...frags) => rows.find(r => { const s = String(r[1]).toLowerCase(); return frags.every(f => s.includes(f)); });
const primerNum = row => row ? (row.slice(2).map(num).find(v => v != null) ?? null) : null;

/* Precio LIBRE (no regulado) del GLP del campo Capachos. Es el único precio de mercado que
 * Ecopetrol publica en esta hoja; el resto son precios máximos regulados por la CREG. */
function capachos(rows) {
  const kg = primerNum(buscaFila(rows, 'suministro de referencia', '($/kg)'));
  if (kg == null) return null;
  const gl = primerNum(buscaFila(rows, 'suministro de referencia', '($/gl)'));
  const i  = rows.findIndex(r => String(r[1]).toUpperCase().includes('PRECIO LIBRE DE GLP CAPACHOS'));
  const periodo = i >= 0 ? String(rows[i + 1]?.[1] || '').trim() : null;
  return { kg: r2(kg), gl: r2(gl), periodo: periodo || null, regulado: false };
}

/* Nota al pie (**): lo que cuesta mover producto IMPORTADO hasta Barrancabermeja.
 * Sirve de contraste contra el flete carrotanque del modelo. */
function transporteImportado(rows) {
  const r = rows.find(x => String(x[1]).toLowerCase().includes('producto importado'));
  if (!r) return null;
  const t = String(r[1]).replace(/\s+/g, ' ');
  const kg = t.match(/\$\s*\/\s*kg\s*([\d.]*\d(?:,\d+)?)/i);
  const gl = t.match(/\$\s*\/\s*gl\s*([\d.]*\d(?:,\d+)?)/i);
  if (!kg) return null;
  return { kg: r2(coma(kg[1])), gl: gl ? r2(coma(gl[1])) : null, destino: 'Refinería de Barrancabermeja' };
}

/* Última hoja del libro que todavía traía la columna "Precio Importado". Se recorre el libro
 * completo a propósito: si Ecopetrol algún día la revive, este mismo código la vuelve a ver. */
function importadoHistorico(wb) {
  let ultimo = null;
  for (const n of wb.SheetNames) {
    const rows = filas(wb.Sheets[n]);
    const hdr = buscaFila(rows, 'concepto');
    if (!hdr) continue;
    const c = hdr.findIndex(x => String(x).toLowerCase().includes('import'));
    if (c < 0) continue;
    const v = num(buscaFila(rows, '($/kg)')?.[c]);
    if (v != null) ultimo = { periodo: n.trim(), kg: r2(v) };
  }
  return ultimo;
}

async function ecopetrol() {
  const res = await fetch(ECO_XLS, { headers: { 'User-Agent': UA } });
  if (!res.ok) throw new Error('Ecopetrol xls HTTP ' + res.status);
  const buf = Buffer.from(await res.arrayBuffer());
  const wb = XLSX.read(buf, { type: 'buffer' });
  const period = wb.SheetNames[wb.SheetNames.length - 1];           // hoja más reciente
  const rows = filas(wb.Sheets[period]);
  const hdr = buscaFila(rows, 'concepto');
  const kg  = buscaFila(rows, 'kg');
  if (!hdr || !kg) throw new Error('Ecopetrol: no encontré header / fila $/KG en ' + period);
  const by = {}, sinVentas = [];
  for (let c = 2; c < hdr.length; c++) {
    const name = String(hdr[c]).toLowerCase();
    if (!name.trim()) continue;
    let k = null;
    if (name.includes('barranca')) k = 'barranca';
    else if (name.includes('reficar')) k = 'reficar';
    else if (name.includes('cusiana')) k = 'cusiana';
    else if (name.includes('cupiagua')) k = 'cupiagua';
    else if (name.includes('dina')) k = 'dina';
    else if (name.includes('apiay')) k = 'apiay';
    else if (name.includes('import')) k = 'importado';
    if (!k) continue;
    const val = num(kg[c]);
    if (val == null) { sinVentas.push(k); continue; }               // "SIN VENTAS" → la fuente no despachó
    by[k] = val;
  }
  const interior = by.cusiana ?? by.barranca ?? by.dina ?? by.cupiagua ?? null;
  const hist = importadoHistorico(wb);
  const nota = by.importado == null
    ? 'Ecopetrol ya no publica la columna "Precio Importado" (la última fue ' +
      (hist ? hist.periodo + ', ' + hist.kg + ' $/kg' : 'anterior a 2019') +
      '). El nivel vivo del importado se estima por paridad de importación (Mont Belvieu × TRM), no por esta hoja.'
    : null;
  return {
    period: period.trim(),
    interior, reficar: by.reficar ?? null,
    importado: by.importado ?? null,
    importadoNota: nota,
    importadoHist: hist,
    capachos: capachos(rows),
    transporteImportado: transporteImportado(rows),
    sinVentas,
    porFuente: by,
  };
}

async function eia(serie, etiqueta) {
  const key = process.env.EIA_API_KEY;
  if (!key) return null;                                            // sin llave → se omite (no va al frontend)
  const url = `https://api.eia.gov/v2/petroleum/pri/spt/data/?api_key=${key}` +
    `&frequency=daily&data[0]=value&facets[series][]=${serie}` +
    `&sort[0][column]=period&sort[0][direction]=desc&length=1`;
  try {
    const r = await fetch(url, { headers: { 'User-Agent': UA } });
    if (!r.ok) throw new Error('EIA HTTP ' + r.status);
    const row = (await r.json())?.response?.data?.[0];
    const v = num(parseFloat(row?.value));
    return v == null ? null : { valor: v, date: row.period };
  } catch (e) { console.error(etiqueta + ' EIA:', e.message); return null; }
}

async function montBelvieu() {
  const d = await eia('EER_EPLLPA_PF4_Y44MB_DPG', 'MB');
  if (!d) return null;
  return { propano: +d.valor.toFixed(3), butano: +(d.valor * 1.18).toFixed(3), date: d.date, fuente: 'EIA (butano ≈ propano×1,18)' };
}

async function brent() {
  const d = await eia('RBRTE', 'Brent');                            // Colombia exporta crudo → Brent↑ tiende a TRM↓
  if (!d) return null;
  return { usd_bbl: r2(d.valor), date: d.date, fuente: 'EIA (Europe Brent Spot, RBRTE)' };
}

(async () => {
  const out = { updated: new Date().toISOString() };
  try {
    out.eco = await ecopetrol();
    console.log('Ecopetrol OK:', out.eco.period, 'interior', out.eco.interior, 'reficar', out.eco.reficar);
    console.log('  Capachos (libre):', out.eco.capachos ? out.eco.capachos.kg + ' $/kg · ' + out.eco.capachos.periodo : '(no publicado)');
    console.log('  Transporte importado→Barranca:', out.eco.transporteImportado ? out.eco.transporteImportado.kg + ' $/kg' : '(no publicado)');
    console.log('  Importado:', out.eco.importado ?? '(columna descontinuada; último ' + (out.eco.importadoHist ? out.eco.importadoHist.kg + ' $/kg en ' + out.eco.importadoHist.periodo : 'n/d') + ')');
    if (out.eco.sinVentas.length) console.log('  Sin ventas este período:', out.eco.sinVentas.join(', '));
  } catch (e) { console.error('Ecopetrol FALLÓ:', e.message); out.ecoError = e.message; }
  out.mb = await montBelvieu();
  console.log('MB:', out.mb ? `${out.mb.propano}/${out.mb.butano} @ ${out.mb.date}` : '(sin llave EIA — omitido)');
  out.brent = await brent();
  console.log('Brent:', out.brent ? `${out.brent.usd_bbl} @ ${out.brent.date}` : '(sin llave EIA — omitido)');

  const dest = path.join(__dirname, 'glp-data.json');
  fs.writeFileSync(dest, JSON.stringify(out, null, 2) + '\n');
  console.log('Escrito', dest);
})();
