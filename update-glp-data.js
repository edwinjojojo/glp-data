/* Automatizaciones Digitales — refresco de datos del GLP (corre en GitHub Actions, a diario).
 * Escribe glp-data.json en la raíz del repo. La página /glp lo lee desde raw.githubusercontent
 * (CORS abierto), así que se actualiza SIN redeploy. La TRM va aparte, en vivo del lado del cliente.
 *
 * Fuentes:
 *  - Molécula del productor (nacional por fuente + importado): Ecopetrol PME-VPRECIOS .xls  [SIN LLAVE]
 *  - Mont Belvieu propano/butano: EIA API v2  [OPCIONAL: usa EIA_API_KEY de los Secrets; si no hay, se omite]
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ECO_XLS = 'https://www.ecopetrol.com.co/wps/wcm/connect/d29a5ea7-3017-41b6-a3c7-91d9f09db257/PME-VPRECIOS+GLP+2019+-+2026.xls?MOD=AJPERES&CVID=pZcMRdF';
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AutomatizacionesDigitales-GLP/1.0';

const num = v => (typeof v === 'number' && isFinite(v)) ? v : null;

async function ecopetrol() {
  const res = await fetch(ECO_XLS, { headers: { 'User-Agent': UA } });
  if (!res.ok) throw new Error('Ecopetrol xls HTTP ' + res.status);
  const buf = Buffer.from(await res.arrayBuffer());
  const wb = XLSX.read(buf, { type: 'buffer' });
  const period = wb.SheetNames[wb.SheetNames.length - 1];           // hoja más reciente
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[period], { header: 1, defval: '' });
  const hdr = rows.find(r => String(r[1]).toLowerCase().includes('concepto'));
  const kg  = rows.find(r => String(r[1]).toLowerCase().includes('kg'));
  if (!hdr || !kg) throw new Error('Ecopetrol: no encontré header / fila $/KG en ' + period);
  const by = {};
  for (let c = 2; c < hdr.length; c++) {
    const name = String(hdr[c]).toLowerCase();
    const val = num(kg[c]);
    if (val == null) continue;
    if (name.includes('barranca')) by.barranca = val;
    else if (name.includes('reficar')) by.reficar = val;
    else if (name.includes('cusiana')) by.cusiana = val;
    else if (name.includes('cupiagua')) by.cupiagua = val;
    else if (name.includes('dina')) by.dina = val;
    else if (name.includes('apiay')) by.apiay = val;
    else if (name.includes('import')) by.importado = val;
  }
  const interior = by.cusiana ?? by.barranca ?? by.dina ?? by.cupiagua ?? null;
  return {
    period: period.trim(),
    interior, reficar: by.reficar ?? null,
    importado: by.importado ?? null,
    porFuente: by,
  };
}

async function montBelvieu() {
  const key = process.env.EIA_API_KEY;
  if (!key) return null;                                            // sin llave → se omite (no va al frontend)
  const url = `https://api.eia.gov/v2/petroleum/pri/spt/data/?api_key=${key}` +
    `&frequency=daily&data[0]=value&facets[series][]=EER_EPLLPA_PF4_Y44MB_DPG` +
    `&sort[0][column]=period&sort[0][direction]=desc&length=1`;
  try {
    const r = await fetch(url, { headers: { 'User-Agent': UA } });
    if (!r.ok) throw new Error('EIA HTTP ' + r.status);
    const j = await r.json();
    const row = j?.response?.data?.[0];
    const propano = num(parseFloat(row?.value));
    if (propano == null) return null;
    return { propano: +propano.toFixed(3), butano: +(propano * 1.18).toFixed(3), date: row.period, fuente: 'EIA (butano ≈ propano×1,18)' };
  } catch (e) { console.error('MB EIA:', e.message); return null; }
}

async function brent() {
  const key = process.env.EIA_API_KEY;
  if (!key) return null;                                            // contexto macro; Colombia exporta crudo → Brent↑ tiende a TRM↓
  const url = `https://api.eia.gov/v2/petroleum/pri/spt/data/?api_key=${key}` +
    `&frequency=daily&data[0]=value&facets[series][]=RBRTE` +
    `&sort[0][column]=period&sort[0][direction]=desc&length=1`;
  try {
    const r = await fetch(url, { headers: { 'User-Agent': UA } });
    if (!r.ok) throw new Error('EIA HTTP ' + r.status);
    const j = await r.json();
    const row = j?.response?.data?.[0];
    const v = num(parseFloat(row?.value));
    if (v == null) return null;
    return { usd_bbl: +v.toFixed(2), date: row.period, fuente: 'EIA (Europe Brent Spot, RBRTE)' };
  } catch (e) { console.error('Brent EIA:', e.message); return null; }
}

(async () => {
  const out = { updated: new Date().toISOString() };
  try { out.eco = await ecopetrol(); console.log('Ecopetrol OK:', out.eco.period, 'interior', out.eco.interior, 'reficar', out.eco.reficar); }
  catch (e) { console.error('Ecopetrol FALLÓ:', e.message); out.ecoError = e.message; }
  out.mb = await montBelvieu();
  console.log('MB:', out.mb ? `${out.mb.propano}/${out.mb.butano} @ ${out.mb.date}` : '(sin llave EIA — omitido)');
  out.brent = await brent();
  console.log('Brent:', out.brent ? `${out.brent.usd_bbl} @ ${out.brent.date}` : '(sin llave EIA — omitido)');

  const dest = path.join(__dirname, 'glp-data.json');
  fs.writeFileSync(dest, JSON.stringify(out, null, 2) + '\n');
  console.log('Escrito', dest);
})();
