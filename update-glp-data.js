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
/* 2111.22 -> "2.111,22": la nota se lee en la pagina, en formato colombiano */
const co  = n => { const [e, d] = n.toFixed(2).split('.'); return e.replace(/\B(?=(\d{3})+$)/g, '.') + ',' + d; };
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
      (hist ? hist.periodo + ', ' + co(hist.kg) + ' $/kg' : 'anterior a 2019') +
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

/* ===== PARTICIPACIÓN HISTÓRICA EN EL MERCADO NACIONAL (SUI, formato B1) =====
 * dda6-rcne = "Origen del producto para consumo nacional de GLP". Trae el NIT de quien produjo
 * o importó cada kilo, así que la participación se MIDE, no se estima.
 *
 * Dos trampas de este dataset:
 *  - Ecopetrol aparece con dos escrituras del NIT (con y sin dígito de verificación).
 *  - 9001125157 es REFICAR, filial 100% de Ecopetrol. Sin consolidarla, el dominio de partida
 *    sale subestimado en más de 20 puntos. Se reportan las dos lecturas y que decida quien lee.
 *  - `anio` y `mes` son TEXTO en Socrata: nada de comparaciones numéricas en el $where.
 * Los últimos meses llegan incompletos (el reporte va entrando), así que se recortan solos. */
const SUI_B1 = 'https://www.datos.gov.co/resource/dda6-rcne.json?' +
  '$select=anio,mes,nit_productor_y_o_importador as nit,nombre_fuente_produccion as fuente,sum(cantidad_kg) as kg' +
  '&$group=anio,mes,nit_productor_y_o_importador,nombre_fuente_produccion&$limit=20000';
const NIT_ECOPETROL = ['8999990681', '899999068'];   // matriz, con y sin dígito de verificación
const NIT_REFICAR   = ['9001125157'];                // Refinería de Cartagena, filial 100%

async function participacion() {
  const r = await fetch(SUI_B1, { headers: { 'User-Agent': UA } });
  if (!r.ok) throw new Error('SUI B1 HTTP ' + r.status);
  const filas = await r.json();
  if (!Array.isArray(filas) || !filas.length) throw new Error('SUI B1: respuesta vacía');

  /* Mes completo = el último que no se desploma contra la mediana de los 12 anteriores.
     Evita clavar un corte fijo que se vuelve mentira el mes entrante. */
  const porMes = {};
  filas.forEach(f => { const k = f.anio + '-' + String(f.mes).padStart(2, '0'); porMes[k] = (porMes[k] || 0) + Number(f.kg); });
  const claves = Object.keys(porMes).sort();
  const ult12 = claves.slice(-12).map(k => porMes[k]).sort((a, b) => a - b);
  const mediana = ult12[Math.floor(ult12.length / 2)] || 0;
  let i = claves.length - 1;
  while (i > 0 && porMes[claves[i]] < mediana * 0.6) i--;
  const corte = claves[i];

  const esImp = f => /IMPORTA/i.test(f || '');
  const Y = {};
  filas.forEach(f => {
    const k = f.anio + '-' + String(f.mes).padStart(2, '0');
    if (k > corte) return;
    const a = f.anio, kg = Number(f.kg);
    const y = Y[a] = Y[a] || { anio: +a, kt: 0, ecoNac: 0, refNac: 0, grpImp: 0, terNac: 0, terImp: 0 };
    const eco = NIT_ECOPETROL.includes(f.nit), ref = NIT_REFICAR.includes(f.nit);
    if ((eco || ref) && esImp(f.fuente)) y.grpImp += kg;   // el propio grupo también importa
    else if (eco) y.ecoNac += kg;
    else if (ref) y.refNac += kg;
    else if (esImp(f.fuente)) y.terImp += kg;
    else y.terNac += kg;
    y.kt += kg;
  });

  const pc = (v, t) => t ? +(100 * v / t).toFixed(1) : null;
  const serie = Object.values(Y).sort((a, b) => a.anio - b.anio).map(y => ({
    anio: y.anio,
    kt: +(y.kt / 1e6).toFixed(1),
    grupo:  pc(y.ecoNac + y.refNac + y.grpImp, y.kt),   // Grupo Ecopetrol = matriz + Reficar
    matriz: pc(y.ecoNac + y.grpImp, y.kt),              // solo Ecopetrol S.A.
    ecoNac: pc(y.ecoNac, y.kt),
    refNac: pc(y.refNac, y.kt),
    grpImp: pc(y.grpImp, y.kt),
    terNac: pc(y.terNac, y.kt),
    terImp: pc(y.terImp, y.kt),
  }));
  const anioCorte = +corte.slice(0, 4), mesCorte = +corte.slice(5);
  return {
    corte, serie,
    parcial: mesCorte < 12 ? anioCorte : null,     // el último año va incompleto: hay que decirlo
    mesesParcial: mesCorte < 12 ? mesCorte : null,
    fuente: 'SSPD · SUI formato B1 (Origen del producto para consumo nacional), datos.gov.co dda6-rcne',
    url: 'https://www.datos.gov.co/d/dda6-rcne',
    nota: 'Reficar (NIT 900112515) se consolida dentro del Grupo Ecopetrol por ser filial 100%. "matriz" excluye Reficar.',
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
  try {
    out.part = await participacion();
    const s = out.part.serie;
    console.log('Participación OK: corte', out.part.corte, '·', s.length, 'años ·',
      'Grupo Ecopetrol ' + s[0].anio + ' ' + s[0].grupo + '% → ' + s[s.length - 1].anio + ' ' + s[s.length - 1].grupo + '%',
      '| terceros importadores ' + s[0].terImp + '% → ' + s[s.length - 1].terImp + '%');
  } catch (e) { console.error('Participación FALLÓ:', e.message); out.partError = e.message; }
  out.mb = await montBelvieu();
  console.log('MB:', out.mb ? `${out.mb.propano}/${out.mb.butano} @ ${out.mb.date}` : '(sin llave EIA — omitido)');
  out.brent = await brent();
  console.log('Brent:', out.brent ? `${out.brent.usd_bbl} @ ${out.brent.date}` : '(sin llave EIA — omitido)');

  const dest = path.join(__dirname, 'glp-data.json');
  fs.writeFileSync(dest, JSON.stringify(out, null, 2) + '\n');
  console.log('Escrito', dest);
})();
