/* ===== Núcleo: lectura de archivos y cálculo de bonos (sin dependencias de UI) ===== */
const IVA = 1.19;
const TOL = 1000; // tolerancia en pesos para considerar que dos montos calzan

const MESES = { ENERO:1, FEBRERO:2, MARZO:3, ABRIL:4, MAYO:5, JUNIO:6, JULIO:7, AGOSTO:8, SEPTIEMBRE:9, SETIEMBRE:9, OCTUBRE:10, NOVIEMBRE:11, DICIEMBRE:12 };
const MESES_NOMBRE = ['', 'Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

function norm(s) {
  if (s === null || s === undefined) return '';
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
}
function num(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).replace(/\$/g, '').replace(/\s/g, '').replace(/%/g, '');
  if (s === '' || /^(N\/A|NA|-|#N\/A)$/i.test(s)) return null;
  const n = Number(s.replace(/,/g, '.'));
  return isFinite(n) ? n : null;
}
function round(n) { return n === null || n === undefined ? null : Math.round(n); }
function pad(n) { return String(n).padStart(2, '0'); }
function periodoStr(y, m) { return `${y}-${pad(m)}-01`; }

function marcaDesdeTexto(s) {
  const t = norm(s);
  if (t.includes('PEUGEOT')) return 'PEUGEOT';
  if (t.includes('OPEL')) return 'OPEL';
  if (t.includes('CITRO')) return 'CITROEN';
  return null;
}
function marcaDesdeCIT(cit) {
  const t = norm(cit);
  if (t.startsWith('PG')) return 'PEUGEOT';
  if (t.startsWith('OP')) return 'OPEL';
  if (t.startsWith('CT')) return 'CITROEN';
  return null;
}
function periodoDesdeTexto(s) {
  const t = norm(s).replace(/[_\-.]/g, ' ');
  let mes = null, anio = null;
  for (const [k, v] of Object.entries(MESES)) if (new RegExp('\\b' + k + '\\b').test(t)) { mes = v; break; }
  const y = t.match(/\b(20\d{2})\b/);
  if (y) anio = Number(y[1]);
  return mes && anio ? periodoStr(anio, mes) : null;
}

// Busca el mes primero en el nombre del archivo y luego en las carpetas (de la más cercana a la raíz).
// El año se toma del mismo tramo; si no está, de cualquier otro tramo de la ruta.
function periodoDesdeRuta(ruta) {
  if (!ruta) return null;
  const partes = String(ruta).split(/[\\/]/).filter(Boolean).reverse();
  const mesDe = (s) => { const t = ' ' + norm(s).replace(/[_\-.]/g, ' ') + ' '; for (const [k, v] of Object.entries(MESES)) if (new RegExp('\\b' + k + '\\b').test(t)) return v; const n = t.match(/\b(20\d{2})[ ]?(0[1-9]|1[0-2])\b/); return n ? Number(n[2]) : null; };
  const anioDe = (s) => { const y = norm(s).match(/(?:^|\D)(20\d{2})(?:\D|$)/); return y ? Number(y[1]) : null; };
  for (const p of partes) {
    const m = mesDe(p);
    if (m) { const y = anioDe(p) || partes.map(anioDe).find(Boolean); return y ? periodoStr(y, m) : null; }
  }
  return null;
}

// Quita filas repetidas por clave. Si las repetidas traen montos distintos, se informa como conflicto.
function sinRepetidas(rows, clave, etiqueta, quedarseCon) {
  const m = new Map(), repetidas = [];
  for (const r of rows) {
    const k = clave(r);
    if (!m.has(k)) { m.set(k, r); continue; }
    const prev = m.get(k);
    const distinta = JSON.stringify({ ...prev, familia: null, version: null }) !== JSON.stringify({ ...r, familia: null, version: null });
    repetidas.push({ nombre: etiqueta(r), distinta });
    if (quedarseCon === 'ultima') m.set(k, r);
  }
  return { rows: [...m.values()], repetidas };
}

function filas(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
}

/* ---------- Detección de tipo de archivo ---------- */
function detectarArchivo(wb) {
  const names = wb.SheetNames.map(norm);
  if (names.includes('PASAJEROS') && names.includes('COMERCIALES')) return 'lista';
  if (names.includes('PEUGEOT') && names.includes('OPEL') && names.includes('CITROEN')) return 'b2b';
  const r = filas(wb.Sheets[wb.SheetNames[0]]);
  const h = (r[0] || []).map(norm);
  if (h.includes('FOLIO VENTA') && h.includes('CODIGO VERSION')) return 'informe';
  return null;
}

/* ---------- Lista de precios (Peugeot / Opel / Citroën) ---------- */
function parseListaPrecios(wb, filename) {
  const out = [];
  let marca = marcaDesdeTexto(String(filename).split(/[\\/]/).pop()) || marcaDesdeTexto(filename);
  let periodo = periodoDesdeRuta(filename);
  for (const sheetName of wb.SheetNames) {
    const sn = norm(sheetName);
    if (sn !== 'PASAJEROS' && sn !== 'COMERCIALES') continue;
    const seg = sn === 'PASAJEROS' ? 'PASAJEROS' : 'COMERCIALES';
    const R = filas(wb.Sheets[sheetName]);
    // fecha de vigencia (D5 en Peugeot)
    if (!periodo) {
      for (let i = 0; i < Math.min(8, R.length); i++) for (const v of (R[i] || [])) {
        if (v instanceof Date) { periodo = periodoStr(v.getFullYear(), v.getMonth() + 1); }
      }
    }
    const h8i = R.findIndex(r => (r || []).some(v => norm(v) === 'LCDV'));
    if (h8i < 0) continue;
    const h8 = (R[h8i] || []).map(norm), h9 = (R[h8i + 1] || []).map(norm);
    const ncol = Math.max(h8.length, h9.length);
    // grupos de la fila 8
    const grupos = [];
    for (let c = 0; c < ncol; c++) if (h8[c]) grupos.push({ c, t: h8[c] });
    const rango = (pred) => {
      const i = grupos.findIndex(g => pred(g.t));
      if (i < 0) return null;
      return [grupos[i].c, i + 1 < grupos.length ? grupos[i + 1].c : ncol];
    };
    const sub = (rg, pred) => { if (!rg) return -1; for (let c = rg[0]; c < rg[1]; c++) if (pred(h9[c] || '')) return c; return -1; };
    const cLCDV = h8.indexOf('LCDV');
    const cCIT = h8.indexOf('CIT');
    const cVer = h8.findIndex(t => t.startsWith('VEHICULOS'));
    const cKey = h9.findIndex(t => ['KEY MCV', 'COD EDS', 'CODIGO REF', 'CDC REF'].includes(t));
    const rLista = rango(t => t === 'PRECIO LISTA');
    const rTMP = rango(t => t === 'TODO MEDIO DE PAGO');
    const rCC = rango(t => t === 'CREDITO CONVENCIONAL');
    const rCI = rango(t => t === 'COMPRA INTELIGENTE');
    const rB2B = rango(t => t.startsWith('DESCUENTOS B2C'));
    const cMargen = h8.findIndex(t => t.startsWith('MARGEN'));
    const C = {
      listaIva: sub(rLista, t => t === 'CON IVA'), listaNeto: sub(rLista, t => t === 'SIN IVA'),
      tmpIva: sub(rTMP, t => t === 'BONO CON IVA'), tmpNeto: sub(rTMP, t => t === 'BONO SIN IVA'),
      tmpRed: sub(rTMP, t => t === 'APORTE RED' || t === 'APORTE CE'), tmpStl: sub(rTMP, t => t === 'APORTE STELLANTIS'),
      ccIva: sub(rCC, t => t === 'BONO CON IVA'), ccNeto: sub(rCC, t => t === 'BONO SIN IVA'), ccFin: sub(rCC, t => t === 'APORTE FINANCIERA'),
      ciIva: sub(rCI, t => t === 'BONO CON IVA'), ciNeto: sub(rCI, t => t === 'BONO SIN IVA'), ciFin: sub(rCI, t => t === 'APORTE FINANCIERA'),
      t1: sub(rB2B, t => t === 'T1'), t2: sub(rB2B, t => t === 'T2'), gc: sub(rB2B, t => t === 'GC'),
    };
    const g = (row, c) => (c >= 0 ? row[c] : null);
    let familia = null;
    for (let i = h8i + 2; i < R.length; i++) {
      const row = R[i] || [];
      const lcdv = g(row, cLCDV);
      const ver = g(row, cVer);
      if (!lcdv || String(lcdv).trim().length < 10 || String(lcdv).startsWith('#')) {
        if (ver && String(ver).length < 40 && !String(ver).startsWith('(')) familia = String(ver).trim();
        continue;
      }
      const cit = g(row, cCIT);
      const mk = marcaDesdeCIT(cit) || marca;
      if (!marca) marca = mk;
      let listaNeto, listaIva, bTmp, bCc, bCi, red, stl, fCc, fCi;
      if (seg === 'PASAJEROS') {
        // base con IVA
        listaIva = num(g(row, C.listaIva));
        listaNeto = num(g(row, C.listaNeto)) ?? (listaIva !== null ? listaIva / IVA : null);
        bTmp = num(g(row, C.tmpIva)); bCc = num(g(row, C.ccIva)); bCi = num(g(row, C.ciIva));
        red = num(g(row, C.tmpRed)); stl = num(g(row, C.tmpStl));
        fCc = num(g(row, C.ccFin)); fCi = num(g(row, C.ciFin));
      } else {
        // base sin IVA -> se lleva todo a con IVA
        const x = (c) => { const v = num(g(row, c)); return v === null ? null : v * IVA; };
        listaNeto = num(g(row, C.listaNeto));
        listaIva = listaNeto !== null ? listaNeto * IVA : num(g(row, C.listaIva));
        bTmp = x(C.tmpNeto); bCc = x(C.ccNeto); bCi = x(C.ciNeto);
        red = x(C.tmpRed); stl = x(C.tmpStl); fCc = x(C.ccFin); fCi = x(C.ciFin);
      }
      // Reconstrucción si faltan aportes ocultos: Red = min(bono TMP/2, 2,5% lista)
      if (red === null && bTmp !== null && listaIva !== null) red = Math.min(bTmp / 2, listaIva * 0.025);
      if (stl === null && bTmp !== null && red !== null) stl = bTmp - red;
      if (fCc === null && bCc !== null && bTmp !== null) fCc = bCc - bTmp;
      if (fCi === null && bCi !== null && bTmp !== null) fCi = bCi - bTmp;
      if (bCi === 0) { bCi = null; fCi = null; }
      out.push({
        marca: mk, segmento: seg, key_mcv: g(row, cKey) ? String(g(row, cKey)).trim() : null,
        lcdv: String(lcdv).trim(), cit: cit ? String(cit).trim() : null, familia,
        version: ver ? String(ver).replace(/\s+/g, ' ').trim() : null,
        precio_lista_neto: round(listaNeto), precio_lista_iva: round(listaIva),
        bono_tmp_iva: round(bTmp), bono_cc_iva: round(bCc), bono_ci_iva: round(bCi),
        aporte_red_iva: round(red), aporte_stl_iva: round(stl),
        aporte_fin_cc_iva: round(fCc), aporte_fin_ci_iva: round(fCi),
        margen_red: num(g(row, cMargen)),
        b2b_t1: num(g(row, C.t1)), b2b_t2: num(g(row, C.t2)), b2b_gc: num(g(row, C.gc)),
      });
    }
  }
  const d = sinRepetidas(out, r => r.lcdv, r => `${r.version || r.key_mcv} (${r.lcdv})`, 'primera');
  return { tipo: 'lista', marca, periodo, filas: d.rows, repetidas: d.repetidas };
}

/* ---------- Lista de descuentos B2B (flotas) ---------- */
function parseB2B(wb, filename) {
  let periodo = periodoDesdeRuta(filename);
  const out = [];
  for (const sheetName of wb.SheetNames) {
    const marca = marcaDesdeTexto(sheetName);
    if (!marca) continue;
    const R = filas(wb.Sheets[sheetName]);
    if (!periodo) {
      for (let i = 0; i < 5; i++) for (const v of (R[i] || [])) { const p = periodoDesdeTexto(v); if (p) { periodo = p; break; } }
    }
    const hi = R.findIndex(r => (r || []).some(v => norm(v).startsWith('TRAMO 1')));
    if (hi < 0) continue;
    const h = (R[hi] || []).map(norm);
    // primera ocurrencia = Concesionario Professional
    const cT1 = h.findIndex(t => t.startsWith('TRAMO 1'));
    const cT2 = h.findIndex(t => t.startsWith('TRAMO 2'));
    const cGC = h.findIndex(t => t.startsWith('GRANDES CUENTAS'));
    const cE = 4, cA = 0;
    for (let i = hi + 1; i < R.length; i++) {
      const row = R[i] || [];
      if (norm(row[cE]) !== 'TOTAL' || !row[cA]) continue;
      const s = R[i + 1] || [], r = R[i + 2] || [];
      out.push({
        marca, key_mcv: String(row[cA]).trim(),
        t1: num(row[cT1]), t2: num(row[cT2]), gc: num(row[cGC]),
        t1_stl: num(s[cT1]), t2_stl: num(s[cT2]), gc_stl: num(s[cGC]),
        t1_red: num(r[cT1]), t2_red: num(r[cT2]), gc_red: num(r[cGC]),
      });
    }
  }
  const d = sinRepetidas(out, r => r.marca + '|' + norm(r.key_mcv), r => `${r.marca} · ${r.key_mcv}`, 'primera');
  return { tipo: 'b2b', periodo, filas: d.rows, repetidas: d.repetidas };
}

/* ---------- Informe de ventas ---------- */
function fechaDesde(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return v;
  const m = String(v).trim().match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
  const d = new Date(v); return isNaN(d) ? null : d;
}
function isoFecha(d) { return d ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` : null; }

function parseInforme(wb) {
  const R = filas(wb.Sheets[wb.SheetNames[0]]);
  const H = (R[0] || []).map(v => (v === null ? '' : String(v).trim()));
  const idx = (name) => H.findIndex(h => norm(h) === norm(name));
  const col = {};
  ['Sucursal', 'Vendedor', 'Fecha Venta', 'Folio Venta', 'Cliente', 'Rut', 'Digito', 'Marca', 'Modelo', 'Versión', 'Codigo Version',
    'Color', 'Numero Chasis', 'Estado Venta', 'Folio Facturación', 'Fecha Facturación', 'Tipo Venta', 'Precio Lista',
    'Bono Lista Precio', 'Unidad Venta', 'Bono'].forEach(n => (col[n] = idx(n)));
  const faltan = Object.entries(col).filter(([, v]) => v < 0).map(([k]) => k);
  ['Precio Compra', 'Entidad Crédito', 'Margen Crédito'].forEach(n => (col[n] = idx(n))); // opcionales
  const out = [];
  for (let i = 1; i < R.length; i++) {
    const r = R[i] || [];
    const g = (n) => (col[n] >= 0 ? r[col[n]] : null);
    const s = (n) => { const v = g(n); return v === null || v === undefined || v === '' ? null : String(v).replace(/\s+/g, ' ').trim(); };
    const folio = s('Folio Venta');
    if (!folio) continue;
    const ff = fechaDesde(g('Fecha Facturación'));
    const fv = fechaDesde(g('Fecha Venta'));
    out.push({
      folio_venta: folio,
      folio_facturacion: s('Folio Facturación'),
      fecha_facturacion: isoFecha(ff),
      fecha_venta: fv ? fv.toISOString() : null,
      sucursal: s('Sucursal'), vendedor: s('Vendedor'), cliente: s('Cliente'), rut: s('Rut'), digito: s('Digito'),
      marca: s('Marca'), modelo: s('Modelo'), version: s('Versión'), codigo_version: s('Codigo Version'), color: s('Color'),
      chasis: s('Numero Chasis'), tipo_venta: s('Tipo Venta'), estado_venta: s('Estado Venta'),
      unidad_venta: num(g('Unidad Venta')), precio_lista_neto: num(g('Precio Lista')),
      bono_lista_precio: num(g('Bono Lista Precio')), bono_informe: num(g('Bono')),
      precio_compra: num(g('Precio Compra')), entidad_credito: s('Entidad Crédito'), margen_credito: num(g('Margen Crédito')),
    });
  }
  const d = sinRepetidas(out, r => r.folio_venta, r => `Folio ${r.folio_venta}`, 'ultima');
  return { tipo: 'informe', faltan, filas: d.rows, repetidas: d.repetidas };
}

/* ---------- Índices y cálculo ---------- */
function periodoDeVenta(v) {
  const f = v.fecha_facturacion || (v.fecha_venta ? v.fecha_venta.slice(0, 10) : null);
  return f ? f.slice(0, 7) + '-01' : null;
}
function indexarPrecios(precios) {
  const m = new Map();
  for (const p of precios) m.set(p.periodo + '|' + p.lcdv, p);
  return m;
}
function indexarB2B(b2b) {
  const m = new Map();
  for (const b of b2b) {
    const k = b.periodo + '|' + norm(b.marca);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(b);
  }
  return m;
}
function buscarB2B(idxB2B, periodo, marca, key) {
  const lista = idxB2B.get(periodo + '|' + norm(marca));
  if (!lista || !key) return null;
  const k = norm(key);
  let b = lista.find(x => norm(x.key_mcv) === k);
  if (b) return { b, aprox: false };
  // aproximado: prefijo común más largo (p.ej. Traveller 9P usa la fila 8P)
  let best = null, bestLen = 0;
  for (const x of lista) {
    const kx = norm(x.key_mcv); let n = 0;
    while (n < k.length && n < kx.length && k[n] === kx[n]) n++;
    if (n > bestLen) { bestLen = n; best = x; }
  }
  if (best && bestLen >= Math.max(12, Math.floor(k.length * 0.6))) return { b: best, aprox: true };
  return null;
}

const CLASIF = {
  CONTADO: 'Contado',
  CC: 'CC',
  CI: 'CI',
  FLOTA_T1: 'Flota Tramo 1',
  FLOTA_T2: 'Flota Tramo 2',
  FLOTA_GC: 'Flota GC',
  FLOTA_CC: 'Flota CC',
  FLOTA_CI: 'Flota CI',
};
// valores guardados con la versión anterior
const CLASIF_ANTIGUA = { MARCA: 'CONTADO', MARCA_SANT_CC: 'CC', MARCA_SANT_CI: 'CI', SIN_BONO: null };
function clasifGuardada(v) { const x = v.clasificacion; if (!x) return null; if (x in CLASIF) return x; return CLASIF_ANTIGUA[x] ?? null; }

const TIPOS_EXTRA = ['SELL IN', 'Otro'];

// Financiera: primero la Entidad Crédito del informe; si no viene, el Tipo Venta.
function financieraDe(v) {
  const conocidas = [['SANTANDER', 'Santander'], ['FORUM', 'Forum'], ['TANNER', 'Tanner'], ['SCOT', 'Scotiabank'], ['BCI', 'BCI'], ['CHILE', 'Banco de Chile'], ['ESTADO', 'BancoEstado'], ['ITAU', 'Itaú'], ['AUTOFIN', 'Autofin'], ['GLOBAL', 'Global'], ['MAF', 'MAF']];
  for (const fuente of [norm(v.entidad_credito), norm(v.tipo_venta)]) {
    if (!fuente) continue;
    for (const [k, n] of conocidas) if (fuente.includes(k)) return n;
  }
  const ent = (v.entidad_credito || '').trim();
  if (ent) return ent.split(/\s+/).slice(0, 2).join(' ');
  const tv = norm(v.tipo_venta);
  if (!tv || /CONTADO/.test(tv)) return null;
  if (/LEASING/.test(tv)) return 'Leasing';
  if (/FLOTA/.test(tv)) return null;
  return 'Crédito (sin entidad)';
}
// Margen estimado = Unidad Venta − Precio Compra + bonos a cobrar sin IVA
function margenDe(v, c) {
  if (v.unidad_venta === null || v.unidad_venta === undefined || v.precio_compra === null || v.precio_compra === undefined) return null;
  return v.unidad_venta - v.precio_compra + (c.total || 0) / IVA;
}
function extrasDe(v) { const x = v.bonos_extra; if (Array.isArray(x)) return x; if (typeof x === 'string') { try { return JSON.parse(x) || []; } catch (e) { return []; } } return []; }

function formaDesdeTipoVenta(tv, match, tramoDesc) {
  const esFlota = /FLOTA|LEASING/.test(tv);
  if (esFlota && /\bCI\b/.test(tv)) return 'FLOTA_CI';
  if (esFlota && /\bCC\b/.test(tv)) return 'FLOTA_CC';
  if (esFlota) return 'FLOTA_' + (match || tramoDesc || 'T1');
  if (match) return 'FLOTA_' + match;          // venta que trae exactamente un bono flota (p.ej. contado a empresa)
  if (/CONTADO/.test(tv)) return 'CONTADO';
  if (/\bCI\b/.test(tv)) return 'CI';
  return 'CC';
}

function calcularBonos(v, idxP, idxB) {
  const periodo = periodoDeVenta(v);
  const res = { periodo, p: null, b2b: null, cands: {}, auto: null, clasif: null, tramo: null, marca: null, sant: null, flota: null, extra: null, total: null, dif: null, avisos: [] };
  const p = periodo ? idxP.get(periodo + '|' + (v.codigo_version || '')) : null;
  const tv = norm(v.tipo_venta);
  const esSant = /SANTANDER/.test(tv) || /SANTANDER/.test(norm(v.entidad_credito));
  const sumaExtra = () => extrasDe(v).reduce((s, e) => s + (Number(e.monto) || 0), 0) || null;
  if (!p) {
    res.avisos.push(periodo ? `Sin lista de precios de ${MESES_NOMBRE[Number(periodo.slice(5, 7))]} ${periodo.slice(0, 4)} para esta versión` : 'Sin fecha');
    res.auto = formaDesdeTipoVenta(tv, null, null);
    res.clasif = clasifGuardada(v) || res.auto;
    res.extra = sumaExtra();
    return res;
  }
  res.p = p;
  const neto = p.precio_lista_neto;
  const fb = buscarB2B(idxB, periodo, p.marca, p.key_mcv);
  if (fb) {
    res.b2b = fb.b;
    if (fb.aprox) res.avisos.push(`Tabla B2B tomada de versión similar (${fb.b.key_mcv})`);
    for (const t of ['t1', 't2', 'gc']) if (fb.b[t + '_stl'] !== null && fb.b[t + '_stl'] !== undefined) res.cands[t.toUpperCase()] = Math.round(Math.round(fb.b[t + '_stl'] * neto) * 119 / 100);
  }
  const bi = v.bono_informe;
  let match = null;
  if (bi !== null && bi !== undefined) for (const [t, val] of Object.entries(res.cands)) if (Math.abs(val - bi) <= TOL) { match = t; break; }
  let tramoDesc = null;
  if (res.b2b && v.unidad_venta && neto) {
    const d = 1 - v.unidad_venta / neto; let best = 1e9;
    for (const t of ['t1', 't2', 'gc']) { const x = res.b2b[t]; if (x !== null && x !== undefined && Math.abs(d - x) < best) { best = Math.abs(d - x); tramoDesc = t.toUpperCase(); } }
  }
  res.auto = formaDesdeTipoVenta(tv, match, tramoDesc);
  res.clasif = clasifGuardada(v) || res.auto;
  const c = res.clasif;
  if (c === 'CONTADO' || c === 'CC' || c === 'CI') {
    res.marca = p.aporte_stl_iva;
    if (esSant && c === 'CC') { res.sant = p.aporte_fin_cc_iva; if (res.sant === null) res.avisos.push('La lista no trae aporte financiera CC'); }
    if (esSant && c === 'CI') { res.sant = p.aporte_fin_ci_iva; if (res.sant === null) res.avisos.push('Esta versión no tiene Compra Inteligente'); }
  } else {
    // Flota T1 / T2 / GC usan su tramo; Flota CC / CI usan el tramo detectado (sin aporte financiera)
    res.tramo = (c === 'FLOTA_CC' || c === 'FLOTA_CI') ? (match || tramoDesc || 'T1') : c.slice(6);
    res.flota = res.cands[res.tramo] ?? null;
    if (res.flota === null) res.avisos.push('Versión sin tabla B2B: monto flota no calculable');
  }
  res.extra = sumaExtra();
  res.total = (res.marca || 0) + (res.sant || 0) + (res.flota || 0) + (res.extra || 0);
  if (bi !== null && bi !== undefined) res.dif = Math.round(bi - res.total);
  return res;
}

if (typeof module !== 'undefined') module.exports = { financieraDe, margenDe, extrasDe, TIPOS_EXTRA, periodoDesdeRuta, norm, num, detectarArchivo, parseListaPrecios, parseB2B, parseInforme, indexarPrecios, indexarB2B, calcularBonos, periodoDeVenta, CLASIF, MESES_NOMBRE, TOL };
