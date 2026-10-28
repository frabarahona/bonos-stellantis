// API de solo lectura para Power BI. Uso: /api/datos?tabla=ventas&key=CLAVE  (tabla: ventas | pagos | precios | metas | pendientes; &formato=json opcional)
global.XLSX = global.XLSX || {};
const C = require('../src/core.js');
const URL_SB = 'https://prnrnsgkdjfotkanfwul.supabase.co';
const ANON = 'sb_publishable_n5t3kpj1IXdWKYgoEgS14Q_3RrqNyI8';
const IVA = 1.19;

async function token() {
  const r = await fetch(`${URL_SB}/auth/v1/token?grant_type=password`, { method: 'POST', headers: { apikey: ANON, 'Content-Type': 'application/json' }, body: JSON.stringify({ email: process.env.SB_EMAIL, password: process.env.SB_PASSWORD }) });
  const j = await r.json(); if (!j.access_token) throw new Error('No se pudo iniciar sesión en la base'); return j.access_token;
}
async function todo(tk, tabla) {
  const out = []; let desde = 0;
  while (true) {
    const r = await fetch(`${URL_SB}/rest/v1/${tabla}?select=*`, { headers: { apikey: ANON, Authorization: 'Bearer ' + tk, Range: `${desde}-${desde + 999}` } });
    if (!r.ok) { if (tabla === 'bonos_pagados' || tabla === 'metas') return out; throw new Error(tabla + ': ' + r.status); }
    const d = await r.json(); out.push(...d); if (d.length < 1000) break; desde += 1000;
  }
  return out;
}
const numerizar = (o) => { const r = { ...o }; for (const [k, v] of Object.entries(o)) if (typeof v === 'string' && /^-?\d+(\.\d+)?$/.test(v) && !['folio_venta', 'folio_facturacion', 'rut', 'digito', 'lcdv', 'key_mcv', 'archivo', 'tipo'].includes(k)) r[k] = Number(v); return r; };
const vinLimpio = (s) => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const claves = (s) => { const t = vinLimpio(s); if (!t) return []; return t.length >= 8 ? ['F:' + t, 'S:' + t.slice(-8)] : ['F:' + t]; };
const buscar = (m, s) => { const t = vinLimpio(s); if (!t) return []; return m.get('F:' + t) || (t.length >= 8 ? m.get('S:' + t.slice(-8)) || [] : []); };
const indexar = (xs, campo) => { const m = new Map(); for (const x of xs) for (const k of claves(x[campo])) { if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; };
function estado(xs, c, fin) {
  const ref = fin ? (c.sant || 0) : (c.marca || 0) + (c.flota || 0);
  if (!xs.length) return ref ? 'PENDIENTE' : '';
  const tot = xs.reduce((s, x) => s + (x.aporte || 0), 0);
  return C.calzaPago(tot, c, fin) ? 'PAGADO - CUADRA' : 'PAGADO - NO CUADRA';
}
const r0 = (n) => (n === null || n === undefined || isNaN(n)) ? null : Math.round(n);
const VACIAS = { metas: ['periodo', 'marca', 'unidades'], pagos: ['folio_carta', 'tipo', 'marca', 'fecha', 'semana', 'vin', 'modelo', 'version', 'canal', 'estado', 'aporte_sin_iva', 'folio_venta', 'cliente'] };
function csv(rows, tabla) {
  const cols = rows.length ? Object.keys(rows[0]) : (VACIAS[tabla] || []);
  const q = (v) => v === null || v === undefined ? '' : /[",\n;]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v);
  return '﻿' + [cols.join(','), ...rows.map(r => cols.map(c => q(r[c])).join(','))].join('\n');
}

module.exports = async (req, res) => {
  try {
    if (!process.env.API_KEY || req.query.key !== process.env.API_KEY) { res.status(401).send('Clave inválida'); return; }
    const tabla = String(req.query.tabla || 'ventas').toLowerCase();
    const tk = await token();
    let rows;
    if (tabla === 'precios') {
      rows = (await todo(tk, 'precios')).map(numerizar).map(p => ({ periodo: p.periodo, marca: p.marca, segmento: p.segmento, familia: p.familia, version: p.version, cit: p.cit, lcdv: p.lcdv, lista_iva: p.precio_lista_iva, lista_neto: p.precio_lista_neto, bono_tmp: p.bono_tmp_iva, bono_cc: p.bono_cc_iva, bono_ci: p.bono_ci_iva, aporte_red: p.aporte_red_iva, aporte_stellantis: p.aporte_stl_iva, financiera_cc: p.aporte_fin_cc_iva, financiera_ci: p.aporte_fin_ci_iva }));
    } else if (tabla === 'metas') {
      rows = (await todo(tk, 'metas')).map(m => ({ periodo: m.periodo, marca: m.marca, unidades: Number(m.unidades) }));
    } else {
      const [p, b, v, pg] = await Promise.all([todo(tk, 'precios'), todo(tk, 'descuentos_b2b'), todo(tk, 'ventas_bonos'), todo(tk, 'bonos_pagados')]);
      const idxP = C.indexarPrecios(p.map(numerizar)), idxB = C.indexarB2B(b.map(numerizar));
      const ventas = v.map(numerizar), pagos = pg.map(x => ({ ...x, aporte: Number(x.aporte), diferencia: Number(x.diferencia || 0) }));
      const pagVin = indexar(pagos, 'vin'), venVin = indexar(ventas, 'chasis');
      const calc = new Map(ventas.map(x => [x.folio_venta, C.calcularBonos(x, idxP, idxB)]));
      if (tabla === 'pagos') {
        rows = pagos.map(x => { const vv = buscar(venVin, x.vin)[0]; return { folio_carta: x.folio, tipo: C.esPagoFinanciamiento(x) ? 'Financiamiento' : 'Aporte marca', marca: x.marca, fecha: x.fecha, semana: x.semana, vin: x.vin, modelo: x.modelo, version: x.version, canal: String(x.canal || '').replace(/\s*\$\s*$/, '') || null, estado: x.estado, aporte_sin_iva: x.aporte, folio_venta: vv ? vv.folio_venta : null, cliente: vv ? vv.cliente : null }; });
      } else {
        rows = ventas.map(x => {
          const c = calc.get(x.folio_venta), ps = buscar(pagVin, x.chasis), pm = ps.filter(y => !C.esPagoFinanciamiento(y)), pf = ps.filter(C.esPagoFinanciamiento);
          return {
            folio_venta: x.folio_venta, periodo: C.periodoDeVenta(x), fecha_facturacion: x.fecha_facturacion, sucursal: x.sucursal, vendedor: x.vendedor, cliente: x.cliente,
            rut: x.rut ? `${x.rut}-${x.digito ?? ''}` : null, marca: x.marca, modelo: x.modelo, version: x.version, chasis: x.chasis, tipo_venta: x.tipo_venta, financiera: C.financieraDe(x),
            forma_pago: C.CLASIF[c.clasif] || c.clasif, tramo_flota: c.tramo || null, unidad_venta_neto: x.unidad_venta, total_con_iva: r0(x.unidad_venta * IVA),
            bono_informe: x.bono_informe, bono_marca_tmp: r0(c.marca), bono_santander: r0(c.sant), bono_flota: r0(c.flota), bonos_adicionales: r0(c.extra), bonos_total: r0(c.total), diferencia: r0(c.dif),
            marcado_pagado_marca: c.marca ? !!(x.pagado_marca || pm.length) : null, marcado_pagado_santander: c.sant ? !!(x.pagado_santander || pf.length) : null, marcado_pagado_flota: c.flota ? !!(x.pagado_flota || pm.length) : null,
            pagado_marca_sin_iva: pm.reduce((s, y) => s + y.aporte, 0) || null, estado_pago_marca: estado(pm, c, false),
            pagado_financ_sin_iva: pf.reduce((s, y) => s + y.aporte, 0) || null, estado_pago_financ: estado(pf, c, true),
            margen: r0(C.margenDe(x, c)), sin_lista: !c.p, observacion: x.observacion || null,
          };
        });
        if (tabla === 'pendientes') rows = rows.filter(r => r.estado_pago_marca === 'PENDIENTE' || r.estado_pago_financ === 'PENDIENTE');
      }
    }
    res.setHeader('Cache-Control', 'no-store');
    if (String(req.query.formato) === 'json') { res.json(rows); return; }
    res.setHeader('Content-Type', 'text/csv; charset=utf-8'); res.send(csv(rows, tabla));
  } catch (e) { res.status(500).send('Error: ' + e.message); }
};
