/**
 * ISL · Stock de Almacenes — backend (Apps Script VINCULADO al Google Sheet de stock)
 *
 * Hojas esperadas en el Sheet: Movimientos, Catalogo, Nodos (las del libro Stock_Almacenes_ISL.xlsx).
 * Propiedades del script (Configuración del proyecto → Propiedades del script):
 *   APP_PIN            PIN para escribir desde la app (ingresos, ajustes, altas). Sin PIN no se escribe.
 *   REMITOS_SHEET_ID   ID del Google Sheet de Remitos Digitales (hojas "Remitos" y "NODOS").
 *
 * Cómo se mueve el stock:
 *   • Los remitos de tipo Stock se leen de la hoja "Remitos" y se pasan a Movimientos (sincronizarRemitos_),
 *     cada 5 minutos (instalarTrigger) y a pedido (accion=sincronizar). Es idempotente por NroDocumento.
 *   • La app solo registra ingresos por compra y ajustes de inventario.
 *   • Saldo de un nodo = Σ NodoDestino − Σ NodoOrigen.
 */

const HOJA = { MOV: 'Movimientos', CAT: 'Catalogo', NOD: 'Nodos' };
const ESP = { 'AJUSTE-INV': 'Inventario / ajuste', 'PROVEEDOR': 'Proveedor' };
const C = { ID: 0, FECHA: 1, HORA: 2, TIPO: 3, COD: 4, DESC: 5, ORI: 6, DES: 7, RESP: 8, RECEP: 9, DOC: 10, OBS: 11, CANT: 12, UNI: 13, VER: 14 };
const NCOL = 15;
const MAX_CATALOGO = 300; // lugares disponibles en la hoja Stock
const PREFIJOS = { 'Cable': 'CAB', 'Material eléctrico': 'ELE', 'Cañería y accesorios': 'CAN', 'Ferretería': 'FER', 'Consumibles': 'CON', 'Otro': 'OTR' };
const EPS = 1e-9;

/* ============ utilidades ============ */
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function num_(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const n = parseFloat(String(v === null || v === undefined ? '' : v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function r3_(n) { return Math.round(n * 1000) / 1000; }
function tz_() { return Session.getScriptTimeZone(); }
function fmtFecha_(v) { return (v instanceof Date) ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : str_(v); }
function fmtHora_(v) { return (v instanceof Date) ? Utilities.formatDate(v, tz_(), 'HH:mm:ss') : str_(v); }
function norm_(s) { return str_(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }
function conLock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ============ lectura de hojas ============ */
function leerMovs_() {
  const sh = ss_().getSheetByName(HOJA.MOV);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, NCOL).getValues().filter(r => str_(r[C.COD]) !== '');
}
function leerCatalogo_() {
  const sh = ss_().getSheetByName(HOJA.CAT);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 6).getValues()
    .filter(r => str_(r[0]) !== '')
    .map(r => ({ codigo: str_(r[0]), descripcion: str_(r[1]), categoria: str_(r[2]), subtipo: str_(r[3]), unidad: str_(r[4]) || 'un', obs: str_(r[5]) }));
}
function mapaCatalogo_() { const m = {}; leerCatalogo_().forEach(c => m[c.codigo] = c); return m; }
function leerAlmacenes_() {
  const sh = ss_().getSheetByName(HOJA.NOD);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 4).getValues()
    .filter(r => str_(r[0]) !== '' && norm_(r[2]).indexOf('almac') === 0)
    .map(r => ({ id: str_(r[0]), nombre: str_(r[1]) || ('Nodo ' + str_(r[0])) }));
}
function leerObras_() {
  const id = prop_('REMITOS_SHEET_ID');
  const out = {};
  if (!id) return out;
  try {
    const sh = SpreadsheetApp.openById(id).getSheetByName('NODOS');
    if (!sh) return out;
    const v = sh.getDataRange().getValues();
    const h = v[0].map(x => norm_(x));
    const iId = h.indexOf('id') >= 0 ? h.indexOf('id') : 0;
    const iNom = h.indexOf('nombre') >= 0 ? h.indexOf('nombre') : 1;
    for (let i = 1; i < v.length; i++) if (str_(v[i][iId]) !== '') out[str_(v[i][iId])] = str_(v[i][iNom]);
  } catch (e) { /* sin nombres de obras: se muestran los IDs */ }
  return out;
}
function nombreNodo_(id, alm, obras) {
  id = str_(id);
  if (ESP[id]) return ESP[id];
  const a = alm.find(x => x.id === id);
  if (a) return a.nombre;
  return obras[id] || ('Nodo ' + id);
}

/* ============ cálculo de saldos ============ */
function calcularSaldos_(movs) {
  const s = {}, ing = {}, ver = {};
  movs.forEach(r => {
    const cod = str_(r[C.COD]), o = str_(r[C.ORI]), d = str_(r[C.DES]), q = num_(r[C.CANT]);
    s[cod] = s[cod] || {};
    if (o) s[cod][o] = (s[cod][o] || 0) - q;
    if (d) s[cod][d] = (s[cod][d] || 0) + q;
    ing[cod] = (ing[cod] || 0) + (ESP[o] ? q : 0) - (ESP[d] ? q : 0);
    if (str_(r[C.VER]).toUpperCase() === 'SI') ver[cod] = (ver[cod] || 0) + 1;
  });
  return { s: s, ing: ing, ver: ver };
}

/* ============ escritura ============ */
function fila_(o) {
  const f = new Array(NCOL).fill('');
  f[C.FECHA] = o.fecha || new Date();
  f[C.HORA] = o.hora || '';
  f[C.TIPO] = o.tipo || '';
  f[C.COD] = o.cod;
  f[C.DESC] = o.desc || '';
  f[C.ORI] = o.ori;
  f[C.DES] = o.des;
  f[C.RESP] = o.resp || '';
  f[C.RECEP] = o.recep || '';
  f[C.DOC] = o.doc || '';
  f[C.OBS] = o.obs || '';
  f[C.CANT] = o.cant;
  f[C.UNI] = o.uni || '';
  f[C.VER] = o.ver || '';
  return f;
}
function appendMovs_(filas) {
  const sh = ss_().getSheetByName(HOJA.MOV);
  const last = sh.getLastRow();
  let maxId = 0;
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => { const n = num_(r[0]); if (n > maxId) maxId = n; });
  filas.forEach((f, i) => { f[C.ID] = maxId + i + 1; });
  sh.getRange(last + 1, 1, filas.length, NCOL).setValues(filas);
  sh.getRange(last + 1, 2, filas.length, 1).setNumberFormat('dd/mm/yyyy');
}

/* ============ seguridad ============ */
function autorizar_(p) {
  const pin = prop_('APP_PIN');
  if (!pin) throw new Error('El servidor no tiene APP_PIN configurado: no se puede escribir.');
  if (str_(p.pin) !== pin) throw new Error('PIN incorrecto.');
}

/* ============ enrutador ============ */
function doGet(e) { return despachar_((e && e.parameter) || {}, false); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'json_invalido', detalle: 'El cuerpo no es JSON válido.' }); }
  return despachar_(body, true);
}
function despachar_(p, esPost) {
  try {
    const a = str_(p.accion);
    switch (a) {
      case 'catalogo': return json_(accionCatalogo_());
      case 'stock': return json_(accionStock_());
      case 'stock_por_nodo': return json_(accionStockPorNodo_(p.nodo));
      case 'kardex': return json_(accionKardex_(p));
      case 'movimientos': return json_(accionMovimientos_(p));
      case 'sincronizar': return json_(sincronizarRemitos_());
      case 'ingreso':
        if (!esPost) throw new Error('Esta acción requiere POST.');
        autorizar_(p); return json_(accionIngreso_(p));
      case 'alta_item':
        if (!esPost) throw new Error('Esta acción requiere POST.');
        autorizar_(p); return json_(accionAltaItem_(p));
      case 'verificar_pin':
        autorizar_(p); return json_({ ok: true });
      default: return json_({ error: 'accion_desconocida', detalle: 'Acción no reconocida: ' + a });
    }
  } catch (err) {
    return json_({ error: 'error', detalle: String((err && err.message) || err) });
  }
}

/* ============ acciones de lectura ============ */
function accionCatalogo_() {
  const obras = leerObras_();
  return {
    items: leerCatalogo_(),
    almacenes: leerAlmacenes_(),
    categorias: Object.keys(PREFIJOS),
    obras: Object.keys(obras).map(id => ({ id: id, nombre: obras[id] }))
  };
}
function accionStock_() {
  const cat = leerCatalogo_(), alm = leerAlmacenes_();
  const sd = calcularSaldos_(leerMovs_());
  const filas = cat.map(c => {
    const sc = sd.s[c.codigo] || {};
    const saldos = {}; let total = 0;
    alm.forEach(a => { const v = r3_(sc[a.id] || 0); saldos[a.id] = v; total += v; });
    total = r3_(total);
    const ingresado = r3_(sd.ing[c.codigo] || 0);
    return {
      codigo: c.codigo, descripcion: c.descripcion, categoria: c.categoria, subtipo: c.subtipo, unidad: c.unidad,
      saldos: saldos, total: total, ingresado: ingresado, enObras: r3_(ingresado - total), verificar: sd.ver[c.codigo] || 0
    };
  });
  return { almacenes: alm, filas: filas, actualizado: new Date().toISOString() };
}
function accionStockPorNodo_(nodo) {
  nodo = str_(nodo);
  if (!nodo) throw new Error('Falta el nodo.');
  const cat = leerCatalogo_();
  const sd = calcularSaldos_(leerMovs_());
  const items = [];
  cat.forEach(c => {
    const q = r3_(((sd.s[c.codigo] || {})[nodo]) || 0);
    if (q > 0.0005) items.push({ codigo: c.codigo, descripcion: c.descripcion, categoria: c.categoria, subtipo: c.subtipo, unidad: c.unidad, saldo: q });
  });
  return { nodo: nodo, items: items, detalle: items.length ? '' : 'No hay stock cargado en ese nodo.' };
}
function movAObjeto_(r, alm, obras, nodo) {
  const o = str_(r[C.ORI]), d = str_(r[C.DES]), q = num_(r[C.CANT]);
  let efecto = null;
  if (nodo) efecto = (d === nodo ? q : 0) - (o === nodo ? q : 0);
  return {
    id: r[C.ID], fecha: fmtFecha_(r[C.FECHA]), hora: fmtHora_(r[C.HORA]), tipo: str_(r[C.TIPO]),
    codigo: str_(r[C.COD]), descripcion: str_(r[C.DESC]),
    origen: o, origenNombre: nombreNodo_(o, alm, obras), destino: d, destinoNombre: nombreNodo_(d, alm, obras),
    cantidad: r3_(q), unidad: str_(r[C.UNI]), doc: str_(r[C.DOC]), resp: str_(r[C.RESP]), recep: str_(r[C.RECEP]),
    obs: str_(r[C.OBS]), verificar: str_(r[C.VER]).toUpperCase() === 'SI', efecto: efecto === null ? null : r3_(efecto)
  };
}
function accionKardex_(p) {
  const item = str_(p.item);
  if (!item) throw new Error('Falta el ítem.');
  const nodo = str_(p.nodo);
  const alm = leerAlmacenes_(), obras = leerObras_();
  const limit = Math.min(Math.max(parseInt(p.limit, 10) || 300, 1), 1000);
  const movs = leerMovs_().filter(r => str_(r[C.COD]) === item &&
    (!nodo || str_(r[C.ORI]) === nodo || str_(r[C.DES]) === nodo));
  const out = movs.map(r => movAObjeto_(r, alm, obras, nodo)).reverse().slice(0, limit);
  const c = mapaCatalogo_()[item];
  return { item: item, descripcion: c ? c.descripcion : '', unidad: c ? c.unidad : '', movimientos: out };
}
function accionMovimientos_(p) {
  const alm = leerAlmacenes_(), obras = leerObras_();
  const limit = Math.min(Math.max(parseInt(p.limit, 10) || 200, 1), 1000);
  const q = norm_(p.q), item = str_(p.item), nodo = str_(p.nodo);
  const out = [];
  const movs = leerMovs_();
  for (let i = movs.length - 1; i >= 0 && out.length < limit; i--) {
    const r = movs[i];
    if (item && str_(r[C.COD]) !== item) continue;
    if (nodo && str_(r[C.ORI]) !== nodo && str_(r[C.DES]) !== nodo) continue;
    const o = movAObjeto_(r, alm, obras, nodo);
    if (q) {
      const txt = norm_([o.codigo, o.descripcion, o.doc, o.resp, o.recep, o.origenNombre, o.destinoNombre, o.tipo, o.obs].join(' '));
      if (q.split(/\s+/).some(t => t && txt.indexOf(t) < 0)) continue;
    }
    out.push(o);
  }
  return { movimientos: out };
}

/* ============ ingresos / ajustes (desde la app) ============ */
function accionIngreso_(p) {
  return conLock_(function () {
    const modo = str_(p.modo);
    const alm = leerAlmacenes_();
    const a = alm.find(x => x.id === str_(p.almacen));
    if (!a) throw new Error('Elegí un almacén válido.');
    const items = (p.items || []).map(i => ({ codigo: str_(i.codigo), cantidad: num_(i.cantidad) })).filter(i => i.codigo && i.cantidad > 0);
    if (!items.length) throw new Error('Cargá al menos un ítem con cantidad mayor a 0.');
    const cat = mapaCatalogo_();
    items.forEach(i => { if (!cat[i.codigo]) throw new Error('El ítem ' + i.codigo + ' no está en el catálogo.'); });
    const obs = str_(p.obs);
    if (modo !== 'compra' && !obs) throw new Error('Los ajustes requieren un motivo.');
    let tipo, ori, des, pref;
    if (modo === 'compra') { tipo = 'Ingreso por compra'; ori = 'PROVEEDOR'; des = a.id; pref = 'ING'; }
    else if (modo === 'ajuste_mas') { tipo = 'Ajuste (+)'; ori = 'AJUSTE-INV'; des = a.id; pref = 'AJ'; }
    else if (modo === 'ajuste_menos') {
      tipo = 'Ajuste (-)'; ori = a.id; des = 'AJUSTE-INV'; pref = 'AJ';
      const s = calcularSaldos_(leerMovs_()).s, req = {};
      items.forEach(i => { req[i.codigo] = (req[i.codigo] || 0) + i.cantidad; });
      Object.keys(req).forEach(cod => {
        const disp = ((s[cod] || {})[a.id]) || 0;
        if (req[cod] > disp + EPS) throw new Error('Saldo insuficiente de ' + cod + ' en ' + a.nombre + ': hay ' + r3_(disp) + ' ' + cat[cod].unidad + '.');
      });
    } else throw new Error('Tipo de movimiento inválido.');
    const ahora = new Date();
    const doc = str_(p.nrodoc) || (pref + '-' + Utilities.formatDate(ahora, tz_(), 'yyyyMMdd-HHmmss'));
    const filas = items.map(i => fila_({
      fecha: ahora, hora: fmtHora_(ahora), tipo: tipo, cod: i.codigo, desc: cat[i.codigo].descripcion,
      ori: ori, des: des, resp: str_(p.responsable), doc: doc, obs: obs, cant: i.cantidad, uni: cat[i.codigo].unidad
    }));
    appendMovs_(filas);
    return { ok: true, doc: doc, registrados: filas.length };
  });
}

/* ============ alta de ítems en el catálogo ============ */
function accionAltaItem_(p) {
  return conLock_(function () {
    const desc = str_(p.descripcion), cat = str_(p.categoria), uni = str_(p.unidad) || 'un';
    if (!desc) throw new Error('Falta la descripción.');
    if (!PREFIJOS[cat]) throw new Error('Categoría inválida.');
    const actuales = leerCatalogo_();
    if (actuales.length >= MAX_CATALOGO) throw new Error('El catálogo está lleno (' + MAX_CATALOGO + ' ítems en la planilla de stock).');
    const dup = actuales.find(c => norm_(c.descripcion) === norm_(desc) && norm_(c.unidad) === norm_(uni));
    if (dup) throw new Error('Ya existe ese ítem: ' + dup.codigo + ' · ' + dup.descripcion);
    const pref = PREFIJOS[cat];
    let max = 0;
    actuales.forEach(c => { const m = /^([A-Z]+)-(\d+)$/.exec(c.codigo); if (m && m[1] === pref) max = Math.max(max, parseInt(m[2], 10)); });
    const codigo = pref + '-' + ('000' + (max + 1)).slice(-3);
    const sh = ss_().getSheetByName(HOJA.CAT);
    const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
    let fila = col.length + 1;
    for (let i = 1; i < col.length; i++) { if (str_(col[i][0]) === '') { fila = i + 1; break; } }
    sh.getRange(fila, 1, 1, 6).setValues([[codigo, desc, cat, str_(p.subtipo), uni, str_(p.obs)]]);
    return { ok: true, item: { codigo: codigo, descripcion: desc, categoria: cat, subtipo: str_(p.subtipo), unidad: uni } };
  });
}

/* ============ sincronización con Remitos Digitales ============ */
function sincronizarRemitos_() {
  const rid = prop_('REMITOS_SHEET_ID');
  if (!rid) throw new Error('Falta REMITOS_SHEET_ID en las propiedades del script.');
  return conLock_(function () {
    const rsh = SpreadsheetApp.openById(rid).getSheetByName('Remitos');
    if (!rsh) throw new Error('No se encontró la hoja "Remitos" en el Sheet de remitos.');
    const data = rsh.getDataRange().getValues();
    const head = data[0].map(h => str_(h));
    const iId = head.indexOf('RemitoID'), iF = head.indexOf('Fecha'), iO = head.indexOf('ObraOrigen'),
      iD = head.indexOf('ObraDestino'), iE = head.indexOf('Entrega'), iR = head.indexOf('Recibe'), iP = head.indexOf('PayloadJSON');
    if ([iId, iO, iD, iP].some(i => i < 0)) throw new Error('La hoja Remitos no tiene las columnas esperadas.');
    const movs = leerMovs_();
    const docs = {}; movs.forEach(r => { docs[str_(r[C.DOC])] = true; });
    const cat = mapaCatalogo_();
    const esAlm = {}; leerAlmacenes_().forEach(a => { esAlm[a.id] = true; });
    const s = calcularSaldos_(movs).s;
    const marca = /"stock"\s*:\s*true/;
    const nuevas = [], procesados = [], alertas = [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const id = str_(row[iId]), pj = str_(row[iP]);
      if (!id || docs[id] || !marca.test(pj)) continue;
      let pay;
      try { pay = JSON.parse(pj); } catch (e) { alertas.push(id + ': payload ilegible'); continue; }
      const ori = str_(row[iO]), des = str_(row[iD]);
      const f = (iF >= 0 && row[iF] instanceof Date) ? row[iF] : new Date();
      const tipo = (esAlm[ori] && esAlm[des]) ? 'Transferencia entre almacenes' : (esAlm[des] ? 'Devolución a almacén' : 'Envío a obra');
      let n = 0;
      (pay.items || []).forEach(it => {
        if (!(it.extra && it.extra.stock === true)) return;
        const cod = str_(it.codigo), q = num_(it.cantidad);
        if (!cod || q <= 0) { alertas.push(id + ': ítem inválido (' + cod + ')'); return; }
        const c = cat[cod];
        let ver = '', obs = '';
        if (!c) { ver = 'SI'; obs += 'Ítem fuera del catálogo. '; }
        const disp = ((s[cod] || {})[ori]) || 0;
        if (q > disp + EPS) {
          ver = 'SI';
          obs += esAlm[ori] ? ('Saldo insuficiente al registrar (había ' + r3_(disp) + '). ') : 'El nodo de origen no tenía saldo registrado. ';
        }
        s[cod] = s[cod] || {};
        s[cod][ori] = (s[cod][ori] || 0) - q;
        s[cod][des] = (s[cod][des] || 0) + q;
        nuevas.push(fila_({
          fecha: f, hora: fmtHora_(f), tipo: tipo, cod: cod, desc: c ? c.descripcion : str_(it.descripcion),
          ori: ori, des: des, resp: str_(row[iE]), recep: str_(row[iR]), doc: id, obs: obs.trim(),
          cant: q, uni: c ? c.unidad : str_(it.unidad), ver: ver
        }));
        n++;
      });
      if (n > 0) procesados.push(id);
    }
    if (nuevas.length) appendMovs_(nuevas);
    return { ok: true, remitos: procesados, movimientos: nuevas.length, alertas: alertas };
  });
}

function sincronizarPeriodico() {
  try { sincronizarRemitos_(); } catch (e) { console.error(e); }
}

/** Ejecutar UNA vez desde el editor para que los remitos se pasen al stock cada 5 minutos. */
function instalarTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sincronizarPeriodico')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sincronizarPeriodico').timeBased().everyMinutes(5).create();
}
