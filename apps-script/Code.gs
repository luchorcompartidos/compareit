/**
 * CompareIt — backend en Google Apps Script (script vinculado a la Google Sheet).
 *
 * Instalación:
 * 1. Abrí la Google Sheet > Extensiones > Apps Script.
 * 2. Pegá este archivo como Code.gs.
 * 3. Creá un archivo de secuencia de comandos llamado "App" y pegá el contenido de App.gs
 *    (trae el esquema de datos compartido con la app y la interfaz)
 *    (o, como alternativa, un archivo HTML llamado "index" con el contenido de compareit.html).
 * 4. Implementar > Nueva implementación > Aplicación web
 *    (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
 * Dos formas de usar la app conectada:
 * a) Abrir la URL /exec: la app se sirve desde acá.
 * b) Abrir compareit.html en el navegador y conectarla con la URL /exec y la clave API_KEY
 *
 * Datos privados (NO van escritos en este archivo, que se publica en GitHub):
 *   Apps Script > Configuración del proyecto (engranaje) > Propiedades del script > Agregar propiedad:
 *   - API_KEY : la clave que la app envía para leer y escribir (la misma que cargás en la app al conectar).
 *   - SHEET_ID: el ID de la planilla (la parte de su URL entre /d/ y /edit). Opcional si el script está
 *               vinculado a la planilla (creado desde Extensiones > Apps Script).
 *   Se cargan una sola vez: al actualizar este archivo no hay que volver a escribirlas.
 *    (panel de sincronización > Conectar a mi Google Sheet).
 *
 * Estructura de la planilla:
 * - Una pestaña por Grupo (nombre = nombre del grupo): ID | Nombre | Notas | <Atributo [f_id]>... | <Criterio [c_id]>... | Actualizado
 * - Columna nueva escrita a mano en la pestaña de un grupo: pasa a ser un ATRIBUTO (dato descriptivo, no suma puntaje).
 *   Para que sea un CRITERIO, terminá el encabezado con [criterio] (ej.: "Peso (kg) [criterio]").
 * - Columna de atributo borrada a mano: el atributo queda archivado en la app con sus valores guardados.
 * - _Grupos, _Criterios, _Instancias: pestañas de sistema.
 * - La columna "Actualizado" se completa sola al editar a mano (onEdit), para detectar conflictos.
 * - Filas cargadas a mano sin ID reciben uno en la siguiente lectura.
 */

/**
 * Pestañas de sistema armadas desde SCHEMA (definido en App.gs, el mismo esquema que usa la app).
 * Una entidad nueva agregada al esquema obtiene su pestaña y sus columnas sola.
 */
function sys_() {
  const out = {};
  Object.keys(SCHEMA).forEach(t => {
    const s = SCHEMA[t];
    if (!s.perGroup) out[t] = { tab: s.tab, cols: [['ID', 'id']].concat(s.cols.filter(c => !c.virtual).map(c => [c.h, c.f])).concat([['Actualizado', 'u']]) };
  });
  return out;
}
function perGroup_(t) { return !!(SCHEMA[t] && SCHEMA[t].perGroup); }

/**
 * Datos privados: se leen de las Propiedades del script (ver el encabezado), nunca de este archivo.
 * API_KEY: clave que la app HTML debe enviar para leer y escribir la planilla.
 */
function prop_(k) { return String(PropertiesService.getScriptProperties().getProperty(k) || '').trim(); }

/** Acceso desde el archivo HTML: la app envía {action, key, ...} por POST. */
function doPost(e) {
  let res;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const key = prop_('API_KEY');
    if (!key) throw new Error('Falta configurar la clave: en Apps Script, Configuración del proyecto > Propiedades del script, agregá API_KEY.');
    if (body.key !== key) throw new Error('Clave incorrecta: la clave de la app no coincide con la propiedad API_KEY del script.');
    let data;
    if (body.action === 'getAll') data = getAll();
    else if (body.action === 'getInfo') data = getInfo();
    else if (body.action === 'getVersion') data = getVersion();
    else if (body.action === 'putChanges') data = putChanges(body.up || {}, body.dels || [], body.baseVer || '');
    else throw new Error('Acción desconocida: ' + body.action);
    res = { ok: true, data: data };
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  let out;
  try {
    out = HtmlService.createHtmlOutputFromFile('index');
  } catch (e) {
    // Sin archivo "index": se usa la copia de la app incluida en App.gs.
    const html = Utilities.newBlob(Utilities.base64Decode(APP_HTML_B64.join(''))).getDataAsString('UTF-8');
    out = HtmlService.createHtmlOutput(html);
  }
  return out
    .setTitle('CompareIt')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Datos de la planilla conectada, para mostrar el enlace en la app. */
/* ---------- marca de versión: evita descargas innecesarias ----------
 * Cada cambio (desde la app o a mano en la planilla) genera una marca nueva.
 * La app consulta solo la marca (liviano) y descarga los datos únicamente si cambió.
 */
let dirty_ = false;
function bumpVersion_(from) {
  const v = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  PropertiesService.getScriptProperties().setProperties({ ver: v, verAt: new Date().toISOString(), verFrom: from });
  return v;
}
/**
 * Migración única: las escalas pasan a guardarse de mejor a peor (antes, de peor a mejor).
 * Se ejecuta sola la primera vez que la app se conecta con este script.
 */
function migrarEscalas_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('escalaDesc') === '1') return;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (props.getProperty('escalaDesc') === '1') return;
    const sh = ss_().getSheetByName('_Criterios');   // formato anterior a los Criterios Raíz
    if (sh && sh.getLastRow() > 1) {
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      const cs = head.indexOf('Escala');
      if (cs >= 0) {
        const rng = sh.getRange(2, cs + 1, sh.getLastRow() - 1, 1), vals = rng.getValues(); let ch = false;
        vals.forEach(r => { if (String(r[0]).trim()) { r[0] = String(r[0]).split(',').map(x => x.trim()).filter(Boolean).reverse().join(', '); ch = true; } });
        if (ch) { rng.setValues(vals); bumpVersion_('Planilla'); }
      }
    }
    props.setProperty('escalaDesc', '1');
  } finally { lock.releaseLock(); }
}

/**
 * Migración única a Criterios Raíz: la pestaña _Criterios (un criterio por grupo) se divide en
 * _Criterios (catálogo de criterios raíz) y _Usos (qué criterio usa cada grupo, con su padre, peso y orden).
 * La pestaña anterior queda como "_Criterios (anterior)" de respaldo. Las columnas de ítems no cambian.
 */
function migrarRaices_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('raices') === '1') return;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (props.getProperty('raices') === '1') return;
    const old = ss_().getSheetByName('_Criterios');
    if (old && old.getLastColumn() > 0) {
      const vals = old.getDataRange().getValues(), head = (vals.shift() || []).map(String);
      if (head.indexOf('Grupo') >= 0 && head.indexOf('Padre') >= 0) {
        const get = (row, h) => head.indexOf(h) >= 0 ? str_(row[head.indexOf(h)]).trim() : '';
        const rows = vals.filter(r => !r.every(c => c === '')).map(r => ({ id: get(r, 'ID') || newId_('crit'), g: get(r, 'Grupo'), p: get(r, 'Padre'),
          name: get(r, 'Nombre'), kind: get(r, 'Tipo'), dir: get(r, 'Sentido'), unit: get(r, 'Unidad'), scale: get(r, 'Escala'),
          w: get(r, 'Peso'), ord: get(r, 'Orden'), u: get(r, 'Actualizado') }));
        old.setName('_Criterios (anterior)');
        const rdef = sys_().root.cols, udef = sys_().crit.cols, rootSh = sysSheet_('root'), useSh = sysSheet_('crit');
        const rootRows = [], useRows = [];
        rows.forEach(r => {
          const rid = 'r' + (r.id.indexOf('_') >= 0 ? r.id.slice(r.id.indexOf('_')) : '_' + r.id);
          const kind = rows.some(x => x.p === r.id && x.g === r.g) ? 'agg' : (r.kind || 'num');
          const root = { id: rid, name: r.name, kind: kind, scale: kind === 'qual' ? r.scale : '', dir: kind === 'num' ? r.dir : '', unit: kind === 'num' ? r.unit : '', u: r.u };
          const use = { id: r.id, g: r.g, root: rid, p: r.p, w: r.w, ord: r.ord, u: r.u };
          rootRows.push(rdef.map(c => root[c[1]] || '')); useRows.push(udef.map(c => use[c[1]] || ''));
        });
        if (rootRows.length) rootSh.getRange(2, 1, rootRows.length, rdef.length).setValues(rootRows);
        if (useRows.length) useSh.getRange(2, 1, useRows.length, udef.length).setValues(useRows);
        bumpVersion_('Planilla');
      }
    }
    props.setProperty('raices', '1');
  } finally { lock.releaseLock(); }
}
function migrar_() { migrarEscalas_(); migrarRaices_(); }

function getVersion() {
  migrar_();
  const p = PropertiesService.getScriptProperties().getProperties();
  if (!p.ver) { bumpVersion_('Planilla'); return getVersion(); }
  return { v: p.ver, at: p.verAt || '', from: p.verFrom || '' };
}

function getInfo() {
  const ss = ss_(), props = PropertiesService.getScriptProperties();
  if (props.getProperty('trigOk') !== '1') {   // los activadores se revisan una sola vez, no en cada carga
    try { instalarActivador(); props.setProperty('trigOk', '1'); } catch (e) { console.warn('No se pudo instalar el activador: ' + e); }
  }
  return { url: ss.getUrl(), name: ss.getName() };
}

/* ---------- helpers ---------- */
/**
 * ID de la planilla (propiedad del script SHEET_ID): la parte de su URL entre /d/ y /edit.
 * Obligatorio si el proyecto se creó en script.google.com (no vinculado a la planilla).
 */

let SS_ = null;
function ss_() {
  if (SS_) return SS_;   // se abre una sola vez por ejecución
  const sid = prop_('SHEET_ID');
  if (sid) return (SS_ = SpreadsheetApp.openById(sid));
  const ss = SS_ = SpreadsheetApp.getActive();
  if (!ss) throw new Error('El script no está vinculado a una planilla. En Apps Script, Configuración del proyecto > Propiedades del script, agregá SHEET_ID con el ID de tu planilla (la parte de la URL entre /d/ y /edit).');
  return ss;
}
function newId_(t) { return ((SCHEMA[t] && SCHEMA[t].prefix) || 'x') + '_' + Utilities.getUuid().slice(0, 6); }
function str_(v) { return v instanceof Date ? v.toISOString() : String(v == null ? '' : v); }

function sysSheet_(t) {
  const def = sys_()[t];
  let sh = ss_().getSheetByName(def.tab);
  if (!sh) {
    sh = ss_().insertSheet(def.tab);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols.map(c => c[0])]);
    sh.setFrozenRows(1);
  } else {
    // columnas nuevas del esquema: se agregan al final sin tocar las existentes
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    const miss = def.cols.map(c => c[0]).filter(h => head.indexOf(h) < 0);
    if (miss.length) sh.getRange(1, sh.getLastColumn() + 1, 1, miss.length).setValues([miss]);
  }
  return sh;
}

/** Pestaña de un grupo, encontrada por metadato (sobrevive a renombrar la pestaña). */
/** Pestañas de grupo indexadas por su identificador: se buscan una sola vez por ejecución (antes se recorrían todas en cada consulta). */
let GSMAP_ = null;
function gsMap_() {
  if (GSMAP_) return GSMAP_;
  GSMAP_ = {};
  ss_().createDeveloperMetadataFinder().withKey('groupId').find().forEach(md => {
    try { const sh = md.getLocation().getSheet(); if (sh && !GSMAP_[md.getValue()]) GSMAP_[md.getValue()] = sh; } catch (e) {}
  });
  return GSMAP_;
}
/** Google Sheets no admite nombres de pestaña de más de 100 caracteres. */
function tabName_(n) { return String(n || '').slice(0, 100).trim(); }
function groupSheet_(gid, name, create) {
  let sh = gsMap_()[gid]; if (name) name = tabName_(name);
  if (!sh && name) {
    sh = ss_().getSheetByName(name);
    if (sh) { sh.addDeveloperMetadata('groupId', gid); GSMAP_[gid] = sh; }
  }
  if (!sh && create) {
    sh = ss_().insertSheet(name || gid);
    sh.addDeveloperMetadata('groupId', gid); GSMAP_[gid] = sh;
    sh.getRange(1, 1, 1, 4).setValues([['ID', 'Nombre', 'Notas', 'Actualizado']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Columnas de una pestaña de grupo: fijas + "Criterio [c_id]". */
function itemCols_(headers) {
  return headers.map(h => {
    h = String(h);
    if (h === 'ID') return 'id';
    if (h === 'Nombre') return 'name';
    if (h === 'Notas') return 'notes';
    if (h === 'Texto') return 'txt';
    if (h === 'Orden') return 'ord';
    if (h === 'Máx. caracteres') return 'maxc';
    if (h === 'Permitir superar el máximo') return 'maxok';
    if (h === 'Actualizado') return 'u';
    const m = h.match(/\[(c_[\w-]+)\]\s*$/);
    if (m) return 'v.' + m[1];
    const a = h.match(/\[(f_[\w-]+)\]\s*$/);
    return a ? 'x.' + a[1] : null;
  });
}

/* ---------- lectura ---------- */
function getAll() {
  migrar_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const out = {};
    Object.keys(sys_()).forEach(t => {
      const sh = sysSheet_(t), def = sys_()[t];
      const vals = sh.getDataRange().getValues(); const head = vals.shift() || [];
      const keys = head.map(h => (def.cols.find(c => c[0] === h) || [])[1]);
      vals.forEach((row, i) => {
        if (row.every(c => c === '')) return;
        const r = { t: t };
        keys.forEach((k, j) => { if (k && row[j] !== '') r[k] = str_(row[j]); });
        if (!r.id) { r.id = newId_(t); sh.getRange(i + 2, keys.indexOf('id') + 1).setValue(r.id); dirty_ = true; }
        if (r._o) { try { Object.assign(r, JSON.parse(r._o)); } catch (e) {} delete r._o; }
        const id = r.id; delete r.id; out[id] = r;
      });
    });
    Object.keys(out).filter(id => out[id].t === 'group').forEach(gid => {
      const sh = groupSheet_(gid, out[gid].name, false); if (!sh) return;
      normalizarPestana_(sh, gid, out);
      const vals = sh.getDataRange().getValues(); const keys = itemCols_(vals.shift() || []);
      // atributos cuya columna se borró a mano: se archivan (la app guarda sus valores)
      Object.keys(out).filter(id => out[id].t === 'attr' && out[id].g === gid && out[id].arch !== 'si' && keys.indexOf('x.' + id) < 0).forEach(id => {
        out[id].arch = 'si'; out[id].u = new Date().toISOString(); writeSys_(id, out[id]); dirty_ = true;
      });
      vals.forEach((row, i) => {
        if (row.every(c => c === '')) return;
        const r = { t: 'item', g: gid };
        keys.forEach((k, j) => { if (k && row[j] !== '') r[k] = str_(row[j]); });
        if (!r.id) { r.id = newId_('item'); sh.getRange(i + 2, keys.indexOf('id') + 1).setValue(r.id); dirty_ = true; }
        const id = r.id; delete r.id; out[id] = r;
      });
    });
    if (dirty_) { bumpVersion_('Planilla'); dirty_ = false; }   // se completaron IDs o columnas
    return out;
  } finally { lock.releaseLock(); }
}

/**
 * Prepara una pestaña de grupo cargada a mano:
 * - agrega la columna ID si falta (los IDs se completan solos);
 * - si no hay columna "Nombre", usa la primera columna sin reconocer como Nombre;
 * - cada columna nueva escrita a mano pasa a ser un criterio del grupo
 *   (cuantitativo si todos sus valores son números; si no, cualitativo con sus valores como escala).
 */
function normalizarPestana_(sh, gid, out) {
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['ID', 'Nombre', 'Notas', 'Actualizado']]);
    return;
  }
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  if (head.indexOf('ID') < 0) { sh.insertColumnBefore(1); sh.getRange(1, 1).setValue('ID'); head.unshift('ID'); dirty_ = true; }
  let keys = itemCols_(head);
  if (head.indexOf('Nombre') < 0) {
    const j = keys.findIndex((k, i) => !k && head[i]);
    if (j >= 0) { sh.getRange(1, j + 1).setValue('Nombre'); head[j] = 'Nombre'; }
    else { sh.insertColumnAfter(1); sh.getRange(1, 2).setValue('Nombre'); head.splice(1, 0, 'Nombre'); }
    keys = itemCols_(head);
  }
  const lastRow = sh.getLastRow();
  let ord = Math.max(0, ...Object.keys(out).filter(id => out[id].t === 'crit' && out[id].g === gid && !out[id].p).map(id => +out[id].ord || 0));
  let aord = Math.max(0, ...Object.keys(out).filter(id => out[id].t === 'attr' && out[id].g === gid).map(id => +out[id].ord || 0));
  keys.forEach((k, j) => {
    const raw = head[j]; if (!raw) return;
    // columna con marca de atributo cuyo registro no existe (por ejemplo, pestaña copiada): se recrea el atributo
    if (k && k.indexOf('x.') === 0) {
      const aid = k.slice(2);
      if (!out[aid]) { const a = { t: 'attr', g: gid, name: raw.replace(/\s*\[f_[\w-]+\]\s*$/, ''), ord: String(++aord), u: new Date().toISOString() }; writeSys_(aid, a); out[aid] = a; dirty_ = true; }
      return;
    }
    if (k) return;
    const asCrit = /\[criterio\]\s*$/i.test(raw), name = raw.replace(/\s*\[criterio\]\s*$/i, '').trim();
    const vals = lastRow > 1 ? sh.getRange(2, j + 1, lastRow - 1, 1).getValues().map(r => str_(r[0]).trim()).filter(Boolean) : [];
    const numeric = vals.length > 0 && vals.every(x => x !== '' && !isNaN(Number(x.replace(',', '.'))));
    if (!asCrit) {   // atributo: si hay uno archivado con el mismo nombre en el grupo, se reactiva
      const stamp = new Date().toISOString();
      let aid = Object.keys(out).find(id => out[id].t === 'attr' && out[id].g === gid && String(out[id].name || '').toLowerCase() === name.toLowerCase());
      if (aid) { out[aid].arch = ''; out[aid].u = stamp; }
      else { aid = newId_('attr'); out[aid] = { t: 'attr', g: gid, name: name, type: numeric ? 'num' : '', ord: String(++aord), u: stamp }; }
      sh.getRange(1, j + 1).setValue(name + ' [' + aid + ']');
      writeSys_(aid, out[aid]); dirty_ = true;
      return;
    }
    // columna escrita a mano: si ya hay un criterio raíz con ese nombre se reutiliza; si no, se crea
    const cid = newId_('crit'), stamp = new Date().toISOString();
    let rid = Object.keys(out).find(id => out[id].t === 'root' && String(out[id].name || '').toLowerCase() === name.toLowerCase());
    if (!rid) {
      rid = newId_('root');
      const root = { t: 'root', name: name, u: stamp };
      if (numeric) { root.kind = 'num'; root.dir = 'max'; }
      else { root.kind = 'qual'; root.scale = Array.from(new Set(vals)).join(', '); }
      writeSys_(rid, root); out[rid] = root;
    }
    const use = { t: 'crit', g: gid, root: rid, p: '', w: '1', ord: String(++ord), u: stamp };
    sh.getRange(1, j + 1).setValue(name + ' [' + cid + ']');
    writeSys_(cid, use); dirty_ = true;
    out[cid] = use;
  });
}

/**
 * Orden de lectura de las pestañas: grupos activos (en el orden manual de la app), después las de sistema y otras,
 * después los grupos archivados y al final los respaldos («(borrado) …», «… (anterior)»).
 * Dentro de cada bloque se conserva el orden actual. Solo mueve lo que no está en su lugar.
 */
function ordenarPestanas_() {
  const ss = ss_(), map = gsMap_(), bySheet = {};
  Object.keys(map).forEach(gid => { bySheet[map[gid].getSheetId()] = gid; });
  const gsh = sysSheet_('group'), vals = gsh.getDataRange().getValues(), head = (vals.shift() || []).map(String);
  const iId = head.indexOf('ID'), iA = head.indexOf('Archivado'), iO = head.indexOf('Orden'), arch = {}, gord = {};
  vals.forEach(r => { const id = str_(r[iId]); if (id) { arch[id] = iA >= 0 && str_(r[iA]) === 'si'; gord[id] = iO >= 0 && str_(r[iO]) !== '' ? Number(r[iO]) : 1e9; } });
  const rank = sh => {
    const n = sh.getName(), gid = bySheet[sh.getSheetId()];
    if (/^\(borrado\)/.test(n) || /\(anterior\)$/.test(n)) return 3;
    if (gid && arch[gid] !== undefined) return arch[gid] ? 2 : 0;
    return 1;
  };
  // dentro de grupos activos y archivados, respeta el orden manual de los grupos de la app
  const ordOf = sh => { const gid = bySheet[sh.getSheetId()]; return gid && gord[gid] !== undefined ? gord[gid] : 1e9; };
  const target = ss.getSheets().map((sh, i) => ({ sh: sh, r: rank(sh), o: ordOf(sh), i: i }))
    .sort((a, b) => a.r - b.r || ((a.r === 0 || a.r === 2) ? a.o - b.o : 0) || a.i - b.i);
  target.forEach((t, pos) => {
    if (ss.getSheets()[pos].getSheetId() !== t.sh.getSheetId()) { ss.setActiveSheet(t.sh); ss.moveActiveSheet(pos + 1); }
  });
}

/* ---------- escritura ---------- */
function putChanges(up, dels, baseVer) {
  migrar_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  // si la planilla cambió después de la última descarga de la app, se avisa para que vuelva a descargar
  const stale = !!baseVer && baseVer !== getVersion().v;
  try {
    const ids = Object.keys(up || {});
    const rank = t => t === 'group' ? 0 : t === 'root' ? 1 : t === 'crit' ? 2 : perGroup_(t) ? 4 : 3;
    ids.sort((a, b) => rank(up[a].t) - rank(up[b].t));
    ids.forEach(id => {
      const r = up[id];
      if (perGroup_(r.t)) writeItem_(id, r);
      else if (sys_()[r.t]) writeSys_(id, r);
    });
    (dels || []).forEach(d => deleteRow_(d));
    if (ids.some(id => up[id].t === 'group') || (dels || []).some(d => d.t === 'group')) {
      try { ordenarPestanas_(); } catch (e) { console.warn('No se pudieron ordenar las pestañas: ' + e); }
    }
    return { v: bumpVersion_('Web'), stale: stale };
  } finally { lock.releaseLock(); }
}

function findRow_(sh, id) {
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < col.length; i++) if (String(col[i][0]) === id) return i + 1;
  return 0;
}

function writeSys_(id, r) {
  const def = sys_()[r.t], sh = sysSheet_(r.t);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const o = {};
  Object.keys(r).filter(k => k.indexOf('o.') === 0).forEach(k => o[k] = r[k]);
  const rec = Object.assign({}, r, { id: id, _o: Object.keys(o).length ? JSON.stringify(o) : '' });
  const row = head.map(h => { const c = def.cols.find(x => x[0] === h); return c ? (rec[c[1]] || '') : ''; });
  const n = findRow_(sh, id);
  if (n) sh.getRange(n, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  if (r.t === 'group') {
    const gs = groupSheet_(id, r.name, true);
    const tn = tabName_(r.name);
    if (tn && gs.getName() !== tn && !ss_().getSheetByName(tn)) gs.setName(tn);
  }
  if (r.t === 'crit') ensureCritColumn_(id, r);
  if (r.t === 'attr') { if (r.arch === 'si') removeAttrColumn_(r.g, id); else ensureAttrColumn_(id, r); }
  if (r.t === 'root') usesOfRoot_(id).forEach(u => ensureCritColumn_(u.id, u));   // renombra las columnas en todos los grupos
}

/** Usos (en todos los grupos) de un criterio raíz. */
function usesOfRoot_(rid) {
  const sh = sysSheet_('crit'), vals = sh.getDataRange().getValues(), head = (vals.shift() || []).map(String);
  const iId = head.indexOf('ID'), iG = head.indexOf('Grupo'), iR = head.indexOf('Criterio');
  return vals.filter(r => str_(r[iR]) === rid).map(r => ({ id: str_(r[iId]), g: str_(r[iG]), root: rid }));
}

/** Asegura que cada criterio tenga su columna en la pestaña del grupo (o la renombra). */
function ensureCritColumn_(cid, c) {
  const g = getRecord_('group', c.g); if (!g) return;
  const root = c.root ? getRecord_('root', c.root) : null; if (!root || root.kind === 'agg' || root.kind === 'auto') return;   // los calculados no se guardan
  const sh = groupSheet_(c.g, g.name, true);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const label = root.name + (root.unit ? ' (' + root.unit + ')' : '') + ' [' + cid + ']';
  const j = head.findIndex(h => String(h).indexOf('[' + cid + ']') >= 0);
  if (j >= 0) { sh.getRange(1, j + 1).setValue(label); return; }
  const u = head.indexOf('Actualizado');
  if (u >= 0) { sh.insertColumnBefore(u + 1); sh.getRange(1, u + 1).setValue(label); }
  else sh.getRange(1, head.length + 1).setValue(label);
}

/** Columna de un atributo en la pestaña del grupo: «Nombre [f_id]», antes de los criterios. */
function ensureAttrColumn_(aid, a) {
  const g = getRecord_('group', a.g); if (!g) return;
  const sh = groupSheet_(a.g, g.name, true);
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  const label = String(a.name || 'Atributo') + ' [' + aid + ']';
  const j = head.findIndex(h => h.indexOf('[' + aid + ']') >= 0);
  if (j >= 0) { if (head[j] !== label) sh.getRange(1, j + 1).setValue(label); return; }
  let at = head.findIndex(h => /\[c_[\w-]+\]\s*$/.test(h));
  if (at < 0) at = head.indexOf('Orden');
  if (at < 0) at = head.indexOf('Actualizado');
  if (at >= 0) { sh.insertColumnBefore(at + 1); sh.getRange(1, at + 1).setValue(label); }
  else sh.getRange(1, head.length + 1).setValue(label);
}
function removeAttrColumn_(gid, aid) {
  const g = getRecord_('group', gid); if (!g) return;
  const sh = groupSheet_(gid, g.name, false); if (!sh || !sh.getLastColumn()) return;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const j = head.findIndex(h => h.indexOf('[' + aid + ']') >= 0);
  if (j >= 0) sh.deleteColumn(j + 1);
}

function getRecord_(t, id) {
  const sh = sysSheet_(t), n = findRow_(sh, id); if (!n) return null;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = sh.getRange(n, 1, 1, head.length).getValues()[0];
  const r = {}; head.forEach((h, j) => { const c = sys_()[t].cols.find(x => x[0] === h); if (c) r[c[1]] = str_(row[j]); });
  return r;
}

function writeItem_(id, r) {
  const g = getRecord_('group', r.g); if (!g) return;
  const sh = groupSheet_(r.g, g.name, true);
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (r.txt && head.indexOf('Texto') < 0) {   // grupos de textos: columna Texto después de Nombre
    const nm = head.indexOf('Nombre') + 1 || 1;
    sh.insertColumnAfter(nm); sh.getRange(1, nm + 1).setValue('Texto');
    head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  }
  // máximo de caracteres del texto (opcional): columnas después de Texto
  [['maxc', 'Máx. caracteres'], ['maxok', 'Permitir superar el máximo']].forEach(([f, h]) => {
    if (!r[f] || head.indexOf(h) >= 0) return;
    const after = Math.max(head.indexOf('Permitir superar el máximo'), head.indexOf('Máx. caracteres'), head.indexOf('Texto'), head.indexOf('Nombre')) + 1 || 1;
    sh.insertColumnAfter(after); sh.getRange(1, after + 1).setValue(h);
    head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  });
  if (r.ord && head.indexOf('Orden') < 0) {   // orden manual de los ítems: columna antes de Actualizado
    const u = head.indexOf('Actualizado');
    if (u >= 0) { sh.insertColumnBefore(u + 1); sh.getRange(1, u + 1).setValue('Orden'); } else sh.getRange(1, head.length + 1).setValue('Orden');
    head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  }
  Object.keys(r).filter(k => k.indexOf('v.') === 0).forEach(k => {
    const cid = k.slice(2);
    if (!head.some(h => String(h).indexOf('[' + cid + ']') >= 0)) {
      const c = getRecord_('crit', cid); if (c) ensureCritColumn_(cid, c);
    }
  });
  Object.keys(r).filter(k => k.indexOf('x.') === 0).forEach(k => {
    const aid = k.slice(2);
    if (!head.some(h => String(h).indexOf('[' + aid + ']') >= 0)) {
      const a = getRecord_('attr', aid); if (a && a.arch !== 'si') ensureAttrColumn_(aid, a);
    }
  });
  head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const keys = itemCols_(head), rec = Object.assign({}, r, { id: id });
  const row = keys.map(k => k ? (rec[k] || '') : '');
  const n = findRow_(sh, id);
  if (n) {
    const cur = sh.getRange(n, 1, 1, row.length).getValues()[0];
    keys.forEach((k, j) => { if (!k) row[j] = cur[j]; });   // conserva columnas libres del usuario
    sh.getRange(n, 1, 1, row.length).setValues([row]);
  } else sh.appendRow(row);
}

function deleteRow_(d) {
  let sh;
  if (perGroup_(d.t)) { const g = getRecord_('group', d.g); sh = g ? groupSheet_(d.g, g.name, false) : null; }
  else if (sys_()[d.t]) sh = sysSheet_(d.t);
  if (!sh) return;
  const n = findRow_(sh, d.id); if (n) sh.deleteRow(n);
  if (d.t === 'attr' && d.g) removeAttrColumn_(d.g, d.id);
  if (d.t === 'group') {
    const gs = groupSheet_(d.id, null, false);
    if (gs) gs.setName(tabName_('(borrado) ' + gs.getName()));   // se archiva, no se elimina
  }
}

/* ---------- marca de tiempo al editar a mano ---------- */
/**
 * Crea (una sola vez) el activador que completa "Actualizado" al editar la planilla a mano.
 * La app lo llama sola al abrirse; también podés ejecutarlo desde el editor.
 */
function instalarActivador() {
  const ya = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()), id = ss_().getId(), hechos = [];
  if (ya.indexOf('marcarActualizado') < 0) { ScriptApp.newTrigger('marcarActualizado').forSpreadsheet(id).onEdit().create(); hechos.push('edición'); }
  if (ya.indexOf('cambioEstructura') < 0) { ScriptApp.newTrigger('cambioEstructura').forSpreadsheet(id).onChange().create(); hechos.push('cambios de estructura'); }
  return hechos.length ? 'Activadores instalados: ' + hechos.join(' y ') + '.' : 'Los activadores ya estaban instalados.';
}

function marcarActualizado(e) {
  if (!e || !e.range) return;
  bumpVersion_('Planilla');
  const sh = e.range.getSheet(); if (e.range.getRow() < 2) return;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = head.indexOf('Actualizado') + 1; if (!col || e.range.getColumn() === col) return;
  for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) sh.getRange(r, col).setValue(new Date().toISOString());
}

/** Filas o columnas insertadas o borradas, pestañas nuevas o renombradas: el activador de edición no las detecta. */
function cambioEstructura(e) {
  if (e && e.changeType === 'EDIT') return;   // las ediciones ya las marca marcarActualizado
  bumpVersion_('Planilla');
}
