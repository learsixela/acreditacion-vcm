/**
 * Acreditación Seminario VCM — backend en Google Apps Script.
 * La Google Sheet que contiene este script es la base de datos.
 *
 * 1) Ejecute una vez la función `setup` (crea las hojas "Personas" y "Resumen").
 * 2) Implemente como Aplicación web: Ejecutar como "Yo", acceso "Cualquier usuario".
 * 3) Copie la URL que termina en /exec en config.js de la página.
 */

// Clave opcional compartida con la página (config.js). Déjela '' para no usarla.
// No es una contraseña real (queda visible en la página), solo evita usos casuales.
const CLAVE = '';

const HOJA = 'Personas';
const TZ = 'America/Santiago';
// Orden fijo de columnas (no las reordene; puede agregar columnas al final).
const COLS = [
  ['correo', 'Correo institucional'],
  ['nombre', 'Nombre completo'],
  ['institucion', 'Universidad / Institución / Organización'],
  ['unidad', 'Unidad a la que pertenece'],
  ['cargo', 'Cargo'],
  ['rol', 'Rol'],
  ['rolOtro', 'Rol (otro)'],
  ['preinscrito', 'Inscrito previamente'],
  ['asistio', 'Asistió'],
  ['horaAsistencia', 'Hora asistencia'],
  ['creado', 'Creado'],
  ['actualizado', 'Actualizado'],
];
const IDX = Object.fromEntries(COLS.map((c, i) => [c[0], i]));
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/* ---------------- configuración inicial ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  const sh = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA, 0);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS.map(c => c[1])])
    .setFontWeight('bold').setBackground('#1f2f78').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, IDX.horaAsistencia + 1, sh.getMaxRows() - 1, 3).setNumberFormat('dd-MM-yyyy HH:mm');
  sh.getRange(2, IDX.correo + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  sh.autoResizeColumns(1, COLS.length);

  const r = ss.getSheetByName('Resumen') || ss.insertSheet('Resumen');
  r.clear();
  const P = `'${HOJA}'!`;
  const c = k => P + String.fromCharCode(65 + IDX[k]) + '2:' + String.fromCharCode(65 + IDX[k]);
  // setFormulas usa siempre la sintaxis en inglés (comas), sin importar la configuración regional.
  r.getRange('A1:A5').setValues([['Indicador'], ['Inscritos previos'], ['Inscritos que asistieron'], ['Registrados en sala'], ['Total asistentes']]);
  r.getRange('B1').setValue('Cantidad');
  r.getRange('B2:B5').setFormulas([
    [`=COUNTIFS(${c('correo')},"<>",${c('preinscrito')},"<>No")`],
    [`=COUNTIFS(${c('correo')},"<>",${c('preinscrito')},"<>No",${c('asistio')},"Sí")`],
    [`=COUNTIFS(${c('correo')},"<>",${c('preinscrito')},"No")`],
    [`=COUNTIFS(${c('correo')},"<>",${c('asistio')},"Sí")`],
  ]);
  r.getRange('A1:B1').setFontWeight('bold');
  r.autoResizeColumns(1, 2);
}

/* ---------------- HTTP ---------------- */
function doGet(e) {
  return handle_(e && e.parameter || {});
}
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'json_invalido' }); }
  return handle_(body);
}
function handle_(p) {
  try {
    if (CLAVE && p.clave !== CLAVE) return out_({ ok: false, error: 'clave' });
    switch (p.action) {
      case 'ping': return out_({ ok: true });
      case 'buscar': return out_(buscar_(p));
      case 'confirmar': return out_(withLock_(() => confirmar_(p)));
      case 'registrar': return out_(withLock_(() => registrar_(p)));
      default: return out_({ ok: false, error: 'accion_desconocida' });
    }
  } catch (err) {
    return out_({ ok: false, error: 'servidor', detalle: String(err && err.message || err) });
  }
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'ocupado' };
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- datos ---------------- */
function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA);
  if (!sh) throw new Error('Falta la hoja "' + HOJA + '". Ejecute setup().');
  return sh;
}
const norm_ = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
const clean_ = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, 200);
const si_ = v => norm_(v) === 'si' || v === true;

function readAll_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, COLS.length).getValues();
  const res = [];
  vals.forEach((r, i) => {
    const correo = norm_(r[IDX.correo]);
    if (!correo) return;
    res.push({ row: i + 2, correo, r });
  });
  return res;
}
function pub_(x) {
  const r = x.r, h = r[IDX.horaAsistencia];
  return {
    correo: x.correo,
    nombre: clean_(r[IDX.nombre]),
    institucion: clean_(r[IDX.institucion]),
    unidad: clean_(r[IDX.unidad]),
    cargo: clean_(r[IDX.cargo]),
    rol: clean_(r[IDX.rol]),
    rolOtro: clean_(r[IDX.rolOtro]),
    preinscrito: norm_(r[IDX.preinscrito]) !== 'no',
    asistio: si_(r[IDX.asistio]),
    hora: h instanceof Date ? Utilities.formatDate(h, TZ, 'HH:mm') : '',
  };
}
function buscar_(p) {
  const q = norm_(p.q), campo = p.campo || 'nombre';
  if (q.length < 2) return { ok: true, items: [] };
  const all = readAll_();
  let items;
  if (campo === 'correo') {
    const exact = all.filter(x => x.correo === q);
    items = exact.length ? exact : all.filter(x => x.correo.indexOf(q) >= 0);
  } else if (campo === 'institucion') {
    items = all.filter(x => norm_(x.r[IDX.institucion]).indexOf(q) >= 0);
  } else {
    items = all.filter(x => norm_(x.r[IDX.nombre]).indexOf(q) >= 0 || x.correo.indexOf(q) >= 0);
  }
  items = items.map(pub_).sort((a, b) => norm_(a.nombre).localeCompare(norm_(b.nombre))).slice(0, 25);
  return { ok: true, items };
}
function validar_(p) {
  const o = {
    correo: norm_(p.correo), nombre: clean_(p.nombre), institucion: clean_(p.institucion),
    unidad: clean_(p.unidad), cargo: clean_(p.cargo), rol: clean_(p.rol), rolOtro: clean_(p.rolOtro),
  };
  if (['Expositor', 'Asistente', 'Otro'].indexOf(o.rol) < 0) o.rol = '';
  if (o.rol !== 'Otro') o.rolOtro = '';
  const falta = [];
  if (!o.nombre) falta.push('nombre');
  if (!EMAIL_RE.test(o.correo)) falta.push('correo');
  if (!o.institucion) falta.push('institucion');
  if (!o.unidad) falta.push('unidad');
  if (!o.cargo) falta.push('cargo');
  if (!o.rol) falta.push('rol');
  if (o.rol === 'Otro' && !o.rolOtro) falta.push('rolOtro');
  return { o, falta };
}
function confirmar_(p) {
  const { o, falta } = validar_(p);
  if (falta.length) return { ok: false, error: 'faltan', falta };
  const x = readAll_().find(y => y.correo === o.correo);
  if (!x) return { ok: false, error: 'no_inscrito' };
  const sh = sheet_(), r = x.r.slice(), now = new Date();
  const ya = si_(r[IDX.asistio]);
  ['nombre', 'institucion', 'unidad', 'cargo', 'rol', 'rolOtro'].forEach(k => { r[IDX[k]] = o[k]; });
  r[IDX.asistio] = 'Sí';
  if (!ya || !(r[IDX.horaAsistencia] instanceof Date)) r[IDX.horaAsistencia] = now;
  if (!clean_(r[IDX.preinscrito])) r[IDX.preinscrito] = 'Sí';
  r[IDX.actualizado] = now;
  sh.getRange(x.row, 1, 1, COLS.length).setValues([r]);
  return { ok: true, yaConfirmado: ya, persona: pub_({ correo: x.correo, r }) };
}
function registrar_(p) {
  const { o, falta } = validar_(p);
  if (falta.length) return { ok: false, error: 'faltan', falta };
  const x = readAll_().find(y => y.correo === o.correo);
  if (x) return { ok: false, error: 'existe', persona: pub_(x) };
  const now = new Date(), r = new Array(COLS.length).fill('');
  Object.keys(o).forEach(k => { r[IDX[k]] = o[k]; });
  r[IDX.preinscrito] = 'No';
  r[IDX.asistio] = 'Sí';
  r[IDX.horaAsistencia] = now;
  r[IDX.creado] = now;
  r[IDX.actualizado] = now;
  sheet_().appendRow(r);
  return { ok: true, persona: pub_({ correo: o.correo, r }) };
}

/* ---------------- importar inscritos desde las respuestas del Google Form ---------------- */
// Copia a "Personas" las respuestas de la hoja del formulario, reconociendo las columnas por su
// encabezado. No modifica la hoja del formulario. Se puede ejecutar varias veces: no duplica correos.
function importarFormulario() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(HOJA)) setup();
  const src = ss.getSheets().find(s => s.getName() !== HOJA && s.getName() !== 'Resumen' &&
    s.getLastRow() > 1 && s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].some(h => /correo|e-?mail/i.test(h)));
  if (!src) throw new Error('No encontré una hoja con una columna de correo.');
  const vals = src.getDataRange().getValues();
  const head = vals[0].map(h => {
    const n = norm_(h);
    if (n.indexOf('correo') >= 0 || n.indexOf('mail') >= 0) return 'correo';
    if (n.indexOf('nombre') >= 0) return 'nombre';
    if (n.indexOf('unidad') >= 0) return 'unidad';
    if (n.indexOf('cargo') >= 0) return 'cargo';
    if (n.indexOf('rol') === 0) return 'rol';
    if (n.indexOf('universidad') >= 0 || n.indexOf('institu') >= 0 || n.indexOf('organiza') >= 0) return 'institucion';
    return null;
  });
  // si hay más de una columna de correo (p. ej. "Dirección de correo electrónico" del Form), usa la última con datos
  const existentes = new Set(readAll_().map(x => x.correo));
  const nuevas = [], now = new Date();
  vals.slice(1).forEach(r => {
    const o = {};
    head.forEach((k, i) => { if (k && clean_(r[i])) o[k] = clean_(r[i]); });
    o.correo = norm_(o.correo);
    if (!EMAIL_RE.test(o.correo) || existentes.has(o.correo)) return;
    const rn = norm_(o.rol);
    if (rn.indexOf('expos') === 0) o.rol = 'Expositor';
    else if (rn.indexOf('asist') === 0) o.rol = 'Asistente';
    else if (o.rol) { o.rolOtro = o.rol.replace(/^otro\s*:?\s*/i, ''); o.rol = 'Otro'; }
    const fila = new Array(COLS.length).fill('');
    Object.keys(o).forEach(k => { if (k in IDX) fila[IDX[k]] = o[k]; });
    fila[IDX.preinscrito] = 'Sí';
    fila[IDX.creado] = now;
    nuevas.push(fila);
    existentes.add(o.correo);
  });
  const sh = sheet_();
  if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, COLS.length).setValues(nuevas);
  Logger.log('Hoja origen: "' + src.getName() + '". Personas nuevas importadas: ' + nuevas.length);
  return nuevas.length;
}
