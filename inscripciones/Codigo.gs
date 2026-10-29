/**
 * Inscripciones a torneos — Complejo Ramirez Padel Club
 *
 * Va pegado en Extensiones → Apps Script de una planilla de Google.
 * La planilla tiene dos hojas:
 *   - "Torneos": una fila por torneo + categoría, con cupo, apertura y cierre.
 *   - "Inscripciones": se llena sola con cada pareja, con fecha y hora al segundo.
 *
 * Inscripciones, bajas y cambios de disponibilidad: hasta el martes anterior al torneo, 18 hs.
 * Para dar de baja una pareja: escribí "x" en la columna "baja" de su fila.
 * El estado (Adentro / Lista de espera) se recalcula solo.
 */

/** ID de la planilla (lo que va entre /d/ y /edit en su link). Vacío = la planilla donde está pegado el script. */
const SHEET_ID = '';
function libro_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive(); }

const TZ = 'America/Argentina/Buenos_Aires';
const H_TORNEOS = 'Torneos';
const H_INSC = 'Inscripciones';
const COLS_T = ['torneo', 'categoria', 'cupo', 'apertura', 'cierre'];
const COLS_I = ['fecha_hora', 'torneo', 'categoria', 'jugador_1', 'jugador_2', 'telefono', 'estado', 'baja', 'disponibilidad', 'email', 'codigo'];
const WEB = 'https://ramirezpadelclub.netlify.app/#torneos';
/** Regla del circuito: al menos dos franjas del sábado, o el viernes completo. */
function dispValida_(disp) { return disp.some((f) => f.indexOf('Viernes') === 0) || disp.filter((f) => f.indexOf('Sábado') === 0).length >= 2; }
const FRANJAS = ['Viernes 14 a 00 hs', 'Sábado mañana 9 a 13 hs', 'Sábado siesta/tarde 13 a 18 hs', 'Sábado noche 18 a 00 hs'];

/** Correr UNA vez a mano: crea las hojas con los torneos de noviembre. */
function setup() {
  const ss = libro_();
  ss.setSpreadsheetTimeZone(TZ);

  let t = ss.getSheetByName(H_TORNEOS) || ss.insertSheet(H_TORNEOS);
  if (t.getLastRow() === 0) {
    t.appendRow(COLS_T);
    const f = (s) => (s ? Utilities.parseDate(s, TZ, 'yyyy-MM-dd HH:mm') : '');
    const filas = [
      ['prueba', 'Prueba Damas', 2, f('2026-10-29 11:30'), ''],
      ['prueba', 'Prueba Caballeros', 2, f('2026-10-29 11:30'), ''],
      ['t-1315nov', 'Damas suma 13', 16, '', f('2026-11-10 18:00')],
      ['t-1315nov', 'Caballeros suma 11', 16, '', f('2026-11-10 18:00')],
      ['master-2022nov', '6° Damas', 12, '', f('2026-11-17 18:00')],
      ['master-2022nov', '5° Damas', 12, '', f('2026-11-17 18:00')],
      ['master-2022nov', '7° Caballeros', 12, '', f('2026-11-17 18:00')],
      ['master-2022nov', 'Suma 8 Caballeros', 12, '', f('2026-11-17 18:00')],
      ['8dam-2022nov', '8° Damas', 16, '', f('2026-11-17 18:00')],
      ['master-2729nov', '7° Damas', 12, '', f('2026-11-24 18:00')],
      ['master-2729nov', 'Suma 9 Damas', 12, '', f('2026-11-24 18:00')],
      ['master-2729nov', '6° Caballeros', 12, '', f('2026-11-24 18:00')],
      ['master-2729nov', '5° Caballeros', 12, '', f('2026-11-24 18:00')],
      ['8cab-2729nov', '8° Caballeros', 16, '', f('2026-11-24 18:00')],
    ];
    t.getRange(2, 1, filas.length, COLS_T.length).setValues(filas);
    t.getRange('D:E').setNumberFormat('dd/mm/yyyy hh:mm');
    t.setFrozenRows(1);
    t.getRange(1, 1, 1, COLS_T.length).setFontWeight('bold').setBackground('#E6F2EB');
    t.autoResizeColumns(1, COLS_T.length);
  }

  let i = ss.getSheetByName(H_INSC) || ss.insertSheet(H_INSC);
  if (i.getLastRow() === 0) {
    i.appendRow(COLS_I);
    i.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm:ss');
    i.setFrozenRows(1);
    i.getRange(1, 1, 1, COLS_I.length).setFontWeight('bold').setBackground('#E6F2EB');
  }
  if (SHEET_ID && !ScriptApp.getProjectTriggers().some((x) => x.getHandlerFunction() === 'alEditar')) {
    ScriptApp.newTrigger('alEditar').forSpreadsheet(ss).onEdit().create();
  }
  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(hoja1);
}

/* ---------- utilidades ---------- */

function leer_(nombre) {
  const sh = libro_().getSheetByName(nombre);
  const v = sh.getDataRange().getValues();
  const cab = v.shift();
  return { sh, filas: v.map((r, n) => { const o = { _fila: n + 2 }; cab.forEach((c, k) => (o[c] = r[k])); return o; }) };
}

/** Normaliza un nombre: minúsculas, sin tildes, sin signos, palabras ordenadas. */
function norm_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zñ ]/g, ' ').split(/\s+/).filter(Boolean).sort().join(' ');
}

function limpio_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim().slice(0, 60)
    .toLowerCase().replace(/(^|[\s'-])(\S)/g, (m, a, b) => a + b.toUpperCase());
}

function fecha_(d) { return d instanceof Date ? d : (d ? new Date(d) : null); }

function estadoVentana_(t, ahora) {
  const ap = fecha_(t.apertura), ci = fecha_(t.cierre);
  if (!ap) return 'pronto';
  if (ahora < ap) return 'antes';
  if (ci && ahora > ci) return 'cerrada';
  return 'abierta';
}

/** Recalcula Adentro / Lista de espera / Baja en la hoja de inscripciones. */
function recalcular_() {
  const T = leer_(H_TORNEOS).filas;
  const { sh, filas } = leer_(H_INSC);
  if (!filas.length) return;
  const cupos = {};
  T.forEach((t) => (cupos[t.torneo + '|' + t.categoria] = Number(t.cupo) || 0));
  const usados = {};
  const est = filas.map((r) => {
    if (String(r.baja).trim()) return ['Baja'];
    const k = r.torneo + '|' + r.categoria;
    usados[k] = (usados[k] || 0) + 1;
    const cupo = cupos[k] || 0;
    return [cupo && usados[k] > cupo ? 'Lista de espera' : 'Adentro'];
  });
  sh.getRange(2, COLS_I.indexOf('estado') + 1, est.length, 1).setValues(est);
}

/** Se ejecuta solo cuando alguien edita la planilla a mano (por ejemplo, una baja). */
function alEditar(e) { onEdit(e); }

function onEdit(e) {
  const n = e && e.range && e.range.getSheet().getName();
  if (n === H_INSC || n === H_TORNEOS) recalcular_();
}

/** Agrega a la hoja de inscripciones las columnas nuevas que falten. */
function asegurarColumnas_() {
  const sh = libro_().getSheetByName(H_INSC);
  const cab = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  COLS_I.forEach((c, k) => { if (cab[k] !== c) sh.getRange(1, k + 1).setValue(c).setFontWeight('bold').setBackground('#E6F2EB'); });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- lo que lee la página ---------- */

function doGet() {
  const ahora = new Date();
  const T = leer_(H_TORNEOS).filas.filter((t) => t.torneo && t.categoria);
  const I = leer_(H_INSC).filas.filter((r) => !String(r.baja).trim());
  const fmt = (d) => (d ? Utilities.formatDate(fecha_(d), TZ, "yyyy-MM-dd'T'HH:mm:ss") : null);
  const out = T.map((t) => {
    const cupo = Number(t.cupo) || 0;
    const lista = I.filter((r) => r.torneo === t.torneo && r.categoria === t.categoria)
      .map((r, n) => ({
        pareja: r.jugador_1 + ' – ' + r.jugador_2,
        hora: Utilities.formatDate(fecha_(r.fecha_hora), TZ, 'dd/MM HH:mm:ss'),
        espera: cupo > 0 && n >= cupo,
      }));
    return {
      torneo: t.torneo, categoria: t.categoria, cupo,
      apertura: fmt(t.apertura), cierre: fmt(t.cierre),
      estado: estadoVentana_(t, ahora), inscriptos: lista,
    };
  });
  return json_({ ok: true, ahora: fmt(ahora), categorias: out });
}

/* ---------- lo que manda la página al inscribirse ---------- */

function doPost(e) {
  try {
    return doPost_(e);
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'interno', detalle: String(err) });
  }
}

function doPost_(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'datos' }); }
  if (d.web) return json_({ ok: true, estado: 'Adentro' }); // campo trampa anti-robots
  if (d.accion === 'baja' || d.accion === 'disponibilidad') return gestionar_(d);
  if (d.accion === 'reenviar') return reenviar_(d);

  const j1 = limpio_(d.jugador_1), j2 = limpio_(d.jugador_2);
  const tel = String(d.telefono || '').replace(/[^\d+]/g, '').slice(0, 20);
  const disp = (Array.isArray(d.disponibilidad) ? d.disponibilidad : []).filter((f) => FRANJAS.includes(f));
  const email = mail_(d.email);
  if (j1.split(' ').length < 2 || j2.split(' ').length < 2 || tel.length < 8 || !dispValida_(disp) || !email) {
    return json_({ ok: false, error: 'datos' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ahora = new Date();
    const t = leer_(H_TORNEOS).filas.find((x) => x.torneo === d.torneo && x.categoria === d.categoria);
    if (!t) return json_({ ok: false, error: 'categoria' });
    const v = estadoVentana_(t, ahora);
    if (v !== 'abierta') return json_({ ok: false, error: v });

    const { sh, filas } = leer_(H_INSC);
    const misma = filas.filter((r) => r.torneo === d.torneo && r.categoria === d.categoria && !String(r.baja).trim());
    const n1 = norm_(j1), n2 = norm_(j2);
    if (n1 === n2) return json_({ ok: false, error: 'iguales' });
    const rep = misma.find((r) => [norm_(r.jugador_1), norm_(r.jugador_2)].some((x) => x === n1 || x === n2));
    if (rep) {
      const quien = [norm_(rep.jugador_1), norm_(rep.jugador_2)].includes(n1) ? j1 : j2;
      return json_({ ok: false, error: 'duplicado', jugador: quien });
    }

    const cupo = Number(t.cupo) || 0;
    const pos = misma.length + 1;
    const estado = cupo && pos > cupo ? 'Lista de espera' : 'Adentro';
    asegurarColumnas_();
    const dispTxt = disp.length === FRANJAS.length ? 'Completa' : disp.join(' / ');
    const codigo = codigo_();
    sh.appendRow([ahora, d.torneo, d.categoria, j1, j2, "'" + tel, estado, '', dispTxt, email, codigo]);
    SpreadsheetApp.flush();
    const hora = Utilities.formatDate(ahora, TZ, 'dd/MM HH:mm:ss');
    const posicion = estado === 'Adentro' ? pos : pos - cupo;
    const enviado = enviar_(email, 'Inscripción confirmada · ' + d.categoria, [
      ['Categoría', d.categoria], ['Pareja', j1 + ' – ' + j2],
      ['Estado', estado === 'Adentro' ? 'Adentro · pareja n.º ' + posicion : 'Lista de espera · lugar ' + posicion],
      ['Registrada', hora + ' hs'], ['Disponibilidad', dispTxt],
    ], codigo, d, t);
    return json_({ ok: true, estado, posicion, hora, email, enviado, codigo });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- baja o cambio de disponibilidad, hecho por la propia pareja ---------- */

function gestionar_(d) {
  const email = mail_(d.email);
  const cod = String(d.codigo || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!email || cod.length !== 6) return json_({ ok: false, error: 'datos' });
  let disp = [];
  if (d.accion === 'disponibilidad') {
    disp = (Array.isArray(d.disponibilidad) ? d.disponibilidad : []).filter((f) => FRANJAS.includes(f));
    if (!dispValida_(disp)) return json_({ ok: false, error: 'franjas' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ahora = new Date();
    const t = leer_(H_TORNEOS).filas.find((x) => x.torneo === d.torneo);
    if (!t) return json_({ ok: false, error: 'categoria' });
    const v = estadoVentana_(t, ahora);
    if (v !== 'abierta') return json_({ ok: false, error: v === 'cerrada' ? 'cerrada_gestion' : v });

    asegurarColumnas_();
    const { sh, filas } = leer_(H_INSC);
    const r = filas.find((x) => x.torneo === d.torneo && !String(x.baja).trim()
      && String(x.email).toLowerCase() === email && String(x.codigo).toUpperCase() === cod);
    if (!r) return json_({ ok: false, error: 'no_encontrado' });

    const cuando = Utilities.formatDate(ahora, TZ, 'dd/MM HH:mm');
    const pareja = r.jugador_1 + ' – ' + r.jugador_2;
    if (d.accion === 'baja') {
      sh.getRange(r._fila, COLS_I.indexOf('baja') + 1).setValue('x (web ' + cuando + ')');
      SpreadsheetApp.flush();
      recalcular_();
      // Sin email: la baja se confirma en pantalla (ahorra cupo diario de Gmail).
      return json_({ ok: true, accion: 'baja', pareja, categoria: r.categoria });
    }
    const dispTxt = disp.length === FRANJAS.length ? 'Completa' : disp.join(' / ');
    sh.getRange(r._fila, COLS_I.indexOf('disponibilidad') + 1).setValue(dispTxt);
    // Sin email: el cambio se confirma en pantalla (ahorra cupo diario de Gmail).
    return json_({ ok: true, accion: 'disponibilidad', pareja, categoria: r.categoria, disponibilidad: dispTxt });
  } finally {
    lock.releaseLock();
  }
}

/** Reenvía los códigos de las inscripciones activas de ese email en ese torneo (solo a ese email). */
function reenviar_(d) {
  const email = mail_(d.email);
  if (!email) return json_({ ok: false, error: 'datos' });
  const filas = leer_(H_INSC).filas.filter((x) => x.torneo === d.torneo && !String(x.baja).trim() && String(x.email).toLowerCase() === email);
  const T = leer_(H_TORNEOS).filas.find((x) => x.torneo === d.torneo);
  filas.slice(0, 5).forEach((r) => enviar_(email, 'Tu código de inscripción · ' + r.categoria, [
    ['Categoría', r.categoria], ['Pareja', r.jugador_1 + ' – ' + r.jugador_2], ['Disponibilidad', r.disponibilidad],
  ], r.codigo, d, T));
  return json_({ ok: true }); // no dice si existe o no, para no revelar inscripciones ajenas
}

/* ---------- email ---------- */

function mail_(s) {
  const m = String(s || '').trim().toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(m) && m.length <= 100 ? m : '';
}

function codigo_() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let c = '';
  for (let i = 0; i < 6; i++) c += abc[Math.floor(Math.random() * abc.length)];
  return c;
}

function txt_(v, n) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n || 200); }

function nombreTorneo_(d) {
  const i = d.torneo_info || {};
  return i.titulo ? txt_(i.titulo, 80) : txt_(d.torneo_nombre || d.torneo, 120);
}

/** Datos del torneo tal como figuran en la tarjeta/flyer de la página. */
function bloqueTorneo_(d, t) {
  const i = d.torneo_info || {};
  const filas = [];
  if (i.fecha) filas.push(['Fecha', txt_(i.fecha, 80).toLowerCase().replace(/^./, (c) => c.toUpperCase())]);
  if (i.lugar) filas.push(['', txt_(i.lugar, 120)]);
  (Array.isArray(i.datos) ? i.datos : []).slice(0, 8).forEach((x) => {
    const s = txt_(x, 240), k = s.indexOf(':');
    filas.push(k > 0 && k < 40 ? [s.slice(0, k), s.slice(k + 1).trim()] : ['', s]);
  });
  if (i.precio) filas.push(['Inscripción', txt_(i.precio, 60)]);
  if (t && t.cierre) {
    const c = fecha_(t.cierre);
    const dia = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'][Number(Utilities.formatDate(c, TZ, 'u')) % 7];
    filas.push(['Cierre de inscripción', dia + ' ' + Utilities.formatDate(c, TZ, "dd/MM 'a las' HH:mm 'hs'")]);
  }
  filas.push(['Disponibilidad horaria', 'Viernes 14 a 00 hs · Sábado mañana 9 a 13 hs · Sábado siesta/tarde 13 a 18 hs · Sábado noche 18 a 00 hs. Domingo: llaves finales.']);
  const flyer = /^https:\/\/ramirezpadelclub\.netlify\.app\/fotos\/torneos\/[\w-]+\.jpg$/.test(String(i.flyer || '')) ? i.flyer : '';
  return { filas, flyer };
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Manda el correo. Si falla (por ejemplo, cupo diario de Gmail), la inscripción igual queda hecha. */
function enviar_(to, asunto, filas, codigo, d, t) {
  try {
    const tabla = (fs) => '<table style="border-collapse:collapse;font-size:15px;width:100%">' + fs.map((f) =>
      '<tr><td style="padding:7px 12px 7px 0;color:#4D6357;vertical-align:top;white-space:nowrap">' + esc_(f[0]) +
      '</td><td style="padding:7px 0;font-weight:700;color:#10231A">' + esc_(f[1]) + '</td></tr>').join('') + '</table>';
    const tor = d ? bloqueTorneo_(d, t) : { filas: [], flyer: '' };
    const cod = codigo ?
      '<div style="margin:22px 0 0;padding:18px;border:2px dashed #037841;border-radius:14px;background:#F6F8F5">' +
      '<p style="margin:0 0 8px;color:#10231A;font-weight:700">Tu código de 6 caracteres</p>' +
      '<p style="font:700 32px/1 monospace;letter-spacing:8px;color:#037841;margin:0 0 10px">' + esc_(codigo) + '</p>' +
      '<p style="margin:0;color:#4D6357;font-size:14px;line-height:1.5">Este código <b>solo lo vas a necesitar si tienen que darse de baja del torneo o cambiar la disponibilidad horaria</b>. ' +
      'Se hace desde la página, en el botón <b>“Baja o cambiar horarios”</b> del torneo, poniendo este email y el código. ' +
      'Se puede hasta el martes anterior al torneo a las 18 hs. No lo compartas: con este código cualquiera podría dar de baja a la pareja.</p></div>' : '';
    const html = '<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;border:1px solid #D5E3DA;border-radius:16px;overflow:hidden">' +
      '<div style="background:#024A29;color:#fff;padding:18px 22px;font:700 18px Arial">Complejo Ramirez Padel Club</div>' +
      '<div style="padding:22px"><h2 style="margin:0 0 6px;color:#024A29;font-size:22px">' + esc_(asunto.split(' · ')[0]) + '</h2>' +
      (d ? '<p style="margin:0 0 16px;color:#4D6357;font-size:15px">' + esc_(nombreTorneo_(d)) + '</p>' : '') +
      '<h3 style="margin:18px 0 6px;font-size:16px;color:#037841">Tu inscripción</h3>' + tabla(filas) + cod +
      (tor.filas.length ? '<h3 style="margin:24px 0 6px;font-size:16px;color:#037841">El torneo</h3>' + tabla(tor.filas) : '') +
      (tor.flyer ? '<p style="margin:18px 0 0;text-align:center"><img src="' + tor.flyer + '" alt="Flyer del torneo" style="width:100%;max-width:360px;border-radius:12px"></p>' : '') +
      '<p style="margin:22px 0 0"><a href="' + WEB + '" style="background:#037841;color:#fff;padding:11px 18px;border-radius:999px;text-decoration:none;font-weight:700">Ver torneos en la página</a></p>' +
      '<p style="margin:18px 0 0;color:#4D6357;font-size:13px">Cualquier consulta, escribinos por WhatsApp al complejo.</p>' +
      '</div></div>';
    MailApp.sendEmail({ to: to, subject: asunto + ' · ' + nombreTorneo_(d || {}), htmlBody: html, name: 'Complejo Ramirez Padel Club' });
    return true;
  } catch (err) {
    console.error(err);
    return false;
  }
}

/** Correr una vez a mano después de actualizar el código: pide el permiso para mandar emails. */
function autorizarEmail() {
  Logger.log('Emails disponibles hoy: ' + MailApp.getRemainingDailyQuota());
}
