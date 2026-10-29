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

const TZ = 'America/Argentina/Buenos_Aires';
const H_TORNEOS = 'Torneos';
const H_INSC = 'Inscripciones';
const COLS_T = ['torneo', 'categoria', 'cupo', 'apertura', 'cierre'];
const COLS_I = ['fecha_hora', 'torneo', 'categoria', 'jugador_1', 'jugador_2', 'telefono', 'estado', 'baja', 'disponibilidad'];
const FRANJAS = ['Viernes 14 a 00 hs', 'Sábado mañana 9 a 13 hs', 'Sábado siesta/tarde 13 a 18 hs', 'Sábado noche 18 a 00 hs'];

/** Correr UNA vez a mano: crea las hojas con los torneos de noviembre. */
function setup() {
  const ss = SpreadsheetApp.getActive();
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
  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(hoja1);
}

/* ---------- utilidades ---------- */

function leer_(nombre) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nombre);
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
function onEdit(e) {
  const n = e && e.range && e.range.getSheet().getName();
  if (n === H_INSC || n === H_TORNEOS) recalcular_();
}

/** Agrega a la hoja de inscripciones las columnas nuevas que falten. */
function asegurarColumnas_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(H_INSC);
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

  const j1 = limpio_(d.jugador_1), j2 = limpio_(d.jugador_2);
  const tel = String(d.telefono || '').replace(/[^\d+]/g, '').slice(0, 20);
  const disp = (Array.isArray(d.disponibilidad) ? d.disponibilidad : []).filter((f) => FRANJAS.includes(f));
  if (j1.split(' ').length < 2 || j2.split(' ').length < 2 || tel.length < 8 || !disp.length) {
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
    sh.appendRow([ahora, d.torneo, d.categoria, j1, j2, "'" + tel, estado, '', dispTxt]);
    SpreadsheetApp.flush();
    return json_({
      ok: true, estado,
      posicion: estado === 'Adentro' ? pos : pos - cupo,
      hora: Utilities.formatDate(ahora, TZ, 'dd/MM HH:mm:ss'),
    });
  } finally {
    lock.releaseLock();
  }
}
