/**
 * SEGUIMIENTO OBRA AUTOVÍAS — Bacheo con fresado
 * INGECO S.A. — Apps Script + Sheets
 *
 * Fuente de combustible: planilla "ENTREGAS DE COMBUSTIBLE - LEANDRO CASARES"
 * Frontend: GitHub Pages (docs/) → API JSON por doPost. Acceso: PIN de CONFIG o sesión de la app INGECO.
 */

// ---------- IDs ----------
const ID_COMBUSTIBLE = '19eyY8MImPM_-Gyzj8QcqA_yR5KDrkuu-1faJUU48N1A'; // Leandro Casares (Shared Drive)
const HOJA_ENTREGAS = 'ENTREGAS';

const TABS = { CONFIG: 'CONFIG', PROD: 'PRODUCCION', PUNTAS: 'PUNTAS', EQUIPOS: 'EQUIPOS' };

// Acuerdo con los proveedores: los alquileres (salvo la fresadora) se cuentan desde el martes 13-oct-2026;
// los días previos fueron sin costo. Se aplica solo una vez, al agregar la columna DESDE a planillas viejas.
const EQUIPOS_DESDE_DEFAULT = '2026-10-13';
const EQUIPOS_DESDE_EXCEPTO = ['fresa'];

// ---------- Web app ----------
// El frontend vive en GitHub Pages; este /exec solo expone la API JSON (doPost).
const URL_APP = 'https://marcoskatz-cmd.github.io/obra-autovias/';
const APP_INGECO_BACKEND = 'https://script.google.com/macros/s/AKfycbx9NN1FfOQjUjfymLXrMyd2QM6l96HP0lxY6RC7_xKsbDlO0PnwdHbaMtag54VYFor0/exec';
const MODULO_INGECO = 'AUTOVIAS';

function doGet() {
  return HtmlService.createHtmlOutput(
    '<meta http-equiv="refresh" content="0;url=' + URL_APP + '">' +
    '<div style="font-family:sans-serif;padding:40px;text-align:center">Abriendo la app… <a href="' + URL_APP + '">' + URL_APP + '</a></div>'
  ).setTitle('Seguimiento Autovías — INGECO');
}

// Body (text/plain para evitar preflight): { accion, pin | sesion, nombre, payload }
function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'Cuerpo inválido' }); }
  try {
    const user = autorizar_(body);
    const p = body.payload || {};
    let data;
    switch (body.accion) {
      case 'login': data = { ok: true }; break;
      case 'datos': data = getDatos(user); break;
      case 'guardarDia': data = guardarDia(p, user); break;
      case 'guardarPuntas': data = guardarPuntas(p, user); break;
      case 'borrarDia': data = borrarDia(p.fecha, user); break;
      case 'obras': data = obrasCombustible_(); break;
      case 'config': data = leerConfigCompleta_(); break;
      case 'guardarConfig': data = guardarConfig(p, user); break;
      default: throw new Error('Acción desconocida: ' + body.accion);
    }
    return json_({ ok: true, data, usuario: user });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Devuelve el nombre de quien opera. Entra con sesión de la app INGECO (sin PIN) o con el PIN de CONFIG.
function autorizar_(body) {
  const nombre = String(body.nombre || '').trim().slice(0, 60);
  if (body.sesion && sesionIngecoOk_(body.sesion)) return nombre || 'App INGECO';
  const cfg = leerConfig_();
  const pin = String(cfg.PIN === undefined || cfg.PIN === '' ? '1234' : cfg.PIN).trim();
  if (String(body.pin || '').trim() !== pin) throw new Error('PIN incorrecto');
  return nombre || 'PIN';
}

// Sesión de módulo emitida por app-ingeco al canjear el ticket (guard.js). Se cachea 10 min.
function sesionIngecoOk_(sesion) {
  if (!sesion || String(sesion).length < 20) return false;
  const cache = CacheService.getScriptCache();
  const k = 'ingeco_ses_' + String(sesion).slice(0, 40);
  const c = cache.get(k);
  if (c === '1') return true;
  if (c === '0') return false;
  let ok = false;
  try {
    const r = UrlFetchApp.fetch(APP_INGECO_BACKEND, {
      method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true,
      payload: JSON.stringify({ accion: 'validar_sesion_modulo', payload: { sesion: String(sesion), modulo: MODULO_INGECO } })
    });
    const j = JSON.parse(r.getContentText());
    ok = !!(j && j.ok);
  } catch (err) { ok = false; }
  cache.put(k, ok ? '1' : '0', ok ? 600 : 60);
  return ok;
}

// Correr UNA vez desde el editor para autorizar UrlFetchApp (acceso desde la app INGECO).
function autorizarAppIngeco() {
  const r = UrlFetchApp.fetch(APP_INGECO_BACKEND + '?accion=ping', { muteHttpExceptions: true });
  Logger.log('App INGECO respondió: ' + r.getContentText().slice(0, 120));
}

// ---------- Setup inicial (correr una vez desde el editor) ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let cfg = ss.getSheetByName(TABS.CONFIG) || ss.insertSheet(TABS.CONFIG);
  cfg.clear();
  cfg.getRange(1, 1, 1, 3).setValues([['CLAVE', 'VALOR', 'DESCRIPCIÓN']]).setFontWeight('bold');
  const rows = [
    ['PIN', '1234', 'PIN de acceso a la app (desde la app INGECO no hace falta)'],
    ['ACCESO', 'marcoskatz@grupoingeco.com.ar, agustin@grupoingeco.com.ar', 'Referencia: quiénes usan la app'],
    ['OBRA_COMBUSTIBLE', 'autovia', 'Texto con que Leandro identifica esta obra (OBRA GENERAL o LUGAR ENTREGA). Contiene, sin mayúsculas ni tildes. Él carga "Autovía" / "Autovía del NOA"'],
    ['TC', 1540, 'Tipo de cambio $/USD'],
    ['IVA', 0.21, 'IVA para netear'],
    ['COMB_CON_IVA', 1, '1 si el precio de la planilla de combustible incluye IVA'],
    ['COMB_DESDE', '2026-10-15', 'Desde qué fecha se toma el combustible real de la planilla de Leandro (antes: manual o estimado)'],
    ['COMB_EXCLUIR', 'camion batea', 'Equipos a excluir del combustible (separados por coma; contiene, sin tildes). El camión batea va dentro del flete'],
    ['DIAS_MES', 22, 'Días laborables por mes (divide los alquileres en el mes en curso)'],
    ['OBRA_INICIO', '2026-10-07', 'Primer día de obra: desde acá corren los fijos todos los días hábiles'],
    ['PRECIO_T', 79.5, 'Precio pactado USD/t colocada sin IVA'],
    ['SOBREFRESADO', 1.08, 't fresadas por t colocada'],
    ['CAP_CAMION', 30, 't por viaje de mezcla'],
    ['FLETE_ARS', 625000, '$ por viaje entero sin IVA'],
    ['COMB_ESTIMADO_ARS', 700000, '$/día estimado — se usa solo si no hay carga real ese día'],
    ['MO_USD', 705, 'USD/día mano de obra (12 personas)'],
    ['SENAL_USD', 333, 'USD/día señalización'],
    ['PUNTAS_JUEGO_ARS', 6134000, '$ juego de 200 puntas CON IVA'],
    ['PUNTAS_VIDA_T', 1400, 't fresadas por juego = 200 puntas × 7 t/punta (estimado; se recalcula con recambios reales)'],
  ];
  cfg.getRange(2, 1, rows.length, 3).setValues(rows);
  cfg.setColumnWidths(1, 1, 180); cfg.setColumnWidth(2, 320); cfg.setColumnWidth(3, 420);

  let eq = ss.getSheetByName(TABS.EQUIPOS) || ss.insertSheet(TABS.EQUIPOS);
  eq.clear();
  eq.getRange(1, 1, 1, 6).setValues([['EQUIPO', 'PRECIO MES', 'MONEDA', 'CON IVA (1/0)', 'OBS', 'DESDE']]).setFontWeight('bold');
  eq.getRange(1, 6, eq.getMaxRows(), 1).setNumberFormat('@');   // DESDE como texto AAAA-MM-DD
  // DESDE: primer día que se imputa el alquiler (vacío = desde OBRA_INICIO). Todo menos la fresa arrancó a contar el martes 13-oct-2026.
  eq.getRange(2, 1, 11, 6).setValues([
    ['Fresadora Bomag BM 1000/20', 24145150, 'ARS', 1, 'sin repuestos ni consumibles', ''],
    ['Rodillo mixto 6 t', 7500000, 'ARS', 0, 'con maquinista', EQUIPOS_DESDE_DEFAULT],
    ['Minicargador', 6700000, 'ARS', 0, '', EQUIPOS_DESDE_DEFAULT],
    ['Retropala', 4000, 'USD', 0, '', EQUIPOS_DESDE_DEFAULT],
    ['Terminadora', 3000, 'USD', 0, 'con operario', EQUIPOS_DESDE_DEFAULT],
    ['Camión chasis apoyo', 6000000, 'ARS', 0, '', EQUIPOS_DESDE_DEFAULT],
    ['Camión volcador', 0, 'ARS', 0, 'propio', EQUIPOS_DESDE_DEFAULT],
    ['Acoplado liga', 1000000, 'ARS', 0, '', EQUIPOS_DESDE_DEFAULT],
    ['Camioneta', 3500000, 'ARS', 0, '', EQUIPOS_DESDE_DEFAULT],
    ['Minibús', 0, 'ARS', 0, 'propio', EQUIPOS_DESDE_DEFAULT],
    ['Camión regador', 0, 'ARS', 0, 'propio', EQUIPOS_DESDE_DEFAULT],
  ]);

  let prod = ss.getSheetByName(TABS.PROD) || ss.insertSheet(TABS.PROD);
  if (prod.getLastRow() === 0) {
    prod.getRange(1, 1, 1, 8).setValues([['FECHA', 'TONELADAS', 'VIAJES', 'OBSERVACIONES', 'CARGADO POR', 'TIMESTAMP', 'JORNADA', 'COMBUSTIBLE_ARS']]).setFontWeight('bold');
    prod.setFrozenRows(1);
  }

  let pun = ss.getSheetByName(TABS.PUNTAS) || ss.insertSheet(TABS.PUNTAS);
  if (pun.getLastRow() === 0) {
    pun.getRange(1, 1, 1, 5).setValues([['FECHA', 'JUEGOS', 'OBSERVACIONES', 'CARGADO POR', 'TIMESTAMP']]).setFontWeight('bold');
    pun.setFrozenRows(1);
  }

  const h1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (h1 && ss.getSheets().length > 1) ss.deleteSheet(h1);
}

// ---------- Lectura de config ----------
function leerConfig_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.CONFIG);
  if (!sh) return {};
  const v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  const o = {};
  v.forEach(r => { if (r[0]) o[r[0]] = r[1]; });
  return o;
}

// Config completa para editar desde la app
function leerConfigCompleta_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sc = ss.getSheetByName(TABS.CONFIG);
  const config = sc.getLastRow() > 1 ? sc.getRange(2, 1, sc.getLastRow() - 1, 3).getValues().filter(r => r[0]).map(r => ({ clave: String(r[0]), valor: r[1] instanceof Date ? Utilities.formatDate(r[1], 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : r[1], desc: String(r[2] || '') })) : [];
  return { config, equipos: leerEquipos_() };
}

// Fecha de celda (Date o texto) → 'yyyy-MM-dd'; '' si está vacía o no se entiende
function fechaStr_(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

// Lista de EQUIPOS con su fecha DESDE (col F). Si la planilla es anterior a esa columna, la crea y
// carga EQUIPOS_DESDE_DEFAULT en todos los equipos salvo los de EQUIPOS_DESDE_EXCEPTO (una sola vez).
function leerEquipos_() {
  const se = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.EQUIPOS);
  if (!se) return [];
  if (String(se.getRange(1, 6).getValue()).trim().toUpperCase() !== 'DESDE') {
    se.getRange(1, 6, se.getMaxRows(), 1).setNumberFormat('@');
    se.getRange(1, 6).setValue('DESDE').setFontWeight('bold');
    const n = se.getLastRow() - 1;
    if (n > 0) {
      const nombres = se.getRange(2, 1, n, 1).getValues();
      se.getRange(2, 6, n, 1).setValues(nombres.map(r => [!r[0] || EQUIPOS_DESDE_EXCEPTO.some(x => sinTildes_(r[0]).includes(x)) ? '' : EQUIPOS_DESDE_DEFAULT]));
    }
  }
  if (se.getLastRow() < 2) return [];
  return se.getRange(2, 1, se.getLastRow() - 1, 6).getValues().filter(r => r[0]).map(r => ({
    equipo: String(r[0]), precio: Number(r[1]) || 0, moneda: r[2] === 'USD' ? 'USD' : 'ARS', conIva: Number(r[3]) === 1 ? 1 : 0, obs: String(r[4] || ''), desde: fechaStr_(r[5]),
  }));
}

const CLAVES_NUM = ['TC', 'IVA', 'COMB_CON_IVA', 'DIAS_MES', 'PRECIO_T', 'SOBREFRESADO', 'CAP_CAMION', 'FLETE_ARS', 'COMB_ESTIMADO_ARS', 'MO_USD', 'SENAL_USD', 'PUNTAS_JUEGO_ARS', 'PUNTAS_VIDA_T'];

function guardarConfig(p, user) {
  if (!user) throw new Error('Sin permiso');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sc = ss.getSheetByName(TABS.CONFIG);
  const n = Math.max(sc.getLastRow() - 1, 0);
  const rows = n ? sc.getRange(2, 1, n, 2).getValues() : [];
  const cambios = p.config || {};
  Object.keys(cambios).forEach(k => {
    let v = cambios[k];
    if (CLAVES_NUM.includes(k)) { v = Number(String(v).replace(',', '.')); if (isNaN(v) || v < 0) throw new Error('Valor inválido en ' + k); }
    else v = String(v).trim().slice(0, 200);
    if (k === 'PIN' && String(v).length < 4) throw new Error('El PIN debe tener al menos 4 caracteres');
    if ((k === 'TC' || k === 'PRECIO_T' || k === 'DIAS_MES' || k === 'CAP_CAMION' || k === 'PUNTAS_VIDA_T') && !(v > 0)) throw new Error(k + ' debe ser mayor a 0');
    const i = rows.findIndex(r => String(r[0]) === k);
    if (i >= 0) sc.getRange(i + 2, 2).setValue(v);
    else sc.appendRow([k, v, '']);
  });
  if (Array.isArray(p.equipos)) {
    const se = ss.getSheetByName(TABS.EQUIPOS);
    const lista = p.equipos.filter(e => e && String(e.equipo || '').trim()).map(e => {
      const precio = Number(String(e.precio).replace(',', '.')) || 0;
      if (precio < 0) throw new Error('Precio inválido en ' + e.equipo);
      const desdeRaw = String(e.desde || '').trim(), desde = fechaStr_(desdeRaw);
      if (desdeRaw && !desde) throw new Error('Fecha "desde" inválida en ' + e.equipo + ' (usar AAAA-MM-DD)');
      return [String(e.equipo).trim().slice(0, 80), precio, e.moneda === 'USD' ? 'USD' : 'ARS', Number(e.conIva) === 1 ? 1 : 0, String(e.obs || '').slice(0, 120), desde];
    });
    leerEquipos_();   // asegura la columna DESDE en planillas viejas
    se.getRange(1, 6, se.getMaxRows(), 1).setNumberFormat('@');
    if (se.getLastRow() > 1) se.getRange(2, 1, se.getLastRow() - 1, 6).clearContent();
    if (lista.length) se.getRange(2, 1, lista.length, 6).setValues(lista);
  }
  return { ok: true };
}

// Costo de equipos a régimen (todos los equipos de la lista). El prorrateo real por día,
// respetando la fecha DESDE de cada equipo, se hace en getDatos.
function costoEquiposDia_(cfg) {
  const lista = leerEquipos_();
  if (!lista.length) return { dia: 0, mes: 0, detalle: [] };
  let mes = 0; const detalle = [];
  lista.forEach(e => {
    let usd = e.moneda === 'USD' ? e.precio : e.precio / Number(cfg.TC);
    if (e.conIva === 1) usd = usd / (1 + Number(cfg.IVA));
    mes += usd; detalle.push({ equipo: e.equipo, usdMes: usd, desde: e.desde || '' });
  });
  return { dia: mes / Number(cfg.DIAS_MES), mes, detalle };
}

// ---------- Combustible desde planilla de Leandro ----------
function combustiblePorDia_(cfg) {
  const out = {}; // 'yyyy-MM-dd' -> { litros, ars, registros }
  let diag = { ok: true, msg: '' };
  try {
    const ss = SpreadsheetApp.openById(ID_COMBUSTIBLE);
    const sh = ss.getSheetByName(HOJA_ENTREGAS);
    if (!sh) throw new Error('No existe la hoja ' + HOJA_ENTREGAS);
    const data = sh.getDataRange().getValues();
    const head = data[0].map(h => String(h).trim().toUpperCase());
    const iF = head.indexOf('FECHA'), iL = head.indexOf('CANTIDAD (L)'), iT = head.indexOf('PRECIO TOTAL'),
          iO = head.indexOf('OBRA GENERAL'), iLug = head.findIndex(h => h.startsWith('LUGAR ENTREGA')), iEq = head.findIndex(h => h.startsWith('EQUIPO'));
    const faltan = [['FECHA', iF], ['CANTIDAD (L)', iL], ['PRECIO TOTAL', iT], ['OBRA GENERAL', iO]].filter(x => x[1] < 0).map(x => x[0]);
    if (faltan.length) throw new Error('Faltan columnas en la planilla de combustible: ' + faltan.join(', '));
    const clave = sinTildes_(cfg.OBRA_COMBUSTIBLE);
    const excluir = String(cfg.COMB_EXCLUIR || '').split(',').map(sinTildes_).filter(Boolean);
    let excluidos = 0;
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const obra = sinTildes_(r[iO]), lugar = sinTildes_(iLug >= 0 ? r[iLug] : '');
      if (!clave || (!obra.includes(clave) && !lugar.includes(clave))) continue;
      if (excluir.length) { const fila = sinTildes_(r.map(v => String(v || '')).join(' | ')); if (excluir.some(x => fila.includes(x))) { excluidos++; continue; } }
      const d = aFecha_(r[iF]); if (!d) continue;
      const k = Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
      const litros = Number(r[iL]) || 0;
      const ars = aNumero_(r[iT]);
      if (!out[k]) out[k] = { litros: 0, ars: 0, registros: [] };
      out[k].litros += litros; out[k].ars += ars;
      out[k].registros.push({ equipo: String(iEq >= 0 ? r[iEq] : ''), litros, ars });
    }
    const total = Object.keys(out).length;
    diag.msg = total ? (total + ' días con carga de combustible para "' + cfg.OBRA_COMBUSTIBLE + '"' + (excluidos ? ' (' + excluidos + ' registros excluidos: ' + cfg.COMB_EXCLUIR + ')' : '')) : ('Sin registros para "' + cfg.OBRA_COMBUSTIBLE + '" en la planilla de Leandro. Verificá CONFIG > OBRA_COMBUSTIBLE.');
    if (!total) diag.ok = false;
  } catch (e) { diag = { ok: false, msg: 'No se pudo leer combustible: ' + e.message }; }
  return { porDia: out, diag };
}

function sinTildes_(v) { return String(v || '').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }

// Diagnóstico: valores distintos de OBRA GENERAL / LUGAR ENTREGA en la planilla de Leandro (últimas filas)
function obrasCombustible_() {
  const sh = SpreadsheetApp.openById(ID_COMBUSTIBLE).getSheetByName(HOJA_ENTREGAS);
  const data = sh.getDataRange().getValues();
  const head = data[0].map(h => String(h).trim().toUpperCase());
  const iO = head.indexOf('OBRA GENERAL'), iLug = head.findIndex(h => h.startsWith('LUGAR ENTREGA')), iF = head.indexOf('FECHA');
  const cnt = {};
  for (let i = Math.max(1, data.length - 400); i < data.length; i++) {
    const r = data[i];
    const k = String(r[iO] || '').trim() + ' | ' + String(iLug >= 0 ? r[iLug] : '').trim();
    const d = aFecha_(r[iF]);
    if (!cnt[k]) cnt[k] = { n: 0, ultima: '' };
    cnt[k].n++; if (d) { const f = Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'); if (f > cnt[k].ultima) cnt[k].ultima = f; }
  }
  return { columnas: head, obras: Object.keys(cnt).map(k => ({ obra_lugar: k, n: cnt[k].n, ultima: cnt[k].ultima })).sort((a, b) => a.ultima < b.ultima ? 1 : -1) };
}

function aFecha_(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim(); const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return null;
}
function aNumero_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  return Number(s) || 0;
}

// ---------- Datos para la app ----------
function getDatos(user) {
  const cfg = leerConfig_();
  const TC = Number(cfg.TC), IVA = Number(cfg.IVA);
  const eq = costoEquiposDia_(cfg);
  const fijoDia = eq.dia + Number(cfg.MO_USD) + Number(cfg.SENAL_USD);
  const fleteUSD = Number(cfg.FLETE_ARS) / TC;
  const combEstUSD = Number(cfg.COMB_ESTIMADO_ARS) / TC;
  const puntasJuegoUSD = Number(cfg.PUNTAS_JUEGO_ARS) / (1 + IVA) / TC;
  const comb = combustiblePorDia_(cfg);
  const combDiv = Number(cfg.COMB_CON_IVA) === 1 ? (1 + IVA) : 1;
  const combDesde = cfg.COMB_DESDE ? (cfg.COMB_DESDE instanceof Date ? Utilities.formatDate(cfg.COMB_DESDE, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : String(cfg.COMB_DESDE).trim()) : '';

  // puntas: vida real si hay recambios, sino estimado
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName(TABS.PUNTAS);
  let recambios = [];
  if (shP && shP.getLastRow() > 1) {
    recambios = shP.getRange(2, 1, shP.getLastRow() - 1, 3).getValues()
      .filter(r => r[0]).map(r => ({ fecha: Utilities.formatDate(new Date(r[0]), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'), juegos: Number(r[1]) || 0, obs: r[2] }));
  }

  const shD = ss.getSheetByName(TABS.PROD);
  let dias = [];
  if (shD && shD.getLastRow() > 1) {
    dias = shD.getRange(2, 1, shD.getLastRow() - 1, 8).getValues().filter(r => r[0])
      .map(r => ({ fecha: Utilities.formatDate(new Date(r[0]), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'), t: Number(r[1]) || 0, viajes: r[2] === '' ? null : Number(r[2]), obs: r[3], por: r[4], jornada: r[6] === '' ? 1 : Number(r[6]), combArs: r[7] === '' || r[7] === null || r[7] === undefined ? null : Number(r[7]) }))
      .sort((a, b) => a.fecha < b.fecha ? -1 : 1);
  }

  // vida real de puntas: t fresadas acumuladas hasta el último recambio / juegos acumulados
  const tFresTotal = dias.reduce((s, d) => s + d.t * Number(cfg.SOBREFRESADO), 0);
  const juegosTotal = recambios.reduce((s, r) => s + r.juegos, 0);
  let vidaPuntas = Number(cfg.PUNTAS_VIDA_T), vidaReal = false;
  if (juegosTotal >= 2) {
    const ultimo = recambios.map(r => r.fecha).sort().pop();
    const tHasta = dias.filter(d => d.fecha <= ultimo).reduce((s, d) => s + d.t * Number(cfg.SOBREFRESADO), 0);
    if (tHasta > 0) { vidaPuntas = tHasta / (juegosTotal - 1); vidaReal = true; }
  }
  const puntasPorT = puntasJuegoUSD / vidaPuntas * Number(cfg.SOBREFRESADO);

  // ---- Calendario de obra: los fijos (alquileres, personal, señalización) corren TODOS los días hábiles
  //      desde el inicio de la obra, se tire mezcla o no. Combustible, flete y puntas solo cuando hay producción.
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  const porFecha = {}; dias.forEach(d => porFecha[d.fecha] = d);
  const inicio = cfg.OBRA_INICIO ? (cfg.OBRA_INICIO instanceof Date ? Utilities.formatDate(cfg.OBRA_INICIO, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : String(cfg.OBRA_INICIO).trim()) : (dias.length ? dias[0].fecha : hoy);
  const fin = dias.length && dias[dias.length - 1].fecha > hoy ? dias[dias.length - 1].fecha : hoy;
  const esFinde = f => { const w = new Date(f + 'T12:00:00').getDay(); return w === 0 || w === 6; };
  const calendario = [];
  for (let d = new Date(inicio + 'T12:00:00'); ; d.setDate(d.getDate() + 1)) {
    const f = Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
    if (f > fin) break;
    if (!esFinde(f) || porFecha[f]) calendario.push(f);   // hábiles + fines de semana con carga
  }
  // Alquileres mensuales: mes cerrado → prorrateo por los días de obra de ese mes; mes en curso → DIAS_MES.
  // Cada equipo cuenta solo desde su fecha DESDE (vacía = desde el inicio de obra); antes, sin costo.
  const meses = {};
  calendario.forEach(f => { const m = f.slice(0, 7); if (!meses[m]) meses[m] = { mes: m, diasObra: 0, conCarga: 0, cerrado: hoy.slice(0, 7) > m, equiposDia: eq.dia }; meses[m].diasObra++; if (porFecha[f] && porFecha[f].t > 0) meses[m].conCarga++; });
  const activo = (e, f) => !e.desde || f >= e.desde;
  // días de obra de cada equipo en cada mes (para prorratear meses cerrados)
  const diasEq = {};   // mes -> [días activos por equipo]
  calendario.forEach(f => { const m = f.slice(0, 7); if (!diasEq[m]) diasEq[m] = eq.detalle.map(() => 0); eq.detalle.forEach((e, i) => { if (activo(e, f)) diasEq[m][i]++; }); });
  const equiposDiaEn = f => {
    const m = f.slice(0, 7), cerrado = meses[m].cerrado;
    return eq.detalle.reduce((s, e, i) => !activo(e, f) ? s : s + (cerrado && diasEq[m][i] > 0 ? e.usdMes / diasEq[m][i] : e.usdMes / Number(cfg.DIAS_MES)), 0);
  };
  const equiposActivosEn = f => eq.detalle.filter(e => activo(e, f) && e.usdMes > 0).length;
  Object.values(meses).forEach(m => { const fs = calendario.filter(f => f.slice(0, 7) === m.mes); if (fs.length) m.equiposDia = fs.reduce((s, f) => s + equiposDiaEn(f), 0) / fs.length; });

  let acum = { t: 0, costo: 0, fact: 0, dias: 0, diasConCarga: 0, comb: 0, flete: 0, fijos: 0, puntas: 0 };
  const serie = calendario.map(f => {
    const d = porFecha[f] || { fecha: f, t: 0, viajes: null, obs: '', por: '', jornada: 0, combArs: null, sinCarga: true };
    const jornada = d.sinCarga ? 0 : d.jornada;
    const mesInfo = meses[f.slice(0, 7)];
    const equiposDiaUsado = equiposDiaEn(f);
    const fijos = equiposDiaUsado + Number(cfg.MO_USD) + Number(cfg.SENAL_USD);   // siempre
    const c = (!combDesde || f >= combDesde) ? comb.porDia[f] : null;
    const manual = d.combArs !== null && d.combArs !== undefined && !isNaN(d.combArs);
    const combUSD = manual ? d.combArs / combDiv / TC : (c ? c.ars / combDiv / TC : combEstUSD * jornada);
    const combFuente = manual ? 'manual' : (c ? 'real' : 'estimado');
    const viajes = d.viajes !== null && d.viajes !== undefined ? d.viajes : Math.ceil(d.t / Number(cfg.CAP_CAMION));
    const flete = viajes * fleteUSD;
    const puntas = d.t * puntasPorT;
    const costo = fijos + combUSD + flete + puntas;
    const fact = d.t * Number(cfg.PRECIO_T);
    acum.t += d.t; acum.costo += costo; acum.fact += fact; acum.dias += 1; if (d.t > 0) acum.diasConCarga += 1;
    acum.comb += combUSD; acum.flete += flete; acum.fijos += fijos; acum.puntas += puntas;
    return {
      fecha: f, t: d.t, viajes, jornada, sinCarga: !!d.sinCarga, finde: esFinde(f), obs: d.obs, por: d.por,
      fijos, equiposDia: equiposDiaUsado, equiposN: equiposActivosEn(f), diasMes: mesInfo.cerrado ? mesInfo.diasObra : Number(cfg.DIAS_MES), mesCerrado: mesInfo.cerrado, comb: combUSD, combReal: combFuente !== 'estimado', combFuente, combArs: manual ? d.combArs : (c ? c.ars : null), litros: c ? c.litros : null, combDet: c ? c.registros : [], flete, puntas,
      costo, fact, utilidad: fact - costo, costoT: d.t > 0 ? costo / d.t : null,
      acumCostoT: acum.t > 0 ? acum.costo / acum.t : null, acumUtil: acum.fact - acum.costo,
    };
  });

  return {
    cfg: { TC, IVA, COMB_CON_IVA: Number(cfg.COMB_CON_IVA) === 1 ? 1 : 0, PRECIO_T: Number(cfg.PRECIO_T), CAP: Number(cfg.CAP_CAMION), OBRA: cfg.OBRA_COMBUSTIBLE },
    base: { equiposDia: eq.dia, equiposMes: eq.mes, moDia: Number(cfg.MO_USD), senalDia: Number(cfg.SENAL_USD), fijoDia, fleteUSD, combEstUSD, puntasPorT, vidaPuntas, vidaReal, equipos: eq.detalle },
    serie, acum, recambios, meses: Object.values(meses), inicio,
    equiposTotal: eq.detalle.filter(e => e.usdMes > 0).length,
    combustible: comb.diag,
    user: user || '',
    hoy: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'),
  };
}

// ---------- Carga ----------
function guardarDia(p, user) {
  if (!user) throw new Error('Sin permiso');
  const fecha = String(p.fecha || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha inválida');
  const t = Number(p.t);
  if (isNaN(t) || t < 0 || t > 500) throw new Error('Toneladas fuera de rango (0–500)');
  const viajes = p.viajes === '' || p.viajes === null || p.viajes === undefined ? '' : Number(p.viajes);
  const jornada = p.jornada === 0 ? 0 : 1;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.PROD);
  const d = new Date(fecha + 'T12:00:00');
  const n = sh.getLastRow();
  let fila = -1;
  if (n > 1) {
    const fechas = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < fechas.length; i++) {
      if (fechas[i][0] && Utilities.formatDate(new Date(fechas[i][0]), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') === fecha) { fila = i + 2; break; }
    }
  }
  const combArs = p.comb === '' || p.comb === null || p.comb === undefined ? '' : Number(p.comb);
  if (combArs !== '' && (isNaN(combArs) || combArs < 0 || combArs > 50000000)) throw new Error('Combustible fuera de rango');
  const row = [d, t, viajes, String(p.obs || '').slice(0, 300), user, new Date(), jornada, combArs];
  if (sh.getLastColumn() < 8) sh.getRange(1, 8).setValue('COMBUSTIBLE_ARS').setFontWeight('bold');
  if (fila > 0) sh.getRange(fila, 1, 1, 8).setValues([row]);
  else sh.appendRow(row);
  let aviso = null;
  try { aviso = avisarProduccion_(fecha, user); } catch (err) { aviso = { ok: false, error: String(err.message || err) }; }
  return { ok: true, editado: fila > 0, aviso };
}

// ---------- Aviso por la app INGECO cuando se carga la producción del día ----------
const fmtUSD_ = v => (v < 0 ? '-' : '') + 'USD ' + Math.round(Math.abs(v)).toLocaleString('es-AR');
const fmtT_ = v => (Math.round(v * 10) / 10).toLocaleString('es-AR');
const fmtFechaAR_ = f => f.slice(8, 10) + '/' + f.slice(5, 7);

/** Arma los KPI del día y del acumulado y los manda a todos los que tienen el módulo AUTOVIAS en la app INGECO. */
function avisarProduccion_(fecha, user) {
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE_AUTOVIAS');
  if (!clave) return { ok: false, error: 'Sin CLAVE_AUTOVIAS' };
  const datos = getDatos(user);
  const d = datos.serie.find(x => x.fecha === fecha);
  if (!d) return { ok: false, error: 'día no encontrado' };
  const precio = datos.cfg.PRECIO_T, ac = datos.acum, base = datos.base;
  const variablePorT = base.puntasPorT + base.fleteUSD / Math.max(1, datos.cfg.CAP);
  const equilibrio = (base.fijoDia + d.comb) / Math.max(0.01, precio - variablePorT);
  const promTDia = ac.dias ? ac.t / ac.dias : 0;
  const titulo = 'Autovías ' + fmtFechaAR_(fecha) + ': ' + fmtT_(d.t) + ' t colocadas · utilidad ' + fmtUSD_(d.utilidad);
  const cuerpo = [
    'DÍA ' + fmtFechaAR_(fecha) + (d.jornada ? '' : ' (sin jornada)') + ': ' + fmtT_(d.t) + ' t · ' + d.viajes + ' viajes · costo ' + (d.costoT !== null ? fmtUSD_(d.costoT) + '/t' : '—') + ' (pactado USD ' + precio + '/t) · utilidad ' + fmtUSD_(d.utilidad),
    'Combustible: ' + (d.combFuente === 'manual' ? 'cargado a mano' : d.combFuente === 'real' ? 'planilla de combustible' : 'estimado') + ' · equilibrio hoy ≈ ' + fmtT_(equilibrio) + ' t/día',
    'ACUMULADO: ' + fmtT_(ac.t) + ' t en ' + ac.dias + ' días (' + fmtT_(promTDia) + ' t/día) · costo ' + (ac.t ? fmtUSD_(ac.costo / ac.t) + '/t' : '—') + ' · utilidad ' + fmtUSD_(ac.fact - ac.costo),
    'Cargó: ' + user
  ].join('\n');
  const r = UrlFetchApp.fetch(APP_INGECO_BACKEND, {
    method: 'post', contentType: 'text/plain', muteHttpExceptions: true,
    payload: JSON.stringify({
      accion: 'notificar', clave_sistema: clave, modulo: MODULO_INGECO,
      destinatarios: [{ con_modulo: MODULO_INGECO }, { admin: true }],
      tipo: 'aviso_general', titulo: titulo.slice(0, 120), cuerpo: cuerpo.slice(0, 600),
      url_destino: 'https://marcoskatz-cmd.github.io/obra-autovias/', prioridad: 'normal',
      clave_unica: 'prod-' + fecha + '-' + d.t + '-' + d.viajes + '-' + Math.round(d.utilidad), dias_unica: 1
    })
  });
  let j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) { }
  return j.ok ? { ok: true, destinatarios: (j.avisos || []).length } : { ok: false, error: r.getContentText().slice(0, 160) };
}

function guardarPuntas(p, user) {
  if (!user) throw new Error('Sin permiso');
  const fecha = String(p.fecha || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha inválida');
  const juegos = Number(p.juegos);
  if (!(juegos > 0 && juegos <= 10)) throw new Error('Cantidad de juegos inválida');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.PUNTAS)
    .appendRow([new Date(fecha + 'T12:00:00'), juegos, String(p.obs || '').slice(0, 300), user, new Date()]);
  return { ok: true };
}

function borrarDia(fecha, user) {
  if (!user) throw new Error('Sin permiso');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.PROD);
  const n = sh.getLastRow(); if (n < 2) return { ok: false };
  const fechas = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < fechas.length; i++) {
    if (fechas[i][0] && Utilities.formatDate(new Date(fechas[i][0]), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') === fecha) { sh.deleteRow(i + 2); return { ok: true }; }
  }
  return { ok: false };
}
