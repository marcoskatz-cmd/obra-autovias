/**
 * SEGUIMIENTO OBRA AUTOVÍAS — Bacheo con fresado
 * INGECO S.A. — Apps Script + Sheets
 *
 * Fuente de combustible: planilla "ENTREGAS DE COMBUSTIBLE - LEANDRO CASARES"
 * Acceso: solo los mails listados en CONFIG > ACCESO
 */

// ---------- IDs ----------
const ID_COMBUSTIBLE = '19eyY8MImPM_-Gyzj8QcqA_yR5KDrkuu-1faJUU48N1A'; // Leandro Casares (Shared Drive)
const HOJA_ENTREGAS = 'ENTREGAS';

const TABS = { CONFIG: 'CONFIG', PROD: 'PRODUCCION', PUNTAS: 'PUNTAS', EQUIPOS: 'EQUIPOS' };

// ---------- Web app ----------
function doGet() {
  const user = Session.getActiveUser().getEmail();
  if (!accesoPermitido_(user)) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:40px;text-align:center">' +
      '<h2>Sin acceso</h2><p>Esta herramienta es de uso restringido.<br>Cuenta: ' + (user || 'no identificada') + '</p></div>'
    ).setTitle('Sin acceso');
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Seguimiento Autovías — INGECO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function accesoPermitido_(email) {
  if (!email) return false;
  const cfg = leerConfig_();
  const lista = (cfg.ACCESO || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  return lista.includes(email.toLowerCase());
}

// ---------- Setup inicial (correr una vez desde el editor) ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let cfg = ss.getSheetByName(TABS.CONFIG) || ss.insertSheet(TABS.CONFIG);
  cfg.clear();
  cfg.getRange(1, 1, 1, 3).setValues([['CLAVE', 'VALOR', 'DESCRIPCIÓN']]).setFontWeight('bold');
  const rows = [
    ['ACCESO', 'marcoskatz@grupoingeco.com.ar, agustin@grupoingeco.com.ar', 'Mails con acceso, separados por coma'],
    ['OBRA_COMBUSTIBLE', 'autovia del noa', 'Texto con que Leandro identifica esta obra en la columna OBRA GENERAL (o LUGAR ENTREGA)'],
    ['TC', 1540, 'Tipo de cambio $/USD'],
    ['IVA', 0.21, 'IVA para netear'],
    ['COMB_CON_IVA', 1, '1 si el precio de la planilla de combustible incluye IVA'],
    ['DIAS_MES', 22, 'Días laborables por mes (divide los alquileres)'],
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
  eq.getRange(1, 1, 1, 5).setValues([['EQUIPO', 'PRECIO MES', 'MONEDA', 'CON IVA (1/0)', 'OBS']]).setFontWeight('bold');
  eq.getRange(2, 1, 11, 5).setValues([
    ['Fresadora Bomag BM 1000/20', 24145150, 'ARS', 1, 'sin repuestos ni consumibles'],
    ['Rodillo mixto 6 t', 7500000, 'ARS', 0, 'con maquinista'],
    ['Minicargador', 6700000, 'ARS', 0, ''],
    ['Retropala', 4000, 'USD', 0, ''],
    ['Terminadora', 3000, 'USD', 0, 'con operario'],
    ['Camión chasis apoyo', 6000000, 'ARS', 0, ''],
    ['Camión volcador', 0, 'ARS', 0, 'propio'],
    ['Acoplado liga', 1000000, 'ARS', 0, ''],
    ['Camioneta', 3500000, 'ARS', 0, ''],
    ['Minibús', 0, 'ARS', 0, 'propio'],
    ['Camión regador', 0, 'ARS', 0, 'propio'],
  ]);

  let prod = ss.getSheetByName(TABS.PROD) || ss.insertSheet(TABS.PROD);
  if (prod.getLastRow() === 0) {
    prod.getRange(1, 1, 1, 7).setValues([['FECHA', 'TONELADAS', 'VIAJES', 'OBSERVACIONES', 'CARGADO POR', 'TIMESTAMP', 'JORNADA']]).setFontWeight('bold');
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

function costoEquiposDia_(cfg) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.EQUIPOS);
  if (!sh || sh.getLastRow() < 2) return { dia: 0, detalle: [] };
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  let mes = 0; const detalle = [];
  v.forEach(r => {
    if (!r[0]) return;
    let usd = r[2] === 'USD' ? Number(r[1]) : Number(r[1]) / Number(cfg.TC);
    if (Number(r[3]) === 1) usd = usd / (1 + Number(cfg.IVA));
    mes += usd; detalle.push({ equipo: r[0], usdMes: usd });
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
    const clave = String(cfg.OBRA_COMBUSTIBLE || '').toLowerCase().trim();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const obra = String(r[iO] || '').toLowerCase(), lugar = String(iLug >= 0 ? r[iLug] : '').toLowerCase();
      if (!clave || (!obra.includes(clave) && !lugar.includes(clave))) continue;
      const d = aFecha_(r[iF]); if (!d) continue;
      const k = Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
      const litros = Number(r[iL]) || 0;
      const ars = aNumero_(r[iT]);
      if (!out[k]) out[k] = { litros: 0, ars: 0, registros: [] };
      out[k].litros += litros; out[k].ars += ars;
      out[k].registros.push({ equipo: String(iEq >= 0 ? r[iEq] : ''), litros, ars });
    }
    const total = Object.keys(out).length;
    diag.msg = total ? (total + ' días con carga de combustible para "' + cfg.OBRA_COMBUSTIBLE + '"') : ('Sin registros para "' + cfg.OBRA_COMBUSTIBLE + '" en la planilla de Leandro. Verificá CONFIG > OBRA_COMBUSTIBLE.');
    if (!total) diag.ok = false;
  } catch (e) { diag = { ok: false, msg: 'No se pudo leer combustible: ' + e.message }; }
  return { porDia: out, diag };
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
function getDatos() {
  const cfg = leerConfig_();
  const TC = Number(cfg.TC), IVA = Number(cfg.IVA);
  const eq = costoEquiposDia_(cfg);
  const fijoDia = eq.dia + Number(cfg.MO_USD) + Number(cfg.SENAL_USD);
  const fleteUSD = Number(cfg.FLETE_ARS) / TC;
  const combEstUSD = Number(cfg.COMB_ESTIMADO_ARS) / TC;
  const puntasJuegoUSD = Number(cfg.PUNTAS_JUEGO_ARS) / (1 + IVA) / TC;
  const comb = combustiblePorDia_(cfg);
  const combDiv = Number(cfg.COMB_CON_IVA) === 1 ? (1 + IVA) : 1;

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
    dias = shD.getRange(2, 1, shD.getLastRow() - 1, 7).getValues().filter(r => r[0])
      .map(r => ({ fecha: Utilities.formatDate(new Date(r[0]), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'), t: Number(r[1]) || 0, viajes: r[2] === '' ? null : Number(r[2]), obs: r[3], por: r[4], jornada: r[6] === '' ? 1 : Number(r[6]) }))
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

  let acum = { t: 0, costo: 0, fact: 0, dias: 0, comb: 0, flete: 0, fijos: 0, puntas: 0 };
  const serie = dias.map(d => {
    const jornada = d.jornada;
    const fijos = fijoDia * jornada;
    const c = comb.porDia[d.fecha];
    const combUSD = c ? c.ars / combDiv / TC : combEstUSD * jornada;
    const viajes = d.viajes !== null ? d.viajes : Math.ceil(d.t / Number(cfg.CAP_CAMION));
    const flete = viajes * fleteUSD;
    const puntas = d.t * puntasPorT;
    const costo = fijos + combUSD + flete + puntas;
    const fact = d.t * Number(cfg.PRECIO_T);
    acum.t += d.t; acum.costo += costo; acum.fact += fact; acum.dias += 1;
    acum.comb += combUSD; acum.flete += flete; acum.fijos += fijos; acum.puntas += puntas;
    return {
      fecha: d.fecha, t: d.t, viajes, jornada, obs: d.obs, por: d.por,
      fijos, comb: combUSD, combReal: !!c, litros: c ? c.litros : null, flete, puntas,
      costo, fact, utilidad: fact - costo, costoT: d.t > 0 ? costo / d.t : null,
      acumCostoT: acum.t > 0 ? acum.costo / acum.t : null, acumUtil: acum.fact - acum.costo,
    };
  });

  return {
    cfg: { TC, PRECIO_T: Number(cfg.PRECIO_T), CAP: Number(cfg.CAP_CAMION), OBRA: cfg.OBRA_COMBUSTIBLE },
    base: { equiposDia: eq.dia, equiposMes: eq.mes, moDia: Number(cfg.MO_USD), senalDia: Number(cfg.SENAL_USD), fijoDia, fleteUSD, combEstUSD, puntasPorT, vidaPuntas, vidaReal, equipos: eq.detalle },
    serie, acum, recambios,
    combustible: comb.diag,
    user: Session.getActiveUser().getEmail(),
    hoy: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'),
  };
}

// ---------- Carga ----------
function guardarDia(p) {
  const user = Session.getActiveUser().getEmail();
  if (!accesoPermitido_(user)) throw new Error('Sin permiso');
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
  const row = [d, t, viajes, String(p.obs || '').slice(0, 300), user, new Date(), jornada];
  if (fila > 0) sh.getRange(fila, 1, 1, 7).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, editado: fila > 0 };
}

function guardarPuntas(p) {
  const user = Session.getActiveUser().getEmail();
  if (!accesoPermitido_(user)) throw new Error('Sin permiso');
  const fecha = String(p.fecha || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha inválida');
  const juegos = Number(p.juegos);
  if (!(juegos > 0 && juegos <= 10)) throw new Error('Cantidad de juegos inválida');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.PUNTAS)
    .appendRow([new Date(fecha + 'T12:00:00'), juegos, String(p.obs || '').slice(0, 300), user, new Date()]);
  return { ok: true };
}

function borrarDia(fecha) {
  const user = Session.getActiveUser().getEmail();
  if (!accesoPermitido_(user)) throw new Error('Sin permiso');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.PROD);
  const n = sh.getLastRow(); if (n < 2) return { ok: false };
  const fechas = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < fechas.length; i++) {
    if (fechas[i][0] && Utilities.formatDate(new Date(fechas[i][0]), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') === fecha) { sh.deleteRow(i + 2); return { ok: true }; }
  }
  return { ok: false };
}
