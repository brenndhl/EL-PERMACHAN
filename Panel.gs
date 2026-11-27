/****************************************************************
 * PANEL DE ADMINISTRACIÓN — RESTAURANTE EL PERMACHAN
 * Todas las solicitudes (pendientes, emitidas, enviadas, rechazadas),
 * buscador, contadores y cambio de estado. Solo para administración:
 * usa una clave distinta a la de los enlaces de la contadora.
 *
 * INSTALACIÓN (una vez):
 *  1. Archivo + > Secuencia de comandos > nombre: Panel  (pega este archivo).
 *  2. Archivo + > HTML > nombre: PanelAdmin  (pega PanelAdmin.html; no puede llamarse igual que Panel.gs).
 *  3. Sustituye Admin.html por la versión nueva (solo añade 2 líneas, arriba y abajo).
 *  4. Implementar > Gestionar implementaciones > editar (lápiz) > Versión: Nueva versión > Implementar.
 *  5. Ejecuta enlacePanelAdmin() y copia el enlace del registro. Guárdalo en favoritos y NO lo reenvíes.
 *  6. (Opcional) Pon tu correo en Config > CORREO_ADMIN y ejecuta activarResumenDiario():
 *     cada mañana te llega un correo con lo pendiente.
 ****************************************************************/

/** Clave del panel de administración (distinta de la de la contadora). */
function claveDuena_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('CLAVE_DUENA');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); props.setProperty('CLAVE_DUENA', k); }
  return k;
}

function esDuena_(k) { return !!k && String(k) === claveDuena_(); }

/** Panel completo: la dueña (su clave) y la contadora (la clave de sus enlaces). */
function esPanel_(k) { return esDuena_(k) || claveOk_(k); }

/** Enlace para instalar la app con acceso de administración en el móvil de la contadora. */
function enlaceAppContadora() {
  const u = String(cfg_().URL_APP || 'https://brenndhl.github.io/EL-PERMACHAN/') + '#k=' + claveAdmin_();
  Logger.log('App de la contadora (abrir una vez en su móvil): ' + u);
  return u;
}

function urlPanelAdmin_() {
  const base = urlWeb_();
  return base ? base + '?k=' + encodeURIComponent(claveDuena_()) : '';
}

/** Escribe en el registro el enlace de TU panel. No lo compartas: ve todo el historial. */
function enlacePanelAdmin() {
  const u = urlPanelAdmin_();
  const txt = u ? 'Tu panel de administración (guárdalo en favoritos, no lo reenvíes): ' + u
    : 'Primero implementa la aplicación web y pega su URL en Config > URL_WEB.';
  Logger.log(txt);
  return u || txt;
}

/** Si el enlace del panel se filtra: crea una clave nueva (el enlace antiguo deja de funcionar). */
function cambiarClavePanelAdmin() {
  PropertiesService.getScriptProperties().deleteProperty('CLAVE_DUENA');
  return enlacePanelAdmin();
}

/* ===================== PAGOS "POR DEFINIR" ===================== */

const HOJA_PAGOS = 'Pagos';
const COL_PAGOS = ['ID', 'Fecha pago', 'RFC', 'Cliente', 'Folio ticket', 'Nota', 'Aviso a contadora', 'Marcado el'];

function hojaPagos_() {
  const ss = ss_();
  let sh = ss.getSheetByName(HOJA_PAGOS);
  if (!sh) {
    sh = ss.insertSheet(HOJA_PAGOS);
    sh.getRange(1, 1, 1, COL_PAGOS.length).setValues([COL_PAGOS]).setFontWeight('bold').setBackground('#1f4e5f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
  }
  return sh;
}

function pagosMap_() {
  const m = {};
  leerTabla_(hojaPagos_()).forEach(p => { if (p['ID']) m[String(p['ID'])] = p; });
  return m;
}

function porDefinir_(s) { return /definir/i.test(String(s['Forma de pago'] || '')); }

function correoContadora_() {
  return String(cfg_().CORREO_CONTADORA || 'jovhys@hotmail.com').split(',')[0].trim();
}

/** 'yyyy-mm-dd' → Date (00:00 en la zona del script); vacío o inválido → null */
function fechaDe_(t) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(t || ''));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

/** Marca como pagada una factura "por definir" y avisa a la contadora por correo. */
function panelMarcarPagada(k, id, fechaTxt, nota) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const f = buscarSolicitud_(id);
    if (!f) return { ok: false, error: 'No encuentro la solicitud ' + id + '.' };
    if (pagosMap_()[String(id)]) return { ok: true, yaEstaba: true };
    const s = f.datos;
    const fecha = fechaDe_(fechaTxt) || new Date();
    nota = String(nota || '').trim().slice(0, 300);
    const para = correoContadora_();
    const tz = Session.getScriptTimeZone();
    const filas = [['Solicitud', s['ID']], ['Cliente', s['Nombre o razón social']], ['RFC', s['RFC']], ['Ticket', s['Folio ticket']],
      ['Fecha de la solicitud', s['Fecha solicitud'] instanceof Date ? Utilities.formatDate(s['Fecha solicitud'], tz, 'dd/MM/yyyy') : s['Fecha solicitud']],
      ['Estado de la factura', s['Estado']], ['Fecha de pago', Utilities.formatDate(fecha, tz, 'dd/MM/yyyy')], ['Nota', nota || '—']]
      .map(([a, b]) => '<tr><td><b>' + esc_(a) + '</b></td><td>' + esc_(b) + '</td></tr>').join('');
    const enviado = enviarCorreo_(para, 'Factura pagada — ticket ' + s['Folio ticket'] + ' — ' + (s['Nombre o razón social'] || s['RFC']),
      '<p>Hola, el cliente ya pagó esta factura que estaba como <b>pago por definir</b>:</p>' +
      '<table border="1" cellpadding="6" style="border-collapse:collapse">' + filas + '</table>' +
      '<p>Si corresponde, emite el complemento de pago.</p><p>' + esc_(cfg_().NOMBRE_NEGOCIO) + '</p>');
    agregarFila_(hojaPagos_(), COL_PAGOS, { 'ID': s['ID'], 'Fecha pago': fecha, 'RFC': s['RFC'], 'Cliente': s['Nombre o razón social'],
      'Folio ticket': s['Folio ticket'], 'Nota': nota, 'Aviso a contadora': enviado ? 'Enviado a ' + para : 'No se pudo enviar', 'Marcado el': new Date() });
    return { ok: true, avisado: enviado, para: para };
  } finally { lock.releaseLock(); }
}

/** Quita la marca de pagada (por si fue un error). No manda correo. */
function panelQuitarPagada(k, id) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = hojaPagos_();
    if (sh.getLastRow() < 2) return { ok: true };
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ===================== DATOS DEL PANEL ===================== */

/** Contadores + lista filtrada (máx. 200 filas).
 *  f = { estado, texto, desde:'yyyy-mm-dd', hasta:'yyyy-mm-dd', pago, cliente:RFC, porCobrar:bool } */
function panelDatos(k, f) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  f = f || {};
  const tz = Session.getScriptTimeZone();
  const ahora = Date.now();
  const todas = leerTabla_(hoja_(HOJAS.SOLICITUDES)).filter(s => s['ID']);
  const pagos = pagosMap_();
  let facts = {};
  try { facts = facturasMap_(); } catch (e) { /* sin hoja de facturas */ }
  const fecha = s => s['Fecha solicitud'] instanceof Date ? s['Fecha solicitud'].getTime() : 0;
  const dias = s => fecha(s) ? Math.floor((ahora - fecha(s)) / 86400000) : null;
  const pagada = s => !!pagos[String(s['ID'])];
  const porCobrar = s => porDefinir_(s) && !pagada(s) && s['Estado'] !== 'Rechazada';

  const inicioMes = new Date(); inicioMes.setDate(1); inicioMes.setHours(0, 0, 0, 0);
  const kpi = {
    atrasadas: todas.filter(s => s['Estado'] === 'Pendiente' && dias(s) >= 2).length,
    mes: todas.filter(s => fecha(s) >= inicioMes.getTime()).length,
    enviadasMes: todas.filter(s => (s['Estado'] === 'Enviada' || s['Estado'] === 'Emitida') &&
      s['Fecha emitida'] instanceof Date && s['Fecha emitida'] >= inicioMes).length,
    porCobrar: todas.filter(porCobrar).length,
    clientesRevisar: leerTabla_(hoja_(HOJAS.CLIENTES)).filter(c => c['Revisión'] === 'Pendiente de revisar').length
  };

  // Lista de clientes para el filtro
  const vistos = {};
  todas.forEach(s => { const r = String(s['RFC'] || '').trim(); if (r && !vistos[r]) vistos[r] = s['Nombre o razón social'] || r; });
  const clientes = Object.keys(vistos).map(r => [r, vistos[r]]).sort((a, b) => String(a[1]).localeCompare(String(b[1]), 'es'));

  // Filtros (todos menos el estado)
  const desde = fechaDe_(f.desde), hasta = fechaDe_(f.hasta);
  const hastaFin = hasta ? hasta.getTime() + 86400000 : 0;
  let base = todas;
  if (desde) base = base.filter(s => fecha(s) >= desde.getTime());
  if (hasta) base = base.filter(s => fecha(s) && fecha(s) < hastaFin);
  if (f.pago) base = base.filter(s => String(s['Forma de pago']) === f.pago);
  if (f.cliente) base = base.filter(s => String(s['RFC']).trim() === f.cliente);
  if (f.porCobrar) base = base.filter(porCobrar);
  const q = sinAcentos_(f.texto || '').toUpperCase().trim();
  if (q) base = base.filter(s => sinAcentos_([s['ID'], s['RFC'], s['Nombre o razón social'], s['Email'], s['Folio ticket']].join(' '))
    .toUpperCase().indexOf(q) >= 0);

  const conteos = { Todas: base.length };
  ESTADOS.forEach(e => { conteos[e] = 0; });
  base.forEach(s => { if (conteos[s['Estado']] !== undefined) conteos[s['Estado']]++; });

  const estado = f.estado || 'Pendiente';
  let lista = estado === 'Todas' ? base : base.filter(s => s['Estado'] === estado);
  lista.sort((a, b) => estado === 'Pendiente' ? fecha(a) - fecha(b) : fecha(b) - fecha(a));

  const fmt = d => d instanceof Date ? Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm') : String(d || '');
  const total = lista.length;
  lista = lista.slice(0, 200).map(s => {
    const p = pagos[String(s['ID'])];
    return Object.assign(resumenSolicitud_(s), {
      dias: dias(s), fechaEstado: fmt(s['Fecha estado']), fechaEmitida: fmt(s['Fecha emitida']),
      notas: s['Notas / motivo'], origen: s['Origen'], uso: s['Uso CFDI'],
      porDefinir: porDefinir_(s), pagada: !!p,
      factura: facts[String(s['ID'])] ? (facts[String(s['ID'])]['Serie-Folio'] || 'sí') : '',
      fechaPago: p && p['Fecha pago'] instanceof Date ? Utilities.formatDate(p['Fecha pago'], tz, 'dd/MM/yyyy') : ''
    });
  });
  return { ok: true, conteos: conteos, kpi: kpi, total: total, lista: lista, estado: estado, clientes: clientes,
    formasPago: FORMAS_PAGO, hoy: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), correoContadora: correoContadora_() };
}

/** Cambia el estado desde el panel. Aquí sí se puede volver a "Pendiente" (reabrir). */
function panelCambiarEstado(k, id, estado, motivo) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  if (ESTADOS.indexOf(estado) < 0) return { ok: false, error: 'Estado no válido.' };
  motivo = String(motivo || '').trim().slice(0, 300);
  if (estado === 'Rechazada' && motivo.length < 3) return { ok: false, error: 'Escribe el motivo para que el cliente sepa qué corregir.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const f = buscarSolicitud_(id);
    if (!f) return { ok: false, error: 'No encuentro la solicitud ' + id + '.' };
    if (f.datos['Estado'] === estado) return { ok: true, yaEstaba: true, estado: estado };
    const sh = hoja_(HOJAS.SOLICITUDES);
    const cambios = { 'Estado': estado };
    if (motivo) cambios['Notas / motivo'] = motivo;
    actualizarFila_(sh, COL_SOLICITUDES, f.fila, cambios);
    aplicarEstado_(sh, f.fila);   // fechas + correo al cliente (al reabrir no se le escribe)
    return { ok: true, estado: estado };
  } finally { lock.releaseLock(); }
}

/** Ver el PDF o XML de la factura subida por la contadora. */
function panelFactura(k, id, tipo) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const f = facturasMap_()[String(id)];
  if (!f) return { ok: false, error: 'Esta solicitud no tiene factura subida.' };
  const b = blobDeUrl_(tipo === 'xml' ? f['XML'] : f['PDF'], 'Factura_' + (f['Serie-Folio'] || id));
  if (!b) return { ok: false, error: 'No se pudo abrir el archivo.' };
  return { ok: true, tipo: b.getContentType(), nombre: b.getName(), base64: Utilities.base64Encode(b.getBytes()) };
}

/** Ver ticket / constancia desde el panel. */
function panelArchivo(k, id, cual) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  return adminArchivo(claveAdmin_(), id, cual);
}

/* ===================== RESUMEN DIARIO POR CORREO ===================== */

/** Correo con las pendientes (más antiguas primero). Si no hay pendientes, no manda nada. */
function resumenDiario() {
  const para = String(cfg_().CORREO_ADMIN || '').trim();
  if (!para) return 'Config > CORREO_ADMIN está vacío.';
  const tz = Session.getScriptTimeZone(), ahora = Date.now();
  const todas = leerTabla_(hoja_(HOJAS.SOLICITUDES)).filter(s => s['ID']);
  const t = s => s['Fecha solicitud'] instanceof Date ? s['Fecha solicitud'].getTime() : 0;
  const pend = todas.filter(s => s['Estado'] === 'Pendiente').sort((a, b) => t(a) - t(b));
  const ayer = ahora - 86400000;
  const hechas = todas.filter(s => s['Fecha estado'] instanceof Date && s['Fecha estado'].getTime() >= ayer &&
    (s['Estado'] === 'Enviada' || s['Estado'] === 'Emitida')).length;
  if (!pend.length) return 'Sin pendientes: no se manda resumen.';
  const atras = pend.filter(s => t(s) && ahora - t(s) >= 2 * 86400000).length;
  const filas = pend.slice(0, 20).map(s => {
    const d = t(s) ? Math.floor((ahora - t(s)) / 86400000) : '';
    const rojo = d !== '' && d >= 2 ? ' style="color:#b3261e;font-weight:bold"' : '';
    return '<tr><td>' + esc_(s['ID']) + '</td><td>' + (t(s) ? Utilities.formatDate(s['Fecha solicitud'], tz, 'dd/MM HH:mm') : '') +
      '</td><td' + rojo + '>' + d + '</td><td>' + esc_(s['Nombre o razón social'] || s['RFC']) + '</td><td>' + esc_(s['Folio ticket']) + '</td></tr>';
  }).join('');
  const link = urlPanelAdmin_();
  const html = '<p><b>' + pend.length + '</b> solicitudes pendientes' + (atras ? ' — <b style="color:#b3261e">' + atras + ' con 2 días o más</b>' : '') +
    '.<br>Enviadas/emitidas en las últimas 24 h: <b>' + hechas + '</b>.</p>' +
    '<table border="1" cellpadding="6" style="border-collapse:collapse"><tr><th>ID</th><th>Fecha</th><th>Días</th><th>Cliente</th><th>Folio</th></tr>' +
    filas + '</table>' + (pend.length > 20 ? '<p>…y ' + (pend.length - 20) + ' más.</p>' : '') +
    (link ? '<p style="margin-top:18px"><a href="' + esc_(link) + '" style="font-size:18px">Abrir el panel</a></p>' +
      '<p style="font-size:12px;color:#777">Enlace privado: no reenvíes este correo.</p>' : '');
  enviarCorreo_(para, 'Facturas: ' + pend.length + ' pendientes' + (atras ? ' (' + atras + ' atrasadas)' : ''), html);
  return 'Resumen enviado a ' + para;
}

/** Crea la fila CORREO_ADMIN en Config (si falta) y el disparador diario a las 9:00 (zona horaria del script). */
function activarResumenDiario() {
  const sh = hoja_(HOJAS.CONFIG);
  const claves = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  if (claves.indexOf('CORREO_ADMIN') < 0) sh.appendRow(['CORREO_ADMIN', '']);
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'resumenDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('resumenDiario').timeBased().everyDays(1).atHour(9).create();
  const txt = cfg_().CORREO_ADMIN
    ? 'Resumen diario activado (9:00, zona ' + Session.getScriptTimeZone() + ') para ' + cfg_().CORREO_ADMIN + '.'
    : 'Disparador creado. Falta tu correo en Config > CORREO_ADMIN (fila nueva al final).';
  Logger.log(txt);
  return txt;
}

/* ===================== CORREO DE LA CONTADORA ===================== */

function setConfig_(clave, valor) {
  const sh = hoja_(HOJAS.CONFIG);
  const n = sh.getLastRow();
  const claves = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  const i = claves.indexOf(clave);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(valor); else sh.appendRow([clave, valor]);
}

/** Pone jovhys@hotmail.com como correo de la contadora en Config (avisos de pago e informe quincenal). */
function configurarCorreoContadora() {
  const correo = 'jovhys@hotmail.com';
  setConfig_('CORREO_CONTADORA', correo);
  const inf = String(cfg_().CORREOS_INFORME || '').split(',').map(x => x.trim()).filter(x => x && !/jhovis@gmail\.com/i.test(x));
  if (inf.map(x => x.toLowerCase()).indexOf(correo) < 0) inf.push(correo);
  setConfig_('CORREOS_INFORME', inf.join(', '));
  hojaPagos_();
  const txt = 'Contadora: ' + correo + ' · Informe quincenal: ' + inf.join(', ') + ' · Avisos: ' + cfg_().CORREOS_AVISO;
  Logger.log(txt);
  return txt;
}
