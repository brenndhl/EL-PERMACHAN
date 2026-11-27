/****************************************************************
 * EXTRAS — RESTAURANTE EL PERMACHAN
 *  · Correos por cliente (varios correos por RFC; añadir / quitar)
 *  · Gestión de clientes desde el panel (alta, edición, borrar)
 *  · Constancia (CSF): fecha y aviso de renovación cada 6 meses
 *  · Ayuda: comentarios (errores, mejoras, dudas), valoraciones y
 *    "¿Quieres tu propia app?"; validar → issue de GitHub
 *  · Caja: pagos a proveedores y pedidos (clave de la cajera)
 *
 * INSTALACIÓN: ejecutar instalarExtras() una vez. Luego enlaceAppCaja()
 * da el enlace para el móvil de la cajera.
 ****************************************************************/

const REPO_GITHUB_ = 'brenndhl/EL-PERMACHAN';
const CORREO_CONTACTO_APPS_ = 'elpermachan@gmail.com';
const DIAS_CSF_ = 182;   // la constancia se pide renovar cada ~6 meses

/* ============================ UTILIDADES ============================ */

function hojaCon_(nombre, cols) {
  const ss = ss_();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#1f4e5f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function sigId_(prop, prefijo) {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty(prop) || 1);
  props.setProperty(prop, String(n + 1));
  return prefijo + ('0000' + n).slice(-4);
}

function fmtFecha_(d, conHora) {
  return d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), conHora ? 'dd/MM/yyyy HH:mm' : 'dd/MM/yyyy') : String(d || '');
}

/** Columna de la hoja Clientes por nombre (la crea al final si no existe). */
function colCliente_(nombre) {
  const sh = hoja_(HOJAS.CLIENTES);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  let i = cab.indexOf(nombre);
  if (i < 0) {
    i = cab.length;
    sh.getRange(1, i + 1).setValue(nombre).setFontWeight('bold').setBackground('#1f4e5f').setFontColor('#ffffff');
  }
  return i + 1;
}

function setCliente_(fila, cambios) {
  const sh = hoja_(HOJAS.CLIENTES);
  Object.keys(cambios).forEach(k => sh.getRange(fila, colCliente_(k)).setValue(celda_(cambios[k])));
}

/** Datos del cliente con TODAS las columnas (también las añadidas: correos extra, fecha constancia…). */
function clienteCompleto_(rfc) {
  const c = buscarCliente_(rfc);
  if (!c) return null;
  const sh = hoja_(HOJAS.CLIENTES);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const fila = sh.getRange(c.fila, 1, 1, cab.length).getValues()[0];
  return { fila: c.fila, datos: objeto_(cab, fila) };
}

function listaCorreos_(txt) {
  return String(txt || '').split(/[,;\s]+/).map(e => e.trim().toLowerCase()).filter(emailValido_);
}

/* ============================ CORREOS DEL CLIENTE ============================ */

/** Correos válidos de un RFC: el registrado + los añadidos a mano + los usados en sus solicitudes, menos los quitados. */
function correosCliente_(rfc, datos) {
  if (!datos || datos['Correos extra'] === undefined) { const cc = clienteCompleto_(rfc); datos = cc ? cc.datos : (datos || {}); }
  const quitados = listaCorreos_(datos['Correos quitados']);
  const lista = [];
  const add = e => { e = String(e || '').trim().toLowerCase(); if (emailValido_(e) && lista.indexOf(e) < 0 && quitados.indexOf(e) < 0) lista.push(e); };
  add(datos['Email']);
  listaCorreos_(datos['Correos extra']).forEach(add);
  const sh = hoja_(HOJAS.SOLICITUDES);
  if (sh.getLastRow() > 1) {
    const iR = COL_SOLICITUDES.indexOf('RFC'), iE = COL_SOLICITUDES.indexOf('Email');
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, COL_SOLICITUDES.length).getValues();
    for (let i = vals.length - 1; i >= 0 && lista.length < 8; i--) if (normalizarRFC(vals[i][iR]) === rfc) add(vals[i][iE]);
  }
  return lista;
}

/* ============================ CONSTANCIA (CSF) ============================ */

/** Fecha de la constancia: la guardada o, si no hay, la de alta / última actualización. */
function fechaCSF_(d) {
  if (!d || !d['Constancia']) return null;
  const f = d['Fecha constancia'] || d['Última actualización'] || d['Fecha alta'];
  return f instanceof Date ? f : null;
}

function csfVencida_(d) {
  const f = fechaCSF_(d);
  return !f || (Date.now() - f.getTime()) > DIAS_CSF_ * 86400000;
}

/** Mensual: recuerda a los clientes activos (con solicitudes en los últimos 6 meses) que actualicen su constancia. Máx. 1 aviso cada 6 meses. */
function recordatorioCSF() {
  const limite = Date.now() - DIAS_CSF_ * 86400000;
  const activos = {};
  leerTabla_(hoja_(HOJAS.SOLICITUDES)).forEach(s => {
    if (s['Fecha solicitud'] instanceof Date && s['Fecha solicitud'].getTime() > limite) activos[normalizarRFC(s['RFC'])] = true;
  });
  const sh = hoja_(HOJAS.CLIENTES);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const filas = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, cab.length).getValues() : [];
  const app = String(cfg_().URL_APP || 'https://brenndhl.github.io/EL-PERMACHAN/') + '?v=mis';
  let n = 0;
  filas.forEach((f, i) => {
    const d = objeto_(cab, f), rfc = normalizarRFC(d['RFC']);
    if (!activos[rfc] || !csfVencida_(d) || !emailValido_(d['Email'])) return;
    if (d['Aviso CSF'] instanceof Date && d['Aviso CSF'].getTime() > limite) return;
    if (MailApp.getRemainingDailyQuota() < 20) return;
    enviarCorreo_(d['Email'], 'Actualiza tu Constancia de Situación Fiscal — ' + cfg_().NOMBRE_NEGOCIO,
      '<p>Hola, para que tus facturas salgan siempre con tus datos correctos, te pedimos que actualices tu <b>Constancia de Situación Fiscal</b> (la pedimos cada 6 meses).</p>' +
      '<p>Es muy fácil: entra en la app, en <b>Mis facturas → Mis datos</b>, y sube la constancia nueva (foto o PDF).</p>' +
      '<p style="text-align:center;margin:22px 0"><a href="' + esc_(app) + '" style="background:#141110;color:#f3ead6;text-decoration:none;font-weight:bold;padding:12px 24px;border-radius:10px">Actualizar mi constancia</a></p>' +
      '<p>Gracias.<br>' + esc_(cfg_().NOMBRE_NEGOCIO) + '</p>');
    setCliente_(i + 2, { 'Aviso CSF': new Date() });
    n++;
  });
  Logger.log('Recordatorios de constancia enviados: ' + n);
  return n;
}

/* ============================ ÁREA DEL CLIENTE: MIS DATOS ============================ */

/** Datos fiscales del cliente que ha entrado (para la pantalla "Mis datos"). */
function clienteMisDatos(token) {
  const rfc = rfcDeToken_(token);
  if (!rfc) return { ok: false, sesion: false, error: 'Tu sesión caducó. Vuelve a entrar con tu RFC.' };
  const c = clienteCompleto_(rfc);
  if (!c) return { ok: false, error: 'No encontramos tu RFC.' };
  const d = c.datos, f = fechaCSF_(d);
  return { ok: true, datos: { rfc: rfc, nombre: d['Nombre o razón social'], cp: String(d['CP fiscal'] || ''), regimen: String(d['Régimen fiscal'] || ''),
    uso: d['Uso CFDI'] || 'G03', tieneCSF: !!d['Constancia'], fechaCSF: f ? fmtFecha_(f) : '', csfVencida: csfVencida_(d),
    correo: emailDeToken_(token) }, regimenes: REGIMENES, usos: USOS_CFDI, maxMb: Number(cfg_().TAMANO_MAX_MB) || 8 };
}

/** Guarda cambios de datos fiscales y/o una constancia nueva. Queda "Pendiente de revisar" y avisa a administración. */
function clienteGuardarDatos(token, n) {
  const rfc = rfcDeToken_(token);
  if (!rfc) return { ok: false, sesion: false, error: 'Tu sesión caducó. Vuelve a entrar con tu RFC.' };
  const c = clienteCompleto_(rfc);
  if (!c) return { ok: false, error: 'No encontramos tu RFC.' };
  n = n || {};
  const d = c.datos;
  const datos = { nombre: n.nombre || d['Nombre o razón social'], cp: n.cp || d['CP fiscal'], regimen: n.regimen || d['Régimen fiscal'],
    uso: n.uso || d['Uso CFDI'] || 'G03', email: emailValido_(d['Email']) ? d['Email'] : 'cliente@correo.mx' };
  const v = validarDatosCliente_(datos, false);
  if (v) return { ok: false, error: v };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const nuevos = { 'Nombre o razón social': limpiarNombre_(datos.nombre), 'CP fiscal': String(datos.cp).trim(), 'Régimen fiscal': datos.regimen, 'Uso CFDI': datos.uso };
    const cambios = [];
    Object.keys(nuevos).forEach(k => { if (String(d[k] || '') !== String(nuevos[k])) cambios.push([k, d[k], nuevos[k]]); });
    if (n.constancia && n.constancia.base64) {
      nuevos['Constancia'] = guardarArchivo_(n.constancia, rfc + '_constancia');
      nuevos['Fecha constancia'] = new Date();
      cambios.push(['Constancia', d['Constancia'] ? 'anterior' : '', 'nueva']);
    }
    if (!cambios.length) return { ok: true, mensaje: 'No había cambios.' };
    const hoy = new Date();
    cambios.forEach(ch => agregarFila_(hoja_(HOJAS.CAMBIOS), COL_CAMBIOS,
      { 'Fecha': hoy, 'RFC': rfc, 'Campo': ch[0], 'Valor anterior': ch[1], 'Valor nuevo': ch[2], 'Estado': 'Pendiente de revisar', 'Origen': 'App' }));
    nuevos['Revisión'] = 'Pendiente de revisar';
    nuevos['Última actualización'] = hoy;
    setCliente_(c.fila, nuevos);
    enviarCorreo_(cfg_().CORREOS_AVISO, 'Datos fiscales actualizados desde la app — ' + rfc,
      '<p>El cliente <b>' + esc_(rfc) + '</b> (' + esc_(emailDeToken_(token)) + ') ha actualizado: ' +
      esc_(cambios.map(ch => ch[0]).join(', ')) + '. Revísalo en el panel (Clientes) o en la hoja <i>Cambios</i>.</p>');
    return { ok: true, mensaje: 'Datos guardados. Los revisaremos antes de tu próxima factura.' };
  } finally { lock.releaseLock(); }
}

/* ============================ PANEL: CLIENTES ============================ */

function filaClientePanel_(d) {
  const f = fechaCSF_(d);
  return { rfc: normalizarRFC(d['RFC']), tipo: d['Tipo'], nombre: d['Nombre o razón social'], cp: String(d['CP fiscal'] || ''),
    regimen: String(d['Régimen fiscal'] || ''), uso: d['Uso CFDI'] || '', email: d['Email'] || '',
    extra: listaCorreos_(d['Correos extra']).join(', '), quitados: listaCorreos_(d['Correos quitados']).join(', '),
    constancia: !!d['Constancia'], fechaCSF: f ? fmtFecha_(f) : '', csfVencida: csfVencida_(d), revision: d['Revisión'] || '',
    alta: fmtFecha_(d['Fecha alta']) };
}

/** Buscar clientes (por nombre, RFC o correo). Sin texto: los últimos actualizados. */
function panelClientes(k, q) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const sh = hoja_(HOJAS.CLIENTES);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  let todos = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, cab.length).getValues().map(f => objeto_(cab, f)).filter(d => d['RFC']) : [];
  const t = sinAcentos_(q || '').toUpperCase().trim();
  if (t) todos = todos.filter(d => sinAcentos_([d['RFC'], d['Nombre o razón social'], d['Email'], d['Correos extra']].join(' ')).toUpperCase().indexOf(t) >= 0);
  const ts = d => (d['Última actualización'] instanceof Date ? d['Última actualización'].getTime() : 0);
  todos.sort((a, b) => ts(b) - ts(a));
  return { ok: true, total: todos.length, lista: todos.slice(0, 60).map(filaClientePanel_), regimenes: REGIMENES, usos: USOS_CFDI };
}

/** Alta o edición de un cliente desde el panel. c = {rfcOriginal?, rfc, nombre, cp, regimen, uso, email, extra, quitados} */
function panelClienteGuardar(k, c) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  c = c || {};
  const rfc = normalizarRFC(c.rfc);
  if (!rfcValido(rfc)) return { ok: false, error: 'RFC no válido.' };
  if (limpiarNombre_(c.nombre).length < 3) return { ok: false, error: 'Escribe el nombre o razón social.' };
  if (c.cp && !/^\d{5}$/.test(String(c.cp).trim())) return { ok: false, error: 'El CP fiscal debe tener 5 dígitos.' };
  if (c.email && !emailValido_(c.email)) return { ok: false, error: 'El correo principal no es válido.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const orig = normalizarRFC(c.rfcOriginal || '');
    const existente = buscarCliente_(rfc);
    if (existente && orig !== rfc) return { ok: false, error: 'Ya existe un cliente con el RFC ' + rfc + '.' };
    const datos = { 'RFC': rfc, 'Tipo': tipoRFC(rfc), 'Nombre o razón social': limpiarNombre_(c.nombre), 'CP fiscal': String(c.cp || '').trim(),
      'Régimen fiscal': c.regimen || '', 'Uso CFDI': c.uso || 'G03', 'Email': String(c.email || '').trim().toLowerCase(),
      'Correos extra': listaCorreos_(c.extra).join(', '), 'Correos quitados': listaCorreos_(c.quitados).join(', '),
      'Revisión': 'Revisado', 'Última actualización': new Date() };
    if (orig) {
      const ant = buscarCliente_(orig);
      if (!ant) return { ok: false, error: 'No encuentro el cliente ' + orig + '.' };
      setCliente_(ant.fila, datos);
      return { ok: true, mensaje: 'Cliente actualizado.' };
    }
    datos['Fecha alta'] = new Date(); datos['Origen'] = 'Panel';
    agregarFila_(hoja_(HOJAS.CLIENTES), COL_CLIENTES, datos);
    const nuevo = buscarCliente_(rfc);
    if (nuevo) setCliente_(nuevo.fila, { 'Correos extra': datos['Correos extra'], 'Correos quitados': datos['Correos quitados'] });
    return { ok: true, mensaje: 'Cliente dado de alta.' };
  } finally { lock.releaseLock(); }
}

/** Borrar un cliente (no borra sus solicitudes ni sus facturas). */
function panelClienteBorrar(k, rfcTexto) {
  if (!esDuena_(k)) return { ok: false, error: 'Solo la administradora puede borrar clientes.' };
  const c = buscarCliente_(normalizarRFC(rfcTexto));
  if (!c) return { ok: false, error: 'No encuentro ese cliente.' };
  hoja_(HOJAS.CLIENTES).deleteRow(c.fila);
  return { ok: true, mensaje: 'Cliente borrado. Sus solicitudes y facturas se conservan.' };
}

/** Marcar como revisado (datos o constancia nuevos). */
function panelClienteRevisado(k, rfcTexto) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const c = buscarCliente_(normalizarRFC(rfcTexto));
  if (!c) return { ok: false, error: 'No encuentro ese cliente.' };
  setCliente_(c.fila, { 'Revisión': 'Revisado' });
  return { ok: true };
}

/** Ver la constancia de un cliente desde el panel. */
function panelClienteCSF(k, rfcTexto) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const c = buscarCliente_(normalizarRFC(rfcTexto));
  const b = c ? blobDeUrl_(c.datos['Constancia'], 'Constancia_' + c.datos['RFC']) : null;
  if (!b) return { ok: false, error: 'No hay constancia o no se pudo abrir.' };
  return { ok: true, tipo: b.getContentType(), nombre: b.getName(), base64: Utilities.base64Encode(b.getBytes()) };
}

/* ============================ AYUDA: COMENTARIOS ============================ */

const COL_COMENT = ['ID', 'Fecha', 'Tipo', 'Estrellas', 'Texto', 'Nombre', 'Contacto', 'RFC', 'Estado', 'Issue', 'Notas'];
function hojaComent_() { return hojaCon_('Comentarios', COL_COMENT); }

/** Desde la pestaña Ayuda. tipo: Error | Mejora | Duda | Valoración | Quiero mi app */
function enviarComentario(c, token) {
  c = c || {};
  const tipos = ['Error', 'Mejora', 'Duda', 'Valoración', 'Quiero mi app'];
  const tipo = tipos.indexOf(c.tipo) >= 0 ? c.tipo : 'Mejora';
  const texto = String(c.texto || '').trim().slice(0, 2000);
  const estrellas = Math.max(0, Math.min(5, Number(c.estrellas) || 0));
  const contacto = String(c.contacto || '').trim().slice(0, 120);
  if (tipo === 'Valoración' ? !estrellas : texto.length < 3) return { ok: false, error: tipo === 'Valoración' ? 'Elige de 1 a 5 estrellas.' : 'Escribe tu mensaje.' };
  if (tipo === 'Quiero mi app' && contacto.length < 6) return { ok: false, error: 'Déjanos un teléfono o correo para contactarte.' };
  const cache = CacheService.getScriptCache();
  const nHora = Number(cache.get('coment_hora') || 0);
  if (nHora > 60) return { ok: false, error: 'Hay muchos mensajes ahora mismo. Inténtalo en un rato.' };
  cache.put('coment_hora', String(nHora + 1), 3600);
  const rfc = token ? rfcDeToken_(token) : '';
  const fila = { 'ID': sigId_('SIG_COMENT', 'C-'), 'Fecha': new Date(), 'Tipo': tipo, 'Estrellas': estrellas || '', 'Texto': texto,
    'Nombre': String(c.nombre || '').trim().slice(0, 100), 'Contacto': contacto, 'RFC': rfc, 'Estado': 'Nuevo', 'Issue': '', 'Notas': '' };
  agregarFila_(hojaComent_(), COL_COMENT, fila);
  if (tipo === 'Quiero mi app') {
    enviarCorreo_(CORREO_CONTACTO_APPS_, '📲 Alguien quiere su propia app — ' + (fila['Nombre'] || contacto),
      '<p>Desde la app de El Permachan han pedido información para tener <b>su propia app</b>:</p><table border="1" cellpadding="6" style="border-collapse:collapse">' +
      [['Nombre', fila['Nombre']], ['Contacto', contacto], ['Mensaje', texto]].map(r => '<tr><td><b>' + r[0] + '</b></td><td>' + esc_(r[1]) + '</td></tr>').join('') + '</table>');
  }
  return { ok: true, mensaje: tipo === 'Valoración' ? '¡Gracias por tu valoración!' : tipo === 'Quiero mi app' ? '¡Gracias! Te contactaremos muy pronto.' : '¡Gracias! Lo revisaremos.' };
}

function panelComentarios(k) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const filas = leerTabla_(hojaComent_()).filter(r => r['ID']).reverse();
  const val = filas.filter(r => r['Tipo'] === 'Valoración' && Number(r['Estrellas']) > 0);
  const media = val.length ? Math.round(val.reduce((a, r) => a + Number(r['Estrellas']), 0) / val.length * 10) / 10 : 0;
  return { ok: true, media: media, valoraciones: val.length, nuevos: filas.filter(r => r['Estado'] === 'Nuevo').length,
    lista: filas.slice(0, 150).map(r => ({ id: r['ID'], fecha: fmtFecha_(r['Fecha'], true), tipo: r['Tipo'], estrellas: Number(r['Estrellas']) || 0,
      texto: r['Texto'], nombre: r['Nombre'], contacto: r['Contacto'], rfc: r['RFC'], estado: r['Estado'], issue: r['Issue'] })) };
}

/** Validar (→ prepara el issue de GitHub), descartar o marcar hecho. Devuelve el enlace para crear el issue. */
function panelComentarioEstado(k, id, estado) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  if (['Validado', 'Descartado', 'Hecho', 'Nuevo'].indexOf(estado) < 0) return { ok: false, error: 'Estado no válido.' };
  const sh = hojaComent_();
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) !== String(id)) continue;
    const r = objeto_(COL_COMENT, v[i]);
    const cambios = { 'Estado': estado };
    let url = '';
    if (estado === 'Validado') {
      const titulo = '[' + r['Tipo'] + '] ' + String(r['Texto']).replace(/\s+/g, ' ').slice(0, 70);
      const cuerpo = '**' + r['Tipo'] + '** recibido desde la app (' + r['ID'] + ', ' + fmtFecha_(r['Fecha'], true) + ')\n\n' + r['Texto'] +
        (r['Estrellas'] ? '\n\nValoración: ' + r['Estrellas'] + '/5' : '');
      const etiqueta = r['Tipo'] === 'Error' ? 'bug' : 'enhancement';
      url = 'https://github.com/' + REPO_GITHUB_ + '/issues/new?title=' + encodeURIComponent(titulo) + '&body=' + encodeURIComponent(cuerpo) + '&labels=' + etiqueta;
      cambios['Issue'] = url;
    }
    actualizarFila_(sh, COL_COMENT, i + 1, cambios);
    return { ok: true, url: url };
  }
  return { ok: false, error: 'No encuentro el comentario ' + id + '.' };
}

/* ============================ CAJA: PAGOS Y PEDIDOS ============================ */

const COL_PAGOPROV = ['ID', 'Fecha solicitud', 'Proveedor', 'Correo proveedor', 'Concepto', 'Importe', 'Fecha límite', 'Adjunto',
  'Solicitado por', 'Estado', 'Fecha pago', 'Comprobante', 'Correo al proveedor'];
const COL_PEDIDOS = ['ID', 'Fecha', 'Proveedor', 'Producto', 'Cantidad', 'Notas', 'Solicitado por', 'Estado', 'Fecha pedido'];
function hojaPagoProv_() { return hojaCon_('Pagos proveedores', COL_PAGOPROV); }
function hojaPedidos_() { return hojaCon_('Pedidos', COL_PEDIDOS); }

function claveCaja_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('CLAVE_CAJA');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); props.setProperty('CLAVE_CAJA', k); }
  return k;
}
function esCaja_(k) { return !!k && String(k) === claveCaja_(); }
function esCajaODuena_(k) { return esCaja_(k) || esDuena_(k); }

/** Enlace para instalar la app con la pestaña de Caja en el móvil de la cajera (abrir una sola vez). */
function enlaceAppCaja() {
  const u = String(cfg_().URL_APP || 'https://brenndhl.github.io/EL-PERMACHAN/') + '#k=' + claveCaja_() + '&r=caja';
  Logger.log('App de la cajera (abrir una vez en su móvil): ' + u);
  return u;
}
function cambiarClaveCaja() { PropertiesService.getScriptProperties().deleteProperty('CLAVE_CAJA'); return enlaceAppCaja(); }

function guardarAdjuntoCaja_(archivo, nombre) {
  if (!archivo || !archivo.base64) return '';
  return guardarArchivo_(archivo, nombre);
}

/** Datos de Caja: pagos y pedidos (pendientes primero). */
function cajaDatos(k) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const pagos = leerTabla_(hojaPagoProv_()).filter(r => r['ID']).reverse().slice(0, 150).map(r => ({
    id: r['ID'], fecha: fmtFecha_(r['Fecha solicitud'], true), proveedor: r['Proveedor'], correo: r['Correo proveedor'], concepto: r['Concepto'],
    importe: Number(r['Importe']) || 0, limite: fmtFecha_(r['Fecha límite']), adjunto: !!r['Adjunto'], estado: r['Estado'],
    pagado: fmtFecha_(r['Fecha pago'], true), comprobante: !!r['Comprobante'], aviso: r['Correo al proveedor'] }));
  const pedidos = leerTabla_(hojaPedidos_()).filter(r => r['ID']).reverse().slice(0, 300).map(r => ({
    id: r['ID'], fecha: fmtFecha_(r['Fecha'], true), proveedor: r['Proveedor'], producto: r['Producto'], cantidad: r['Cantidad'],
    notas: r['Notas'], estado: r['Estado'], pedido: fmtFecha_(r['Fecha pedido'], true) }));
  const proveedores = {};
  leerTabla_(hojaPagoProv_()).forEach(r => { if (r['Proveedor']) proveedores[r['Proveedor']] = r['Correo proveedor'] || proveedores[r['Proveedor']] || ''; });
  leerTabla_(hojaPedidos_()).forEach(r => { if (r['Proveedor'] && proveedores[r['Proveedor']] === undefined) proveedores[r['Proveedor']] = ''; });
  return { ok: true, duena: esDuena_(k), pagos: pagos, pedidos: pedidos, hoy: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    proveedores: Object.keys(proveedores).sort().map(p => [p, proveedores[p]]) };
}

/** La cajera pide un pago a proveedor. p = {proveedor, correo, concepto, importe, limite:'yyyy-mm-dd', adjunto:{base64,tipo}} */
function cajaNuevoPago(k, p) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  p = p || {};
  const prov = String(p.proveedor || '').trim().slice(0, 100), imp = Number(String(p.importe || '').replace(/[^\d.]/g, ''));
  if (prov.length < 2) return { ok: false, error: 'Escribe el proveedor.' };
  if (!(imp > 0)) return { ok: false, error: 'Escribe el importe.' };
  if (p.correo && !emailValido_(p.correo)) return { ok: false, error: 'El correo del proveedor no es válido.' };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const id = sigId_('SIG_PAGOPROV', 'P-');
    const lim = fechaDe_(p.limite);
    agregarFila_(hojaPagoProv_(), COL_PAGOPROV, { 'ID': id, 'Fecha solicitud': new Date(), 'Proveedor': prov,
      'Correo proveedor': String(p.correo || '').trim().toLowerCase(), 'Concepto': String(p.concepto || '').trim().slice(0, 300), 'Importe': imp,
      'Fecha límite': lim || '', 'Adjunto': guardarAdjuntoCaja_(p.adjunto, 'pago_' + id), 'Solicitado por': esCaja_(k) ? 'Caja' : 'Administración',
      'Estado': 'Pendiente', 'Fecha pago': '', 'Comprobante': '', 'Correo al proveedor': '' });
    enviarCorreo_(CORREO_CONTACTO_APPS_, 'Nuevo pago a proveedor pedido por caja: ' + prov + ' — $' + imp.toFixed(2),
      '<p><b>' + esc_(prov) + '</b> · $' + imp.toFixed(2) + (lim ? ' · antes del ' + fmtFecha_(lim) : '') + '</p><p>' + esc_(p.concepto || '') + '</p><p>Ábrelo en la app, en Administración → Proveedores.</p>');
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

/** La cajera pide productos. items = [{proveedor, producto, cantidad, notas}] */
function cajaNuevoPedido(k, items) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  items = (items || []).filter(it => it && String(it.producto || '').trim()).slice(0, 60);
  if (!items.length) return { ok: false, error: 'Añade al menos un producto.' };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = hojaPedidos_(), ahora = new Date();
    items.forEach(it => agregarFila_(sh, COL_PEDIDOS, { 'ID': sigId_('SIG_PEDIDO', 'D-'), 'Fecha': ahora, 'Proveedor': String(it.proveedor || '').trim().slice(0, 100),
      'Producto': String(it.producto).trim().slice(0, 150), 'Cantidad': String(it.cantidad || '').trim().slice(0, 40), 'Notas': String(it.notas || '').trim().slice(0, 200),
      'Solicitado por': esCaja_(k) ? 'Caja' : 'Administración', 'Estado': 'Pendiente', 'Fecha pedido': '' }));
    return { ok: true, n: items.length };
  } finally { lock.releaseLock(); }
}

/** Marcar pago hecho: sube comprobante y lo manda al proveedor. Solo la administradora. */
function panelPagoHecho(k, id, comprobante, avisar) {
  if (!esDuena_(k)) return { ok: false, error: 'Solo la administradora puede marcar pagos.' };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = hojaPagoProv_(), v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]) !== String(id)) continue;
      const r = objeto_(COL_PAGOPROV, v[i]);
      const url = comprobante && comprobante.base64 ? guardarArchivo_(comprobante, 'comprobante_' + id) : '';
      const cambios = { 'Estado': 'Pagado', 'Fecha pago': new Date() };
      if (url) cambios['Comprobante'] = url;
      let enviado = '';
      if (avisar !== false && emailValido_(r['Correo proveedor'])) {
        const b = url ? blobDeUrl_(url, 'Comprobante_pago_' + id) : null;
        const ok = enviarCorreo_(r['Correo proveedor'], 'Comprobante de pago — ' + cfg_().NOMBRE_NEGOCIO,
          '<p>Hola, les enviamos el comprobante del pago' + (r['Concepto'] ? ' de <b>' + esc_(r['Concepto']) + '</b>' : '') +
          ' por <b>$' + (Number(r['Importe']) || 0).toFixed(2) + '</b>.</p><p>Saludos,<br>' + esc_(cfg_().NOMBRE_NEGOCIO) + '</p>', b ? [b] : []);
        enviado = ok ? 'Enviado ' + fmtFecha_(new Date(), true) : 'No se pudo enviar';
        cambios['Correo al proveedor'] = enviado;
      }
      actualizarFila_(sh, COL_PAGOPROV, i + 1, cambios);
      return { ok: true, aviso: enviado };
    }
    return { ok: false, error: 'No encuentro el pago ' + id + '.' };
  } finally { lock.releaseLock(); }
}

/** Marcar pedidos como hechos (uno o varios). Solo la administradora. */
function panelPedidosHechos(k, ids) {
  if (!esDuena_(k)) return { ok: false, error: 'Solo la administradora puede marcar pedidos.' };
  ids = (ids || []).map(String);
  const sh = hojaPedidos_(), v = sh.getDataRange().getValues(), ahora = new Date();
  let n = 0;
  for (let i = 1; i < v.length; i++) if (ids.indexOf(String(v[i][0])) >= 0 && v[i][7] !== 'Pedido') {
    actualizarFila_(sh, COL_PEDIDOS, i + 1, { 'Estado': 'Pedido', 'Fecha pedido': ahora }); n++;
  }
  return { ok: true, n: n };
}

/** Ver adjunto o comprobante de un pago. */
function cajaArchivo(k, id, cual) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const r = leerTabla_(hojaPagoProv_()).filter(x => String(x['ID']) === String(id))[0];
  const b = r ? blobDeUrl_(cual === 'comprobante' ? r['Comprobante'] : r['Adjunto'], (cual === 'comprobante' ? 'Comprobante_' : 'Adjunto_') + id) : null;
  if (!b) return { ok: false, error: 'No hay archivo.' };
  return { ok: true, tipo: b.getContentType(), nombre: b.getName(), base64: Utilities.base64Encode(b.getBytes()) };
}

/** 9:00 diario: pagos pendientes. Lunes: pedidos pendientes de la semana. Correo a la administradora. */
function resumenCaja() {
  const pagos = leerTabla_(hojaPagoProv_()).filter(r => r['ID'] && r['Estado'] !== 'Pagado');
  const pedidos = leerTabla_(hojaPedidos_()).filter(r => r['ID'] && r['Estado'] !== 'Pedido');
  const lunes = new Date().getDay() === 1;
  if (!pagos.length && !(lunes && pedidos.length)) return 'Nada pendiente';
  let html = '';
  if (pagos.length) html += '<h3>Pagos a proveedores pendientes (' + pagos.length + ')</h3><table border="1" cellpadding="6" style="border-collapse:collapse">' +
    pagos.map(r => '<tr><td>' + esc_(r['ID']) + '</td><td>' + esc_(r['Proveedor']) + '</td><td>$' + (Number(r['Importe']) || 0).toFixed(2) + '</td><td>' +
      esc_(r['Concepto']) + '</td><td>' + (r['Fecha límite'] ? 'antes del ' + fmtFecha_(r['Fecha límite']) : '') + '</td></tr>').join('') + '</table>';
  if (lunes && pedidos.length) html += '<h3>Pedidos de la semana (' + pedidos.length + ')</h3><table border="1" cellpadding="6" style="border-collapse:collapse">' +
    pedidos.map(r => '<tr><td>' + esc_(r['Proveedor'] || '—') + '</td><td>' + esc_(r['Producto']) + '</td><td>' + esc_(r['Cantidad']) + '</td><td>' + esc_(r['Notas']) + '</td></tr>').join('') + '</table>';
  enviarCorreo_(CORREO_CONTACTO_APPS_, 'Caja: ' + pagos.length + ' pagos' + (lunes ? ' y ' + pedidos.length + ' pedidos' : '') + ' pendientes', html + '<p>Márcalos en la app, en Administración → Proveedores.</p>');
  return 'Resumen enviado';
}

/* ============================ COMPLEMENTOS DE PAGO ============================ */
// Cuando una factura "por definir" se marca pagada, se le pide a la contadora el complemento de pago.
// Ella lo sube igual que una factura (carpeta con el número de solicitud). El sistema lo reconoce
// (XML tipo "P"), lo manda al cliente y lo registra aquí.

const COL_COMPL = ['ID solicitud', 'RFC', 'Cliente', 'Folio ticket', 'UUID complemento', 'UUID factura', 'Monto', 'Fecha pago', 'PDF', 'XML', 'Subido', 'Enviado al cliente'];
function hojaCompl_() { return hojaCon_('Complementos', COL_COMPL); }
function complementosMap_() {
  const m = {};
  try { leerTabla_(hojaCompl_()).forEach(r => { if (r['ID solicitud']) m[String(r['ID solicitud'])] = r; }); } catch (e) { /* sin hoja */ }
  return m;
}

function procesarComplemento_(base, g, cfdi, procesadas, revisar) {
  if (Object.keys(complementosMap_()).some(k => String(complementosMap_()[k]['UUID complemento']).toUpperCase() === cfdi.uuid)) {
    return aRevisar_(g, revisar, base, 'Este complemento de pago (UUID ' + cfdi.uuid + ') ya se había subido.');
  }
  const facts = facturasMap_();
  let id = g.id || (g.carpeta ? idDeNombre_(g.carpeta.getName()) : idDeNombre_(base));
  if (!id) {
    const porUuid = Object.keys(facts).filter(k => cfdi.relacionados.indexOf(String(facts[k]['UUID']).toUpperCase()) >= 0);
    if (porUuid.length === 1) id = porUuid[0];
  }
  const f = id ? buscarSolicitud_(id) : null;
  if (!f) return aRevisar_(g, revisar, base, 'Es un complemento de pago pero no sé de qué solicitud: pon la carpeta con el número de solicitud.');
  const s = f.datos;
  if (cfdi.rfc && normalizarRFC(s['RFC']) !== cfdi.rfc) {
    return aRevisar_(g, revisar, base, 'El complemento está a nombre de ' + cfdi.rfc + ' pero la solicitud ' + id + ' es de ' + s['RFC'] + '.');
  }
  const destino = subcarpeta_(procesadas, normalizarRFC(s['RFC']) || 'SIN_RFC');
  if (g.carpeta) { g.carpeta.setName(id + ' — complemento de pago'); g.carpeta.moveTo(destino); }
  else { if (g.pdf) g.pdf.moveTo(destino); if (g.xml) g.xml.moveTo(destino); }
  let enviado = false;
  if (emailValido_(s['Email'])) {
    const conf = cfg_(), adj = [];
    if (g.pdf) adj.push(g.pdf.getBlob().setName('Complemento_pago_' + s['Folio ticket'] + '.pdf'));
    if (g.xml) adj.push(g.xml.getBlob().setName('Complemento_pago_' + s['Folio ticket'] + '.xml'));
    enviado = enviarCorreo_(s['Email'], 'Complemento de pago — ' + conf.NOMBRE_NEGOCIO + ' — ticket ' + s['Folio ticket'],
      '<p>Hola, te enviamos el <b>complemento de pago</b> de tu factura del ticket <b>' + esc_(s['Folio ticket']) + '</b>' +
      (facts[id] && facts[id]['Serie-Folio'] ? ' (factura ' + esc_(facts[id]['Serie-Folio']) + ')' : '') + '. Lo tienes adjunto (PDF y XML) y también en la app, en <b>Mis facturas</b>.</p>' +
      '<p>Gracias por tu pago.<br>' + esc_(conf.NOMBRE_NEGOCIO) + '</p>', adj);
  }
  agregarFila_(hojaCompl_(), COL_COMPL, { 'ID solicitud': id, 'RFC': s['RFC'], 'Cliente': s['Nombre o razón social'], 'Folio ticket': s['Folio ticket'],
    'UUID complemento': cfdi.uuid, 'UUID factura': cfdi.relacionados.join(', '), 'Monto': Number(cfdi.montoPago) || cfdi.montoPago,
    'Fecha pago': cfdi.fechaPago ? new Date(cfdi.fechaPago.slice(0, 19)) : '', 'PDF': g.pdf ? g.pdf.getUrl() : '', 'XML': g.xml ? g.xml.getUrl() : '',
    'Subido': new Date(), 'Enviado al cliente': enviado ? 'Sí — ' + s['Email'] : 'No se pudo enviar' });
  return base + ' → complemento de ' + id + (enviado ? ' enviado' : ' SIN correo');
}

/** Lunes: recuerda a la contadora los complementos que faltan (pagadas hace más de 3 días sin complemento). */
function recordarComplementos() {
  const compl = complementosMap_(), facts = facturasMap_();
  const limite = Date.now() - 3 * 86400000;
  const faltan = leerTabla_(hojaPagos_()).filter(p => p['ID'] && !compl[String(p['ID'])] && p['Marcado el'] instanceof Date && p['Marcado el'].getTime() < limite);
  if (!faltan.length) return 'Nada pendiente';
  const tz = Session.getScriptTimeZone();
  enviarCorreo_(correoContadora_(), 'Complementos de pago pendientes (' + faltan.length + ') — ' + cfg_().NOMBRE_NEGOCIO,
    '<p>Hola, faltan estos <b>complementos de pago</b>. Súbelos igual que las facturas: en la carpeta de siempre, dentro de una carpeta con el <b>número de solicitud</b>.</p>' +
    '<table border="1" cellpadding="6" style="border-collapse:collapse"><tr><th>Carpeta</th><th>Cliente</th><th>RFC</th><th>Factura</th><th>Pagada el</th></tr>' +
    faltan.map(p => '<tr><td><b>' + esc_(String(p['ID']).replace(/^S-0*/, '')) + '</b></td><td>' + esc_(p['Cliente']) + '</td><td>' + esc_(p['RFC']) + '</td><td>' +
      esc_(facts[String(p['ID'])] ? facts[String(p['ID'])]['Serie-Folio'] || facts[String(p['ID'])]['UUID'] : '—') + '</td><td>' +
      (p['Fecha pago'] instanceof Date ? Utilities.formatDate(p['Fecha pago'], tz, 'dd/MM/yyyy') : '') + '</td></tr>').join('') + '</table>' +
    '<p>' + esc_(cfg_().NOMBRE_NEGOCIO) + '</p>');
  return 'Recordatorio enviado: ' + faltan.length;
}

/* ============================ INSTALACIÓN ============================ */

function instalarExtras() {
  hojaComent_(); hojaPagoProv_(); hojaPedidos_(); hojaCompl_();
  ['Correos extra', 'Correos quitados', 'Fecha constancia', 'Aviso CSF'].forEach(colCliente_);
  ScriptApp.getProjectTriggers().forEach(t => { if (['resumenCaja', 'recordatorioCSF', 'recordarComplementos'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('resumenCaja').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('recordatorioCSF').timeBased().onMonthDay(5).atHour(11).create();
  ScriptApp.newTrigger('recordarComplementos').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(10).create();
  const txt = 'Extras instalados. Enlace de la cajera: ' + enlaceAppCaja();
  Logger.log(txt);
  return txt;
}
