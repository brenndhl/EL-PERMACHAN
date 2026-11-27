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
const COL_PEDIDOS = ['ID', 'Fecha', 'Proveedor', 'Producto', 'Cantidad', 'Notas', 'Solicitado por', 'Estado', 'Fecha pedido', 'Fecha recibido', 'Recibido por'];
function hojaPagoProv_() { return hojaCon_('Pagos proveedores', COL_PAGOPROV); }
function hojaPedidos_() { return asegurarCols_(hojaCon_('Pedidos', COL_PEDIDOS), COL_PEDIDOS); }
/** Añade a la cabecera las columnas nuevas que falten (hojas creadas con una versión anterior). */
function asegurarCols_(sh, cols) {
  const n = sh.getLastColumn();
  if (n < cols.length) sh.getRange(1, n + 1, 1, cols.length - n).setValues([cols.slice(n)]).setFontWeight('bold').setBackground('#1f4e5f').setFontColor('#ffffff');
  return sh;
}

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
    id: r['ID'], por: r['Solicitado por'], fecha: fmtFecha_(r['Fecha solicitud'], true), proveedor: r['Proveedor'], correo: r['Correo proveedor'], concepto: r['Concepto'],
    importe: Number(r['Importe']) || 0, limite: fmtFecha_(r['Fecha límite']), adjunto: !!r['Adjunto'], estado: r['Estado'],
    pagado: fmtFecha_(r['Fecha pago'], true), comprobante: !!r['Comprobante'], aviso: r['Correo al proveedor'] }));
  const pedidos = leerTabla_(hojaPedidos_()).filter(r => r['ID']).reverse().slice(0, 300).map(r => ({
    id: r['ID'], por: r['Solicitado por'], fecha: fmtFecha_(r['Fecha'], true), proveedor: r['Proveedor'], producto: r['Producto'], cantidad: r['Cantidad'],
    notas: r['Notas'], estado: r['Estado'], pedido: fmtFecha_(r['Fecha pedido'], true), recibido: fmtFecha_(r['Fecha recibido'], true), recibidoPor: r['Recibido por'] || '' }));
  const proveedores = {};
  leerTabla_(hojaPagoProv_()).forEach(r => { if (r['Proveedor']) proveedores[r['Proveedor']] = r['Correo proveedor'] || proveedores[r['Proveedor']] || ''; });
  leerTabla_(hojaPedidos_()).forEach(r => { if (r['Proveedor'] && proveedores[r['Proveedor']] === undefined) proveedores[r['Proveedor']] = ''; });
  let cortes = { lista: [], mes: null };
  try { cortes = cortesDatos_(); } catch (err) { Logger.log(err); }
  const corteHoy = (cortes.lista || []).some(c => c.fecha === ymd_(new Date()));
  return { ok: true, duena: esDuena_(k), pagos: pagos, pedidos: pedidos, cortes: cortes, corteHoy: corteHoy, horaCorte: HORA_AVISO_CORTE_, hoy: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    proveedores: Object.keys(proveedores).sort().map(p => [p, proveedores[p]]) };
}

/** La cajera pide un pago a proveedor. p = {proveedor, correo, concepto, importe, limite:'yyyy-mm-dd', adjunto:{base64,tipo}} */
function cajaNuevoPago(k, p) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  p = p || {};
  const quien = quienCaja_(k, p.cajera);
  if (!quien) return { ok: false, cajera: false, error: 'Escribe tu código de cajera.' };
  const prov = String(p.proveedor || '').trim().slice(0, 100), imp = Number(String(p.importe || '').replace(/[^\d.]/g, ''));
  if (prov.length < 2) return { ok: false, error: 'Escribe el proveedor.' };
  if (!(imp > 0)) return { ok: false, error: 'Escribe el importe.' };
  if (p.correo && !emailValido_(p.correo)) return { ok: false, error: 'El correo del proveedor no es válido.' };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const id = sigId_('SIG_PAGOPROV', 'P-');
    const lim = fechaDe_(p.limite);
    const yaPagado = p.estado === 'Pagado' && esDuena_(k);
    agregarFila_(hojaPagoProv_(), COL_PAGOPROV, { 'ID': id, 'Fecha solicitud': new Date(), 'Proveedor': prov,
      'Correo proveedor': String(p.correo || '').trim().toLowerCase(), 'Concepto': String(p.concepto || '').trim().slice(0, 300), 'Importe': imp,
      'Fecha límite': lim || '', 'Adjunto': guardarAdjuntoCaja_(p.adjunto, 'pago_' + id), 'Solicitado por': quien,
      'Estado': yaPagado ? 'Pagado' : 'Pendiente', 'Fecha pago': yaPagado ? new Date() : '', 'Comprobante': '', 'Correo al proveedor': '' });
    actividad_(quien, yaPagado ? 'Pago registrado (ya pagado)' : 'Pago pedido', id + ' · ' + prov + ' · $' + imp.toFixed(2));
    if (!esCaja_(k) || yaPagado) return { ok: true, id: id };
    enviarCorreo_(CORREO_CONTACTO_APPS_, 'Nuevo pago a proveedor pedido por ' + quien + ': ' + prov + ' — $' + imp.toFixed(2),
      '<p><b>' + esc_(prov) + '</b> · $' + imp.toFixed(2) + (lim ? ' · antes del ' + fmtFecha_(lim) : '') + '</p><p>' + esc_(p.concepto || '') + '</p><p>Ábrelo en la app, en Administración → Proveedores.</p>');
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

/** La cajera pide productos. items = [{proveedor, producto, cantidad, notas}] */
function cajaNuevoPedido(k, items, cajera, estado) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const quien = quienCaja_(k, cajera);
  if (!quien) return { ok: false, cajera: false, error: 'Escribe tu código de cajera.' };
  items = (items || []).filter(it => it && String(it.producto || '').trim()).slice(0, 60);
  if (!items.length) return { ok: false, error: 'Añade al menos un producto.' };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = hojaPedidos_(), ahora = new Date();
    const est = ['Pedido', 'Recibido'].indexOf(estado) >= 0 && esDuena_(k) ? estado : 'Pendiente';
    items.forEach(it => agregarFila_(sh, COL_PEDIDOS, { 'ID': sigId_('SIG_PEDIDO', 'D-'), 'Fecha': ahora, 'Proveedor': String(it.proveedor || '').trim().slice(0, 100),
      'Producto': String(it.producto).trim().slice(0, 150), 'Cantidad': String(it.cantidad || '').trim().slice(0, 40), 'Notas': String(it.notas || '').trim().slice(0, 200),
      'Solicitado por': quien, 'Estado': est, 'Fecha pedido': est !== 'Pendiente' ? ahora : '', 'Fecha recibido': est === 'Recibido' ? ahora : '', 'Recibido por': est === 'Recibido' ? quien : '' }));
    actividad_(quien, 'Pedido', items.length + ' productos: ' + items.map(it => it.producto).join(', ').slice(0, 200));
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
  for (let i = 1; i < v.length; i++) if (ids.indexOf(String(v[i][0])) >= 0 && v[i][7] === 'Pendiente') {
    actualizarFila_(sh, COL_PEDIDOS, i + 1, { 'Estado': 'Pedido', 'Fecha pedido': ahora }); n++;
  }
  return { ok: true, n: n };
}

/** La cajera (o la administradora) confirma que el pedido llegó. */
function cajaRecibido(k, ids, cajera) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const quien = quienCaja_(k, cajera);
  if (!quien) return { ok: false, cajera: false, error: 'Escribe tu código de cajera.' };
  ids = (ids || []).map(String);
  const sh = hojaPedidos_(), v = sh.getDataRange().getValues(), ahora = new Date();
  const nombres = [];
  for (let i = 1; i < v.length; i++) if (ids.indexOf(String(v[i][0])) >= 0 && v[i][7] !== 'Recibido') {
    const c = { 'Estado': 'Recibido', 'Fecha recibido': ahora, 'Recibido por': quien };
    if (!v[i][8]) c['Fecha pedido'] = ahora;
    actualizarFila_(sh, COL_PEDIDOS, i + 1, c); nombres.push(v[i][3]);
  }
  if (nombres.length) actividad_(quien, 'Pedido recibido', nombres.join(', ').slice(0, 250));
  return { ok: true, n: nombres.length };
}

/** La administradora marca un pago como pagado sin comprobante ni correo (p. ej. pagos antiguos). */
function panelPagoEstado(k, id, estado) {
  if (!esDuena_(k)) return { ok: false, error: 'Solo la administradora.' };
  const sh = hojaPagoProv_(), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) {
    actualizarFila_(sh, COL_PAGOPROV, i + 1, estado === 'Pendiente' ? { 'Estado': 'Pendiente', 'Fecha pago': '' } : { 'Estado': 'Pagado', 'Fecha pago': new Date() });
    return { ok: true };
  }
  return { ok: false, error: 'No encuentro el pago.' };
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
  const desde = Number(PropertiesService.getScriptProperties().getProperty('COMPL_DESDE') || 0);   // los pagos marcados antes no piden complemento
  const faltan = leerTabla_(hojaPagos_()).filter(p => p['ID'] && !compl[String(p['ID'])] && p['Marcado el'] instanceof Date && p['Marcado el'].getTime() < limite && p['Marcado el'].getTime() > desde);
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

/* ============================ CORTE DE CAJA DIARIO ============================ */
// La cajera lo llena al final del día con el cierre de la terminal Getnet (débito y crédito por separado).
// Getnet cobra una comisión (tasa de descuento) + IVA sobre cada venta con tarjeta: a la cuenta llega el NETO.
// Tasas por defecto: promedio de Getnet para restaurantes publicado por Banxico (débito 1.79 %, crédito 2.26 %).
// Cámbialas por las de vuestro contrato en la hoja Config: COMISION_DEBITO, COMISION_CREDITO (en %), IVA_COMISION (16).
// Opcional: SALDO_TARJETA_INICIAL = dinero que había en la cuenta de la tarjeta antes del primer corte.

const COL_CORTES = ['Fecha', 'Tarjeta débito', 'Tarjeta crédito', 'Tarjeta', 'Comisión Getnet', 'Neto tarjeta', 'Fondo inicial', 'Ventas totales',
  'Transferencias', 'Gastos en efectivo', 'Detalle gastos', 'Efectivo esperado', 'Efectivo contado', 'Diferencia', 'Notas', 'Foto', 'Registrado por', 'Registrado el'];
function hojaCortes_() { return hojaCon_('Cortes de caja', COL_CORTES); }
function num_(v) { const n = Number(String(v == null ? '' : v).replace(/[^\d.\-]/g, '')); return isFinite(n) ? Math.round(n * 100) / 100 : 0; }
function r2_(n) { return Math.round(n * 100) / 100; }
function ymd_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

function tasasGetnet_() {
  const c = cfg_();
  const t = (k, d) => { const v = Number(String(c[k] == null ? '' : c[k]).replace(',', '.')); return isFinite(v) && v > 0 ? v : d; };
  return { debito: t('COMISION_DEBITO', 1.79), credito: t('COMISION_CREDITO', 2.26), iva: t('IVA_COMISION', 16) };
}
function comisionGetnet_(deb, cred) {
  const t = tasasGetnet_();
  return r2_((deb * t.debito / 100 + cred * t.credito / 100) * (1 + t.iva / 100));
}

function corteFila_(r) {
  const f = r['Fecha'] instanceof Date ? ymd_(r['Fecha']) : String(r['Fecha'] || '');
  return { fecha: f, debito: num_(r['Tarjeta débito']), credito: num_(r['Tarjeta crédito']), tarjeta: num_(r['Tarjeta']), comision: num_(r['Comisión Getnet']),
    neto: num_(r['Neto tarjeta']), fondo: num_(r['Fondo inicial']), ventas: num_(r['Ventas totales']), transf: num_(r['Transferencias']),
    gastos: num_(r['Gastos en efectivo']), detalle: r['Detalle gastos'] || '', esperado: num_(r['Efectivo esperado']),
    contado: r['Efectivo contado'] === '' ? null : num_(r['Efectivo contado']), diferencia: num_(r['Diferencia']), notas: r['Notas'] || '', foto: !!r['Foto'], por: r['Registrado por'] || '' };
}

/** Guarda (o corrige) el corte de un día.
 *  c = {fecha:'yyyy-mm-dd', debito, credito, fondo, ventas, transf, gastos, detalle, contado, notas, foto:{base64,tipo}} */
function cajaCorte(k, c) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  c = c || {};
  const quien = quienCaja_(k, c.cajera);
  if (!quien) return { ok: false, cajera: false, error: 'Escribe tu código de cajera.' };
  const fecha = fechaDe_(c.fecha) || new Date();
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  if (fecha > hoy) return { ok: false, error: 'La fecha del corte no puede ser futura.' };
  const v = { deb: num_(c.debito), cred: num_(c.credito), fondo: num_(c.fondo), ventas: num_(c.ventas), transf: num_(c.transf), gastos: num_(c.gastos), contado: num_(c.contado) };
  const tarjeta = r2_(v.deb + v.cred);
  if (!(tarjeta > 0) && !(v.ventas > 0)) return { ok: false, error: 'Escribe lo que entró con tarjeta (débito y crédito) o las ventas del día.' };
  if (v.ventas > 0 && tarjeta + v.transf > v.ventas) return { ok: false, error: 'Tarjeta + transferencias no pueden ser más que las ventas totales.' };
  const comision = comisionGetnet_(v.deb, v.cred), neto = r2_(tarjeta - comision);
  const hayEfectivo = v.ventas > 0 && String(c.contado == null ? '' : c.contado).trim() !== '';
  const esperado = v.ventas > 0 ? r2_(v.fondo + v.ventas - tarjeta - v.transf - v.gastos) : '';
  const dif = hayEfectivo ? r2_(v.contado - esperado) : '';
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = hojaCortes_(), clave = ymd_(fecha);
    const fila = { 'Fecha': new Date(clave + 'T12:00:00'), 'Tarjeta débito': v.deb, 'Tarjeta crédito': v.cred, 'Tarjeta': tarjeta, 'Comisión Getnet': comision,
      'Neto tarjeta': neto, 'Fondo inicial': v.fondo, 'Ventas totales': v.ventas || '', 'Transferencias': v.transf, 'Gastos en efectivo': v.gastos,
      'Detalle gastos': String(c.detalle || '').trim().slice(0, 300), 'Efectivo esperado': esperado, 'Efectivo contado': hayEfectivo ? v.contado : '',
      'Diferencia': dif, 'Notas': String(c.notas || '').trim().slice(0, 300), 'Registrado por': quien, 'Registrado el': new Date() };
    if (c.foto && c.foto.base64) fila['Foto'] = guardarArchivo_(c.foto, 'corte_' + clave);
    const vals = sh.getDataRange().getValues();
    let n = -1;
    for (let i = 1; i < vals.length; i++) if (vals[i][0] instanceof Date && ymd_(vals[i][0]) === clave) n = i + 1;
    if (n > 0) actualizarFila_(sh, COL_CORTES, n, fila); else agregarFila_(sh, COL_CORTES, fila);
    const saldo = saldoTarjeta_();
    const m = x => '$' + Number(x).toFixed(2);
    const t = tasasGetnet_();
    const filasCorreo = [['Débito', m(v.deb)], ['Crédito', m(v.cred)], ['Total tarjeta', '<b>' + m(tarjeta) + '</b>'],
      ['− Comisión Getnet + IVA', m(comision) + ' <small>(débito ' + t.debito + ' %, crédito ' + t.credito + ' %, IVA ' + t.iva + ' %)</small>'],
      ['= Llega a la cuenta', '<b>' + m(neto) + '</b>'], ['Dinero de tarjeta acumulado', '<b>' + m(saldo.total) + '</b>']];
    if (v.ventas > 0) filasCorreo.push(['Ventas totales', m(v.ventas)], ['Efectivo esperado', m(esperado)]);
    if (hayEfectivo) filasCorreo.push(['Efectivo contado', m(v.contado)], ['Diferencia', '<b style="color:' + (Math.abs(dif) < 0.5 ? '#1e6b3a' : '#b3261e') + '">' + (dif > 0 ? '+' : '') + m(dif) + '</b>']);
    if (fila['Notas']) filasCorreo.push(['Notas', esc_(fila['Notas'])]);
    enviarCorreo_(CORREO_CONTACTO_APPS_, 'Corte ' + Utilities.formatDate(fila['Fecha'], Session.getScriptTimeZone(), 'dd/MM/yyyy') + ' — tarjeta ' + m(tarjeta) + ', llegan ' + m(neto) +
      (hayEfectivo && Math.abs(dif) >= 0.5 ? ' — ' + (dif > 0 ? 'sobran ' : 'faltan ') + m(Math.abs(dif)) + ' en efectivo' : ''),
      '<table border="1" cellpadding="6" style="border-collapse:collapse">' + filasCorreo.map(r => '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td></tr>').join('') + '</table>' +
      (n > 0 ? '<p>(Es una corrección de un corte que ya estaba guardado.)</p>' : ''));
    actividad_(quien, n > 0 ? 'Corte corregido' : 'Corte', clave + ' · tarjeta $' + tarjeta.toFixed(2));
    return { ok: true, tarjeta: tarjeta, comision: comision, neto: neto, saldo: saldo.total, esperado: esperado, diferencia: dif, corregido: n > 0 };
  } finally { lock.releaseLock(); }
}

/** Dinero de tarjeta acumulado = saldo inicial (Config) + netos de todos los cortes. */
function saldoTarjeta_() {
  const ini = num_(cfg_().SALDO_TARJETA_INICIAL);
  const netos = leerTabla_(hojaCortes_()).reduce((a, r) => a + num_(r['Neto tarjeta']), 0);
  return { inicial: ini, total: r2_(ini + netos) };
}

/** Cortes recientes, resumen del mes y dinero de tarjeta acumulado (para Caja y para el panel). */
function cortesDatos_() {
  const filas = leerTabla_(hojaCortes_()).filter(r => r['Fecha'] instanceof Date).sort((a, b) => b['Fecha'].getTime() - a['Fecha'].getTime());
  const mes = ymd_(new Date()).slice(0, 7);
  const delMes = filas.filter(r => ymd_(r['Fecha']).slice(0, 7) === mes).map(corteFila_);
  const sum = key => r2_(delMes.reduce((a, x) => a + (Number(x[key]) || 0), 0));
  return { lista: filas.slice(0, 62).map(corteFila_), tasas: tasasGetnet_(), saldo: saldoTarjeta_(),
    mes: { dias: delMes.length, tarjeta: sum('tarjeta'), comision: sum('comision'), neto: sum('neto'), ventas: sum('ventas'), gastos: sum('gastos'), diferencia: sum('diferencia') } };
}

/** Vista previa del cálculo (para la pantalla, sin guardar). */
function cajaTasas(k) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  return { ok: true, tasas: tasasGetnet_() };
}

function cajaCorteFoto(k, fecha) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const r = leerTabla_(hojaCortes_()).filter(x => x['Fecha'] instanceof Date && ymd_(x['Fecha']) === String(fecha))[0];
  const b = r ? blobDeUrl_(r['Foto'], 'Corte_' + fecha) : null;
  if (!b) return { ok: false, error: 'No hay foto.' };
  return { ok: true, tipo: b.getContentType(), nombre: b.getName(), base64: Utilities.base64Encode(b.getBytes()) };
}

const HORA_AVISO_CORTE_ = 17;   // 5 de la tarde (hora de México): recordatorio del corte (se cierra a las 5:30)

/** La administradora escribe cuánto hay HOY en la cuenta de la tarjeta; desde ahí se va sumando cada corte. */
function panelSaldoTarjeta(k, saldoHoy) {
  if (!esDuena_(k)) return { ok: false, error: 'Solo la administradora.' };
  const s = num_(saldoHoy);
  if (String(saldoHoy).trim() === '') return { ok: false, error: 'Escribe el saldo.' };
  const netos = leerTabla_(hojaCortes_()).reduce((a, r) => a + num_(r['Neto tarjeta']), 0);
  setConfig_('SALDO_TARJETA_INICIAL', r2_(s - netos));
  return { ok: true, saldo: saldoTarjeta_().total };
}

/** Comisiones de vuestro contrato Getnet (en %). */
function panelTasas(k, debito, credito) {
  if (!esDuena_(k)) return { ok: false, error: 'Solo la administradora.' };
  const d = num_(debito), c = num_(credito);
  if (!(d > 0 && d < 10 && c > 0 && c < 10)) return { ok: false, error: 'Escribe los porcentajes, por ejemplo 1.65 y 2.20.' };
  setConfig_('COMISION_DEBITO', d); setConfig_('COMISION_CREDITO', c);
  return { ok: true, tasas: tasasGetnet_() };
}

/** 17:00 (México): si todavía no hay corte de hoy, avisa a las cajeras por correo. En la app también sale el aviso. */
function recordatorioCorte() {
  const hoy = ymd_(new Date());
  if (leerTabla_(hojaCortes_()).some(r => r['Fecha'] instanceof Date && ymd_(r['Fecha']) === hoy)) return 'Ya hay corte';
  const correos = leerTabla_(hojaCajeras_()).filter(r => r['Activa'] !== 'No' && emailValido_(r['Correo'])).map(r => r['Correo']);
  if (!correos.length) return 'Sin correos de cajeras';
  const app = String(cfg_().URL_APP || 'https://brenndhl.github.io/EL-PERMACHAN/');
  enviarCorreo_(correos.join(','), '⏰ Hora del corte de caja — ' + cfg_().NOMBRE_NEGOCIO,
    '<p style="font-size:18px">Hola, ya casi cerramos. Antes de irte haz el <b>corte del día</b>:</p>' +
    '<ol style="font-size:17px"><li>Saca el cierre de la terminal Getnet.</li><li>Abre la app → <b>Caja</b> → <b>🧮 Corte</b>.</li>' +
    '<li>Escribe lo de <b>débito</b> y lo de <b>crédito</b> y pulsa <b>Guardar corte</b>.</li></ol>' +
    '<p style="text-align:center;margin:22px 0"><a href="' + esc_(app) + '" style="background:#141110;color:#f3ead6;text-decoration:none;font-weight:bold;font-size:18px;padding:14px 26px;border-radius:10px">Abrir la app</a></p>');
  return 'Aviso enviado';
}

/* ============================ CAJA: CONSULTAR RFC ============================ */

/** La cajera consulta si un RFC es de persona física o moral y si ya es cliente. */
function cajaConsultarRFC(k, texto) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const rfc = normalizarRFC(texto);
  if (!rfc) return { ok: false, error: 'Escribe un RFC.' };
  const valido = rfcValido(rfc);
  const tipo = rfc.length === 12 ? 'Moral' : rfc.length === 13 ? 'Física' : '';
  const c = valido ? clienteCompleto_(rfc) : null;
  const d = c ? c.datos : null;
  const reg = d ? REGIMENES.filter(x => x[0] === String(d['Régimen fiscal']))[0] : null;
  return { ok: true, rfc: rfc, valido: valido, tipo: tipo, caracteres: rfc.length, generico: rfc === 'XAXX010101000' || rfc === 'XEXX010101000',
    cliente: d ? { nombre: d['Nombre o razón social'], cp: String(d['CP fiscal'] || ''), regimen: reg ? reg[0] + ' — ' + reg[1] : String(d['Régimen fiscal'] || '—'),
      constancia: !!d['Constancia'], csfVencida: csfVencida_(d) } : null };
}

/** Manda a elpermachan@gmail.com el enlace de la app con TU acceso de administradora (para tu móvil). */
function enviarEnlaceDuena() {
  const u = String(cfg_().URL_APP || 'https://brenndhl.github.io/EL-PERMACHAN/') + '#k=' + claveDuena_();
  enviarCorreo_(CORREO_CONTACTO_APPS_, 'Tu acceso de administradora a la app — ' + cfg_().NOMBRE_NEGOCIO,
    '<p>Abre este enlace <b>una vez en tu móvil</b> (no lo reenvíes a nadie). En Administración verás también <b>Pagos, pedidos y cortes</b>:</p>' +
    '<p style="font-size:16px"><a href="' + esc_(u) + '">' + esc_(u) + '</a></p>');
  return 'Enviado a ' + CORREO_CONTACTO_APPS_;
}

/** Manda a elpermachan@gmail.com el enlace de la app de la cajera (para reenviárselo por WhatsApp). */
function enviarEnlaceCaja() {
  const u = enlaceAppCaja();
  enviarCorreo_(CORREO_CONTACTO_APPS_, 'Enlace de la app para la cajera — ' + cfg_().NOMBRE_NEGOCIO,
    '<p>Este es el enlace de la cajera. Mándaselo por WhatsApp y que lo abra <b>una sola vez en su móvil</b> (luego que instale la app desde ahí):</p>' +
    '<p style="font-size:16px"><a href="' + esc_(u) + '">' + esc_(u) + '</a></p>' +
    '<p>En su móvil la última pestaña se llamará <b>🧾 Caja</b>: pagos a proveedores, pedidos, corte del día y consultar RFC.</p>' +
    '<p>Si algún día se reenvía por error, ejecuta <i>cambiarClaveCaja</i> y el enlace viejo deja de funcionar.</p>');
  return 'Enlace enviado a ' + CORREO_CONTACTO_APPS_;
}

/* ============================ CAJERAS ============================ */
// Cada cajera tiene su código de 4 cifras. Lo escribe una vez en su móvil y así queda registrado
// quién pide cada pago, cada pedido y cada corte. La administradora las da de alta en el panel.

const COL_CAJERAS = ['Código', 'Nombre', 'Activa', 'Alta', 'Último uso', 'Correo'];
const COL_ACTIV = ['Fecha', 'Cajera', 'Acción', 'Detalle'];
function hojaCajeras_() { return asegurarCols_(hojaCon_('Cajeras', COL_CAJERAS), COL_CAJERAS); }
function hojaActiv_() { return hojaCon_('Actividad caja', COL_ACTIV); }

function cajeraPorCodigo_(codigo) {
  codigo = String(codigo || '').replace(/\D/g, '');
  if (!codigo) return null;
  const sh = hojaCajeras_(), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).replace(/\D/g, '') === codigo && v[i][2] !== 'No') return { fila: i + 1, nombre: String(v[i][1]) };
  return null;
}

/** Nombre de quien hace la acción: la cajera por su código, o "Administración" si es la dueña. */
function quienCaja_(k, codigo) {
  const c = cajeraPorCodigo_(codigo);
  if (c) { try { hojaCajeras_().getRange(c.fila, 5).setValue(new Date()); } catch (e) { /* sin fecha */ } return c.nombre; }
  return esDuena_(k) ? 'Administración' : '';
}

function actividad_(quien, accion, detalle) {
  try { agregarFila_(hojaActiv_(), COL_ACTIV, { 'Fecha': new Date(), 'Cajera': quien, 'Acción': accion, 'Detalle': String(detalle || '').slice(0, 300) }); } catch (e) { Logger.log(e); }
}

/** La cajera escribe su código en el móvil. */
function cajaEntrar(k, codigo) {
  if (!esCajaODuena_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const cache = CacheService.getScriptCache(), n = Number(cache.get('cajint') || 0);
  if (n > 20) return { ok: false, error: 'Demasiados intentos. Espera 15 minutos.' };
  const c = cajeraPorCodigo_(codigo);
  if (!c) { cache.put('cajint', String(n + 1), 900); return { ok: false, error: 'Código incorrecto. Pídeselo a Brenda.' }; }
  hojaCajeras_().getRange(c.fila, 5).setValue(new Date());
  actividad_(c.nombre, 'Entró en la app', '');
  return { ok: true, nombre: c.nombre };
}

/** Panel: cajeras, qué ha hecho cada una (últimos 30 días) y actividad reciente. */
function panelCajeras(k) {
  if (!esDuena_(k)) return { ok: false, error: 'Solo la administradora.' };
  const hace30 = Date.now() - 30 * 86400000;
  const act = leerTabla_(hojaActiv_()).filter(r => r['Fecha'] instanceof Date);
  const cajeras = leerTabla_(hojaCajeras_()).filter(r => r['Nombre']).map(r => {
    const suyas = act.filter(a => a['Cajera'] === r['Nombre'] && a['Fecha'].getTime() > hace30);
    const cuenta = t => suyas.filter(a => String(a['Acción']).indexOf(t) === 0).length;
    return { codigo: String(r['Código']), nombre: r['Nombre'], correo: r['Correo'] || '', activa: r['Activa'] !== 'No', alta: fmtFecha_(r['Alta']), ultimo: fmtFecha_(r['Último uso'], true),
      pagos: cuenta('Pago'), pedidos: cuenta('Pedido') - cuenta('Pedido recibido'), recibidos: cuenta('Pedido recibido'), cortes: cuenta('Corte'), entradas: cuenta('Entró') };
  });
  return { ok: true, cajeras: cajeras, actividad: act.slice(-80).reverse().map(a => ({ fecha: fmtFecha_(a['Fecha'], true), cajera: a['Cajera'], accion: a['Acción'], detalle: a['Detalle'] })) };
}

/** Alta o edición de una cajera. c = {codigo? (si se edita), nombre, activa} */
function panelCajeraGuardar(k, c) {
  if (!esDuena_(k)) return { ok: false, error: 'Solo la administradora.' };
  c = c || {};
  const nombre = String(c.nombre || '').trim().slice(0, 60);
  if (nombre.length < 2) return { ok: false, error: 'Escribe el nombre de la cajera.' };
  const correo = String(c.correo || '').trim().toLowerCase();
  if (correo && !emailValido_(correo)) return { ok: false, error: 'El correo no es válido.' };
  const sh = hojaCajeras_(), v = sh.getDataRange().getValues();
  if (c.codigo) {
    for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(c.codigo)) {
      sh.getRange(i + 1, 2, 1, 2).setValues([[nombre, c.activa === false ? 'No' : 'Sí']]);
      if (c.correo !== undefined) sh.getRange(i + 1, 6).setValue(correo);
      return { ok: true, codigo: String(c.codigo), mensaje: 'Guardado.' };
    }
    return { ok: false, error: 'No encuentro esa cajera.' };
  }
  const usados = v.slice(1).map(r => String(r[0]));
  let codigo;
  do { codigo = String(Math.floor(1000 + Math.random() * 9000)); } while (usados.indexOf(codigo) >= 0);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([["'" + codigo, nombre, 'Sí', new Date(), '', correo]]);
  return { ok: true, codigo: codigo, mensaje: nombre + ' dada de alta. Su código es ' + codigo + '.' };
}

/* ============================ INSTALACIÓN ============================ */

function instalarExtras() {
  hojaComent_(); hojaPagoProv_(); hojaPedidos_(); hojaCompl_(); hojaCortes_(); hojaCajeras_(); hojaActiv_();
  const pr = PropertiesService.getScriptProperties(); if (!pr.getProperty('COMPL_DESDE')) pr.setProperty('COMPL_DESDE', String(Date.now()));
  ['Correos extra', 'Correos quitados', 'Fecha constancia', 'Aviso CSF'].forEach(colCliente_);
  ScriptApp.getProjectTriggers().forEach(t => { if (['resumenCaja', 'recordatorioCSF', 'recordarComplementos', 'recordatorioCorte'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('resumenCaja').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('recordatorioCSF').timeBased().onMonthDay(5).atHour(11).create();
  ScriptApp.newTrigger('recordatorioCorte').timeBased().everyDays(1).atHour(HORA_AVISO_CORTE_).nearMinute(0).inTimezone('America/Mexico_City').create();
  ScriptApp.newTrigger('recordarComplementos').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(10).create();
  const txt = 'Extras instalados. Enlace de la cajera: ' + enlaceAppCaja();
  Logger.log(txt);
  return txt;
}
