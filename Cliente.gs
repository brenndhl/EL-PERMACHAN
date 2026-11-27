/****************************************************************
 * ÁREA DEL CLIENTE + FACTURAS SUBIDAS POR LA CONTADORA
 * — Restaurante El Permachan
 *
 *  MIS FACTURAS (app, pestaña "Mis facturas"):
 *   El cliente escribe su RFC, recibe un código en su correo registrado
 *   y ve sus solicitudes (pendiente / enviada / rechazada) y descarga
 *   el PDF y el XML de cada factura. La sesión dura 90 días en su móvil.
 *
 *  CARPETA DE LA CONTADORA:
 *   La contadora crea dentro de "Facturas El Permachan — subir aquí" (Drive)
 *   una carpeta con el NÚMERO DE SOLICITUD (597 = S-00597) y mete el PDF y
 *   el XML tal cual. (También vale subir 597.pdf / 597.xml o 1-48349.pdf.) Cada 10 minutos
 *   procesarFacturasSubidas() lee el XML (CFDI 4.0): RFC del cliente,
 *   serie/folio, UUID y total; lo une a su solicitud (por el folio del
 *   ticket en el nombre del archivo o, si no, por el RFC), manda la
 *   factura al cliente con los dos archivos adjuntos, la deja en su app
 *   y marca la solicitud como Enviada. Lo que no reconoce va a la
 *   subcarpeta "Revisar" y avisa por correo.
 *
 *  INSTALACIÓN: ejecutar instalarCarpetaFacturas() una vez.
 ****************************************************************/

const HOJA_FACTURAS = 'Facturas';
const COL_FACTURAS = ['ID solicitud', 'RFC', 'Cliente', 'Folio ticket', 'Serie-Folio', 'UUID', 'Total', 'Fecha factura',
  'PDF', 'XML', 'Subida', 'Enviada al cliente'];
const ESPERA_PAREJA_MIN_ = 30;   // minutos que se espera al PDF/XML que falta antes de procesar solo uno
const DIAS_SESION_ = 90;

/* ============================ HOJA "Facturas" ============================ */

function hojaFacturas_() {
  const ss = ss_();
  let sh = ss.getSheetByName(HOJA_FACTURAS);
  if (!sh) {
    sh = ss.insertSheet(HOJA_FACTURAS);
    sh.getRange(1, 1, 1, COL_FACTURAS.length).setValues([COL_FACTURAS]).setFontWeight('bold').setBackground('#1f4e5f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** ID de solicitud → datos de su factura subida */
function facturasMap_() {
  const m = {};
  leerTabla_(hojaFacturas_()).forEach(f => { if (f['ID solicitud']) m[String(f['ID solicitud'])] = f; });
  return m;
}

/* ============================ ÁREA DEL CLIENTE ============================ */

/** Paso 1: manda un código de 6 cifras al correo registrado del RFC. */
function clientePedirCodigo(rfcTexto) {
  const rfc = normalizarRFC(rfcTexto);
  if (!rfcValido(rfc)) return { ok: false, error: 'El RFC no tiene un formato válido (12 o 13 caracteres).' };
  const c = buscarCliente_(rfc);
  if (!c || !emailValido_(c.datos['Email'])) {
    return { ok: false, error: 'No encontramos ese RFC con un correo registrado. Pide una factura primero o escríbenos a ' + cfg_().RESPONDER_A + '.' };
  }
  if (!permitirEnvio_('mf_' + rfc)) return { ok: false, error: 'Has pedido demasiados códigos. Espera una hora.' };
  const codigo = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('mfc_' + rfc, codigo, 900);
  enviarCorreo_(c.datos['Email'], 'Tu código para ver tus facturas — ' + cfg_().NOMBRE_NEGOCIO,
    '<p>Tu código para entrar a <b>Mis facturas</b> es:</p><p style="font-size:28px;font-weight:bold;letter-spacing:4px">' + codigo + '</p>' +
    '<p>Caduca en 15 minutos. Si no lo has pedido tú, ignora este correo.</p>');
  return { ok: true, email: enmascararEmail(c.datos['Email']) };
}

/** Paso 2: con el código correcto devuelve un token que el móvil guarda (90 días). */
function clienteEntrar(rfcTexto, codigo, disp) {
  const rfc = normalizarRFC(rfcTexto);
  const cache = CacheService.getScriptCache();
  const kInt = 'mfi_' + rfc;
  const intentos = Number(cache.get(kInt) || 0);
  if (intentos >= 5) return { ok: false, error: 'Demasiados intentos. Pide un código nuevo dentro de 15 minutos.' };
  if (!codigo || cache.get('mfc_' + rfc) !== String(codigo).trim()) {
    cache.put(kInt, String(intentos + 1), 900);
    return { ok: false, error: 'Código incorrecto o caducado.' };
  }
  cache.remove('mfc_' + rfc);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('mft_' + token, rfc + '|' + (Date.now() + DIAS_SESION_ * 86400000));
  limpiarSesiones_();
  try { tocarDispositivo_(disp, { rfc: rfc }); } catch (e) { /* sin registro */ }
  return { ok: true, token: token };
}

function rfcDeToken_(token) {
  if (!token || !/^[a-f0-9]{64}$/.test(String(token))) return '';
  const v = PropertiesService.getScriptProperties().getProperty('mft_' + token);
  if (!v) return '';
  const p = v.split('|');
  if (Number(p[1]) < Date.now()) { PropertiesService.getScriptProperties().deleteProperty('mft_' + token); return ''; }
  return p[0];
}

function clienteSalir(token) {
  if (token && /^[a-f0-9]{64}$/.test(String(token))) PropertiesService.getScriptProperties().deleteProperty('mft_' + token);
  return { ok: true };
}

/** Borra sesiones caducadas (de vez en cuando). */
function limpiarSesiones_() {
  if (Math.random() > 0.1) return;
  const props = PropertiesService.getScriptProperties();
  const todo = props.getProperties();
  Object.keys(todo).forEach(k => { if (k.indexOf('mft_') === 0 && Number(String(todo[k]).split('|')[1]) < Date.now()) props.deleteProperty(k); });
}

/** Lista de solicitudes del cliente (más recientes primero). */
function clienteMisFacturas(token, disp) {
  const rfc = rfcDeToken_(token);
  if (!rfc) return { ok: false, sesion: false, error: 'Tu sesión caducó. Vuelve a entrar con tu RFC.' };
  try { tocarDispositivo_(disp, { rfc: rfc }); } catch (e) { /* sin registro */ }
  const tz = Session.getScriptTimeZone();
  const facts = facturasMap_();
  let pagos = {};
  try { pagos = pagosMap_(); } catch (e) { /* sin hoja de pagos */ }
  const c = buscarCliente_(rfc);
  const lista = leerTabla_(hoja_(HOJAS.SOLICITUDES))
    .filter(s => s['ID'] && normalizarRFC(s['RFC']) === rfc)
    .sort((a, b) => (b['Fecha solicitud'] instanceof Date ? b['Fecha solicitud'].getTime() : 0) - (a['Fecha solicitud'] instanceof Date ? a['Fecha solicitud'].getTime() : 0))
    .slice(0, 100)
    .map(s => {
      const f = facts[String(s['ID'])];
      return {
        id: s['ID'], folio: s['Folio ticket'], estado: s['Estado'], pago: s['Forma de pago'],
        fecha: s['Fecha solicitud'] instanceof Date ? Utilities.formatDate(s['Fecha solicitud'], tz, 'dd/MM/yyyy') : '',
        enviada: s['Fecha emitida'] instanceof Date ? Utilities.formatDate(s['Fecha emitida'], tz, 'dd/MM/yyyy') : '',
        motivo: s['Estado'] === 'Rechazada' ? String(s['Notas / motivo'] || '') : '',
        porDefinir: /definir/i.test(String(s['Forma de pago'] || '')), pagada: !!pagos[String(s['ID'])],
        factura: f ? { serieFolio: f['Serie-Folio'], total: f['Total'], uuid: f['UUID'], pdf: !!f['PDF'], xml: !!f['XML'],
          fecha: f['Fecha factura'] instanceof Date ? Utilities.formatDate(f['Fecha factura'], tz, 'dd/MM/yyyy') : String(f['Fecha factura'] || '') } : null
      };
    });
  return { ok: true, rfc: rfc, nombre: c ? c.datos['Nombre o razón social'] : '', lista: lista };
}

/** Descarga del PDF o XML de una factura del propio cliente. */
function clienteDescargar(token, id, tipo) {
  const rfc = rfcDeToken_(token);
  if (!rfc) return { ok: false, sesion: false, error: 'Tu sesión caducó. Vuelve a entrar con tu RFC.' };
  const f = facturasMap_()[String(id)];
  if (!f || normalizarRFC(f['RFC']) !== rfc) return { ok: false, error: 'No encuentro esa factura.' };
  const url = tipo === 'xml' ? f['XML'] : f['PDF'];
  const b = blobDeUrl_(url, 'Factura_' + (f['Serie-Folio'] || id));
  if (!b) return { ok: false, error: 'No se pudo abrir el archivo. Escríbenos a ' + cfg_().RESPONDER_A + '.' };
  return { ok: true, tipo: b.getContentType(), nombre: b.getName(), base64: Utilities.base64Encode(b.getBytes()) };
}

/* ============================ CARPETA DE LA CONTADORA ============================ */

function carpetaSubida_() {
  const id = PropertiesService.getScriptProperties().getProperty('ID_CARPETA_FACTURAS');
  if (!id) throw new Error('Falta la carpeta de facturas. Ejecuta instalarCarpetaFacturas().');
  return DriveApp.getFolderById(id);
}

function subcarpeta_(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

/** Ejecutar UNA vez: crea la carpeta, la comparte con la contadora y programa la revisión cada 10 minutos. */
function instalarCarpetaFacturas() {
  const props = PropertiesService.getScriptProperties();
  let carpeta = null;
  try { if (props.getProperty('ID_CARPETA_FACTURAS')) carpeta = DriveApp.getFolderById(props.getProperty('ID_CARPETA_FACTURAS')); } catch (e) { carpeta = null; }
  if (!carpeta) {
    carpeta = DriveApp.createFolder('Facturas El Permachan — subir aquí');
    props.setProperty('ID_CARPETA_FACTURAS', carpeta.getId());
  }
  subcarpeta_(carpeta, 'Procesadas');
  subcarpeta_(carpeta, 'Revisar');
  const contadora = (typeof correoContadora_ === 'function') ? correoContadora_() : String(cfg_().CORREO_CONTADORA || '').trim();
  if (contadora) { try { carpeta.addEditor(contadora); } catch (e) { Logger.log('No se pudo compartir con ' + contadora + ': ' + e); } }
  hojaFacturas_();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'procesarFacturasSubidas') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('procesarFacturasSubidas').timeBased().everyMinutes(10).create();
  const txt = 'Carpeta: ' + carpeta.getName() + ' — ' + carpeta.getUrl() + ' · compartida con: ' + (contadora || '(nadie)') +
    ' · revisión automática cada 10 minutos.';
  Logger.log(txt);
  return txt;
}

/** Enlace de la carpeta (para la guía de la contadora). */
function enlaceCarpetaFacturas() {
  const u = carpetaSubida_().getUrl();
  Logger.log(u);
  return u;
}

/** Lee un CFDI (XML) y devuelve sus datos principales, o null si no es un CFDI. */
function leerCFDI_(texto) {
  let raiz;
  try { raiz = XmlService.parse(texto).getRootElement(); } catch (e) { return null; }
  if (raiz.getName() !== 'Comprobante') return null;
  const at = (el, n) => { const a = el && el.getAttribute(n); return a ? a.getValue() : ''; };
  const buscar = (el, nombre) => {
    if (el.getName() === nombre) return el;
    const hijos = el.getChildren();
    for (let i = 0; i < hijos.length; i++) { const r = buscar(hijos[i], nombre); if (r) return r; }
    return null;
  };
  const rec = buscar(raiz, 'Receptor'), tfd = buscar(raiz, 'TimbreFiscalDigital');
  return {
    rfc: normalizarRFC(at(rec, 'Rfc')), nombre: at(rec, 'Nombre'),
    serie: at(raiz, 'Serie'), folio: at(raiz, 'Folio'), fecha: at(raiz, 'Fecha'), total: at(raiz, 'Total'),
    metodo: at(raiz, 'MetodoPago'), uuid: at(tfd, 'UUID').toUpperCase()
  };
}

/** Número de solicitud en un nombre: "597", "S-00597", "s597", "Solicitud 597" → "S-00597". */
function idDeNombre_(nombre) {
  const m = /^\s*(?:solicitud\s*)?(?:s\s*-?\s*)?0*(\d{1,6})\s*$/i.exec(String(nombre || ''));
  return m ? 'S-' + ('00000' + m[1]).slice(-5) : '';
}

/** Folio del ticket escrito en el nombre del archivo (1-48349, 1_48349, #1-48349 …). */
function folioDeNombre_(nombre) {
  const m = /(\d{1,3})\s*[-_ ]\s*(\d{3,12})/.exec(String(nombre || ''));
  return m ? m[1] + '-' + m[2] : '';
}

/** Disparador cada 10 minutos: procesa lo que la contadora ha subido. */
function procesarFacturasSubidas() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return 'Ocupado';
  try {
    const carpeta = carpetaSubida_();
    const procesadas = subcarpeta_(carpeta, 'Procesadas'), revisar = subcarpeta_(carpeta, 'Revisar');
    const grupos = {};
    const it = carpeta.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      const n = f.getName(), ext = (/\.([a-z0-9]+)$/i.exec(n) || [])[1];
      if (!ext || ['pdf', 'xml'].indexOf(ext.toLowerCase()) < 0) continue;
      const base = n.replace(/\.[a-z0-9]+$/i, '').trim().toLowerCase();
      (grupos[base] = grupos[base] || {})[ext.toLowerCase()] = f;
    }
    // Carpetas con el número de solicitud (597, S-00597…): su PDF y su XML, con cualquier nombre
    const subs = carpeta.getFolders();
    while (subs.hasNext()) {
      const sub = subs.next();
      if (['Procesadas', 'Revisar'].indexOf(sub.getName()) >= 0) continue;
      const g = { carpeta: sub, id: idDeNombre_(sub.getName()) };
      const fs = sub.getFiles();
      while (fs.hasNext()) {
        const f = fs.next(), ext = ((/\.([a-z0-9]+)$/i.exec(f.getName()) || [])[1] || '').toLowerCase();
        if ((ext === 'pdf' || ext === 'xml') && !g[ext]) g[ext] = f;
      }
      if (g.pdf || g.xml) grupos['carpeta:' + sub.getName()] = g;
    }
    const ahora = Date.now(), res = [];
    Object.keys(grupos).forEach(base => {
      const g = grupos[base];
      const ultimo = Math.max(g.pdf ? g.pdf.getLastUpdated().getTime() : 0, g.xml ? g.xml.getLastUpdated().getTime() : 0);
      if ((!g.pdf || !g.xml) && ahora - ultimo < ESPERA_PAREJA_MIN_ * 60000) return;   // esperando a la pareja
      try { res.push(procesarGrupo_(base, g, procesadas, revisar)); }
      catch (e) { res.push(base + ': error ' + e); Logger.log(e.stack || e); }
    });
    Logger.log(res.join('\n') || 'Nada que procesar');
    return res;
  } finally { lock.releaseLock(); }
}

function procesarGrupo_(base, g, procesadas, revisar) {
  const cfdi = g.xml ? leerCFDI_(g.xml.getBlob().getDataAsString('UTF-8')) : null;
  const folio = folioDeNombre_(base);
  const facts = facturasMap_();
  if (cfdi && cfdi.uuid && Object.keys(facts).some(k => String(facts[k]['UUID']).toUpperCase() === cfdi.uuid)) {
    return aRevisar_(g, revisar, base, 'Esta factura (UUID ' + cfdi.uuid + ') ya se había subido antes.');
  }
  if (g.xml && !cfdi) return aRevisar_(g, revisar, base, 'El XML no es un CFDI válido.');
  const sols = leerTabla_(hoja_(HOJAS.SOLICITUDES)).filter(s => s['ID'] && !facts[String(s['ID'])] && s['Estado'] !== 'Rechazada');
  let s = null, motivo = '';
  const idSol = g.id || (g.carpeta ? '' : idDeNombre_(base));
  if (g.carpeta && !idSol) return aRevisar_(g, revisar, base, 'La carpeta "' + g.carpeta.getName() + '" no es un número de solicitud (por ejemplo 597).');
  if (idSol) {
    s = sols.filter(x => String(x['ID']) === idSol)[0] || null;
    if (!s) return aRevisar_(g, revisar, base, 'No hay ninguna solicitud ' + idSol + ' esperando factura (o ya tiene factura).');
    if (cfdi && cfdi.rfc && normalizarRFC(s['RFC']) !== cfdi.rfc) {
      return aRevisar_(g, revisar, base, 'La solicitud ' + idSol + ' es de ' + s['RFC'] + ' pero la factura está a nombre de ' + cfdi.rfc + '.');
    }
  }
  if (!s && folio) {
    const porFolio = sols.filter(x => normalizarFolio(x['Folio ticket']) === folio);
    if (porFolio.length === 1) s = porFolio[0];
    else motivo = porFolio.length ? 'Hay varias solicitudes con el ticket ' + folio + '.' : 'No hay ninguna solicitud pendiente con el ticket ' + folio + '.';
    if (s && cfdi && cfdi.rfc && normalizarRFC(s['RFC']) !== cfdi.rfc) {
      return aRevisar_(g, revisar, base, 'El ticket ' + folio + ' es de ' + s['RFC'] + ' pero la factura está a nombre de ' + cfdi.rfc + '.');
    }
  }
  if (!s && cfdi && cfdi.rfc) {
    const delRfc = sols.filter(x => normalizarRFC(x['RFC']) === cfdi.rfc);
    const abiertas = delRfc.filter(x => x['Estado'] === 'Pendiente' || x['Estado'] === 'Emitida');
    if (abiertas.length === 1) s = abiertas[0];
    else if (!abiertas.length && delRfc.length === 1) s = delRfc[0];
    else motivo = (motivo ? motivo + ' ' : '') + (abiertas.length > 1
      ? 'El cliente ' + cfdi.rfc + ' tiene ' + abiertas.length + ' solicitudes abiertas: pon el folio del ticket en el nombre del archivo.'
      : 'No hay ninguna solicitud del RFC ' + cfdi.rfc + ' esperando factura.');
  }
  if (!s) return aRevisar_(g, revisar, base, motivo || 'No sé de qué cliente es: sube también el XML o pon el folio del ticket en el nombre.');

  // Mover a Procesadas y registrar
  const destino = subcarpeta_(procesadas, normalizarRFC(s['RFC']) || 'SIN_RFC');
  if (g.carpeta) { g.carpeta.setName(s['ID'] + ' — ticket ' + s['Folio ticket']); g.carpeta.moveTo(destino); }
  else { if (g.pdf) g.pdf.moveTo(destino); if (g.xml) g.xml.moveTo(destino); }
  const serieFolio = cfdi ? [cfdi.serie, cfdi.folio].filter(String).join('-') : '';
  const enviado = mandarFacturaCliente_(s, g, cfdi, serieFolio);
  agregarFila_(hojaFacturas_(), COL_FACTURAS, {
    'ID solicitud': s['ID'], 'RFC': s['RFC'], 'Cliente': s['Nombre o razón social'], 'Folio ticket': s['Folio ticket'],
    'Serie-Folio': serieFolio, 'UUID': cfdi ? cfdi.uuid : '', 'Total': cfdi ? Number(cfdi.total) || cfdi.total : '',
    'Fecha factura': cfdi && cfdi.fecha ? new Date(cfdi.fecha.slice(0, 19)) : '', 'PDF': g.pdf ? g.pdf.getUrl() : '',
    'XML': g.xml ? g.xml.getUrl() : '', 'Subida': new Date(), 'Enviada al cliente': enviado ? 'Sí — ' + s['Email'] : 'No se pudo enviar'
  });
  const f = buscarSolicitud_(s['ID']);
  if (f) {
    const ahora = new Date();
    const cambios = { 'Estado': 'Enviada', 'Fecha estado': ahora, 'Último aviso': 'Enviada' };
    if (!(f.datos['Fecha emitida'] instanceof Date)) cambios['Fecha emitida'] = ahora;
    actualizarFila_(hoja_(HOJAS.SOLICITUDES), COL_SOLICITUDES, f.fila, cambios);
  }
  return base + ' → ' + s['ID'] + ' (' + s['RFC'] + ')' + (enviado ? ' enviada' : ' SIN correo');
}

/** Correo al cliente con el PDF y el XML adjuntos. */
function mandarFacturaCliente_(s, g, cfdi, serieFolio) {
  if (!emailValido_(s['Email'])) return false;
  const conf = cfg_();
  const adj = [];
  if (g.pdf) adj.push(g.pdf.getBlob().setName('Factura_' + (serieFolio || s['Folio ticket']) + '.pdf'));
  if (g.xml) adj.push(g.xml.getBlob().setName('Factura_' + (serieFolio || s['Folio ticket']) + '.xml'));
  const app = String(conf.URL_APP || 'https://brenndhl.github.io/EL-PERMACHAN/');
  const mis = app + (app.indexOf('?') < 0 ? '?' : '&') + 'v=mis';
  const html = '<div style="font-family:Arial,sans-serif;max-width:560px;color:#1d1a17">' +
    '<div style="background:#141110;color:#f3ead6;padding:18px 20px;border-radius:12px 12px 0 0;border-bottom:3px solid #a8774a">' +
    '<img src="' + esc_(app) + 'logo-256.png" width="56" height="56" style="border-radius:50%;vertical-align:middle;margin-right:12px">' +
    '<b style="font-size:20px;vertical-align:middle">¡Ya está tu factura!</b></div>' +
    '<div style="background:#fff;border:1px solid #e5dccb;border-top:0;padding:20px;border-radius:0 0 12px 12px">' +
    '<p>Hola, tu factura del ticket <b>' + esc_(s['Folio ticket']) + '</b>' + (serieFolio ? ' (factura ' + esc_(serieFolio) + ')' : '') +
    ' ya está en la app de El Permachan. También te la dejamos adjunta en este correo (PDF y XML).</p>' +
    '<p style="text-align:center;margin:26px 0"><a href="' + esc_(mis) + '" style="background:#141110;color:#f3ead6;text-decoration:none;' +
    'font-weight:bold;font-size:18px;padding:14px 28px;border-radius:10px;display:inline-block">Descárgala aquí</a></p>' +
    '<p style="font-size:14px;color:#6b625a">En la app, en <b>Mis facturas</b>, tienes todas tus facturas para descargarlas cuando quieras.</p>' +
    '<p>Gracias por tu preferencia.<br>' + esc_(conf.NOMBRE_NEGOCIO) + '</p></div></div>';
  return enviarCorreo_(s['Email'], '¡Ya está tu factura! — ' + conf.NOMBRE_NEGOCIO + ' — ticket ' + s['Folio ticket'], html, adj);
}

function aRevisar_(g, revisar, base, motivo) {
  if (g.carpeta) g.carpeta.moveTo(revisar);
  else { if (g.pdf) g.pdf.moveTo(revisar); if (g.xml) g.xml.moveTo(revisar); }
  const nombres = (g.carpeta ? 'la carpeta ' + g.carpeta.getName() + ': ' : '') + [g.pdf && g.pdf.getName(), g.xml && g.xml.getName()].filter(Boolean).join(' y ');
  enviarCorreo_(cfg_().CORREOS_AVISO, 'Factura por revisar: ' + nombres,
    '<p>No pude asociar <b>' + esc_(nombres) + '</b> a ninguna solicitud.</p><p><b>Motivo:</b> ' + esc_(motivo) + '</p>' +
    '<p>Están en la subcarpeta <i>Revisar</i>. Revisa el número de solicitud (la carpeta debe llamarse solo con el número, por ejemplo <b>597</b>) ' +
    'y vuelve a moverla a la carpeta principal; se procesa sola en unos minutos.</p>');
  return base + ' → Revisar: ' + motivo;
}

/* ============================ USO DE LA APP ============================ */
// La app (GitHub) avisa cada vez que se abre (?ping=1&d=<id del móvil>&p=<plataforma>&inst=<1 si está instalada>)
// y, al entrar en "Mis facturas", se une ese móvil con el RFC del cliente.

const HOJA_DISP = 'App - dispositivos';
const COL_DISP = ['ID dispositivo', 'Plataforma', 'Instalada', 'Primera vez', 'Última vez', 'Aperturas', 'RFC', 'Cliente'];

function hojaDisp_() {
  const ss = ss_();
  let sh = ss.getSheetByName(HOJA_DISP);
  if (!sh) {
    sh = ss.insertSheet(HOJA_DISP);
    sh.getRange(1, 1, 1, COL_DISP.length).setValues([COL_DISP]).setFontWeight('bold').setBackground('#1f4e5f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function idDispValido_(d) { return /^[a-z0-9]{8,40}$/.test(String(d || '')); }

/** Crea o actualiza la fila del móvil. cambios: { abrir:bool, inst:bool, plat, rfc } */
function tocarDispositivo_(id, cambios) {
  if (!idDispValido_(id)) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(4000)) return;
  try {
    const sh = hojaDisp_(), ahora = new Date();
    const n = sh.getLastRow();
    const ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
    const i = ids.indexOf(String(id));
    if (i < 0) {
      const c = cambios.rfc ? buscarCliente_(cambios.rfc) : null;
      agregarFila_(sh, COL_DISP, { 'ID dispositivo': id, 'Plataforma': cambios.plat || '', 'Instalada': cambios.inst ? 'Sí' : 'No',
        'Primera vez': ahora, 'Última vez': ahora, 'Aperturas': cambios.abrir ? 1 : 0, 'RFC': cambios.rfc || '',
        'Cliente': c ? c.datos['Nombre o razón social'] : '' });
      return;
    }
    const fila = i + 2, act = objeto_(COL_DISP, sh.getRange(fila, 1, 1, COL_DISP.length).getValues()[0]);
    const upd = { 'Última vez': ahora };
    if (cambios.abrir) upd['Aperturas'] = (Number(act['Aperturas']) || 0) + 1;
    if (cambios.inst && act['Instalada'] !== 'Sí') upd['Instalada'] = 'Sí';
    if (cambios.plat && !act['Plataforma']) upd['Plataforma'] = cambios.plat;
    if (cambios.rfc && act['RFC'] !== cambios.rfc) {
      const c = buscarCliente_(cambios.rfc);
      upd['RFC'] = cambios.rfc; upd['Cliente'] = c ? c.datos['Nombre o razón social'] : '';
    }
    actualizarFila_(sh, COL_DISP, fila, upd);
  } finally { lock.releaseLock(); }
}

/** Lo llama doGet con ?ping=1 cada vez que se abre la app. */
function registrarUsoApp_(p) {
  const plat = ['Android', 'iPhone', 'Ordenador'].indexOf(String(p.p)) >= 0 ? String(p.p) : '';
  tocarDispositivo_(String(p.d || ''), { abrir: true, inst: String(p.inst) === '1', plat: plat });
}

/** Resumen y lista para el panel. */
function usoAppResumen_() {
  const filas = leerTabla_(hojaDisp_()).filter(r => r['ID dispositivo']);
  const rfcs = {};
  filas.forEach(r => { if (r['RFC']) rfcs[r['RFC']] = true; });
  return { moviles: filas.length, instaladas: filas.filter(r => r['Instalada'] === 'Sí').length, clientes: Object.keys(rfcs).length, filas: filas };
}

function panelUsoApp(k) {
  if (!esPanel_(k)) return { ok: false, error: 'Este enlace no es válido.' };
  const tz = Session.getScriptTimeZone();
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm') : '';
  const u = usoAppResumen_();
  const lista = u.filas.sort((a, b) => (b['Última vez'] instanceof Date ? b['Última vez'].getTime() : 0) - (a['Última vez'] instanceof Date ? a['Última vez'].getTime() : 0))
    .map(r => ({ cliente: r['Cliente'] || '', rfc: r['RFC'] || '', plat: r['Plataforma'] || '', inst: r['Instalada'] === 'Sí',
      primera: fmt(r['Primera vez']), ultima: fmt(r['Última vez']), aperturas: Number(r['Aperturas']) || 0 }));
  return { ok: true, moviles: u.moviles, instaladas: u.instaladas, clientes: u.clientes, lista: lista };
}
