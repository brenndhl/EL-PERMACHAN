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
 *   La contadora sube el PDF y el XML de cada factura a la carpeta
 *   "Facturas El Permachan — subir aquí" (Drive). Cada 10 minutos
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
function clienteEntrar(rfcTexto, codigo) {
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
function clienteMisFacturas(token) {
  const rfc = rfcDeToken_(token);
  if (!rfc) return { ok: false, sesion: false, error: 'Tu sesión caducó. Vuelve a entrar con tu RFC.' };
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
  if (folio) {
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
  if (g.pdf) g.pdf.moveTo(destino);
  if (g.xml) g.xml.moveTo(destino);
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
  const html = '<p>Hola, aquí tienes tu factura del ticket <b>' + esc_(s['Folio ticket']) + '</b>' +
    (serieFolio ? ' (factura ' + esc_(serieFolio) + ')' : '') + '. Van adjuntos el PDF y el XML.</p>' +
    '<p>También la tienes siempre en la app, en <b>Mis facturas</b>, para descargarla cuando quieras:<br>' +
    '<a href="' + esc_(app) + '">' + esc_(app) + '</a></p>' +
    '<p>Gracias por tu preferencia.<br>' + esc_(conf.NOMBRE_NEGOCIO) + '</p>';
  return enviarCorreo_(s['Email'], 'Tu factura — ' + conf.NOMBRE_NEGOCIO + ' — ticket ' + s['Folio ticket'], html, adj);
}

function aRevisar_(g, revisar, base, motivo) {
  if (g.pdf) g.pdf.moveTo(revisar);
  if (g.xml) g.xml.moveTo(revisar);
  const nombres = [g.pdf && g.pdf.getName(), g.xml && g.xml.getName()].filter(Boolean).join(' y ');
  enviarCorreo_(cfg_().CORREOS_AVISO, 'Factura por revisar: ' + nombres,
    '<p>No pude asociar <b>' + esc_(nombres) + '</b> a ninguna solicitud.</p><p><b>Motivo:</b> ' + esc_(motivo) + '</p>' +
    '<p>Los archivos están en la subcarpeta <i>Revisar</i>. Corrige el nombre (por ejemplo <b>1-48349.pdf</b> y <b>1-48349.xml</b>, ' +
    'con el folio del ticket) y vuelve a moverlos a la carpeta principal; se procesan solos en unos minutos.</p>');
  return base + ' → Revisar: ' + motivo;
}
