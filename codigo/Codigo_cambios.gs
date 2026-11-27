/* ===== Sustituye en Código.gs: doGet y correosDe_ ===== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const negocio = cfg_().NOMBRE_NEGOCIO;
  const salida = (o, titulo) => o.setTitle(titulo + ' — ' + negocio).addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (p.ping) { try { registrarUsoApp_(p); } catch (err) { Logger.log(err); } return ContentService.createTextOutput('ok'); }
  if (p.v === 'mis') {
    const tm = HtmlService.createTemplateFromFile('MisFacturas');
    tm.disp = /^[a-z0-9]{8,40}$/.test(String(p.d || '')) ? String(p.d) : '';
    return salida(tm.evaluate(), 'Mis facturas');
  }
  if (p.v === 'ayuda') return salida(HtmlService.createHtmlOutputFromFile('Ayuda'), 'Ayuda');
  if (p.k && esCaja_(p.k)) {
    const tc = HtmlService.createTemplateFromFile('Caja');
    tc.clave = p.k;
    return salida(tc.evaluate(), 'Caja');
  }
  if (p.k) {
    const t = HtmlService.createTemplateFromFile('Admin');
    t.clave = p.k;
    t.id = p.id || '';
    t.accion = p.e === 'Rechazada' ? 'Rechazada' : (p.id ? 'Enviada' : '');
    return salida(t.evaluate(), 'Facturas');
  }
  return salida(HtmlService.createHtmlOutputFromFile('Index'), 'Solicitar factura');
}

/** Correos del RFC (respeta los añadidos y quitados desde el panel). */
function correosDe_(rfc, datos) {
  return correosCliente_(rfc).slice(0, 8);
}
