/****************************************************************
 * CUADRO DE MANDO + INFORME QUINCENAL — RESTAURANTE EL PERMACHAN
 * Archivo aparte: en Apps Script pulsa + > Script, llámalo "Informe"
 * y pega este contenido. Luego elige "instalarInformeQuincenal" y pulsa Ejecutar.
 *
 *  - Hoja "Panel": se actualiza sola cada mañana.
 *  - Informe por correo: los días 1 y 16 de cada mes a las 10:00 (zona horaria del proyecto),
 *    con la quincena anterior (del 1 al 15, o del 16 a fin de mes).
 *    A mano: informeQuincenal() · un mes concreto: informeDeMes(2026, 11).
 *  - Destinatarios: fila CORREOS_INFORME de Config (separados por comas; si está vacía, CORREOS_AVISO).
 ****************************************************************/

const MESES_ = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS_ATRASO_ = 3;   // una pendiente con más días que esto sale como "atrasada"
const NOTA_QUINCENAL_ = 'Este informe se envía solo los días 1 y 16 de cada mes.';

/** Ejecutar UNA vez: crea el Panel, programa los envíos y manda el informe de la última quincena como prueba. */
function instalarInformeQuincenal() {
  const sh = hoja_(HOJAS.CONFIG);
  const claves = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  if (claves.indexOf('CORREOS_INFORME') < 0) sh.appendRow(['CORREOS_INFORME', 'elpermachan@gmail.com, jovhys@hotmail.com']);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['informeMensual', 'informeSemanal', 'informeQuincenal', 'actualizarPanel'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('actualizarPanel').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('informeQuincenal').timeBased().onMonthDay(1).atHour(10).create();
  ScriptApp.newTrigger('informeQuincenal').timeBased().onMonthDay(16).atHour(10).create();
  actualizarPanel();
  const r = informeQuincenal_();
  const txt = 'Listo: informe quincenal los días 1 y 16 a las 10:00 (zona ' + Session.getScriptTimeZone() + '). ' +
    'Prueba enviada a: ' + r.para;
  Logger.log(txt);
  return txt;
}

/** Disparador de los días 1 y 16. */
function informeQuincenal() {
  actualizarPanel();
  return informeQuincenal_();
}

/** Quincena anterior: si hoy es 1–15 → del 16 a fin del mes pasado; si es 16–31 → del 1 al 15 de este mes. */
function informeQuincenal_() {
  const h = new Date();
  let desde, hasta, ant;
  if (h.getDate() < 16) {
    desde = new Date(h.getFullYear(), h.getMonth() - 1, 16);
    hasta = new Date(h.getFullYear(), h.getMonth(), 1);
    ant = new Date(h.getFullYear(), h.getMonth() - 1, 1);
  } else {
    desde = new Date(h.getFullYear(), h.getMonth(), 1);
    hasta = new Date(h.getFullYear(), h.getMonth(), 16);
    ant = new Date(h.getFullYear(), h.getMonth() - 1, 16);
  }
  const fin = new Date(hasta.getTime() - 86400000);
  const titulo = 'del ' + desde.getDate() + ' al ' + fin.getDate() + ' de ' + MESES_[fin.getMonth()] + ' de ' + fin.getFullYear();
  return enviarInforme_(desde, hasta, ant, titulo, 'la quincena anterior', 'Informe quincenal', NOTA_QUINCENAL_);
}

/** Para mandar a mano el informe de un mes: informeDeMes(2026, 11) = noviembre 2026. */
function informeDeMes(anio, mes) {
  const y = Number(anio), m = Number(mes) - 1;
  return enviarInforme_(new Date(y, m, 1), new Date(y, m + 1, 1), new Date(y, m - 1, 1), MESES_[m] + ' ' + y, 'el mes anterior', 'Informe mensual', '');
}

/* ============================ CÁLCULOS ============================ */

function solicitudesInforme_() {
  return leerTabla_(hoja_(HOJAS.SOLICITUDES)).filter(s => s['ID'] && s['Fecha solicitud'] instanceof Date);
}

function enviada_(s) { return s['Estado'] === 'Enviada' || s['Estado'] === 'Emitida'; }

/** Días desde la solicitud hasta que se marcó enviada (solo de la web/formulario nuevo, no del histórico). */
function diasHastaEnviar_(s) {
  if (s['Origen'] === 'Histórico' || !(s['Fecha emitida'] instanceof Date)) return null;
  const d = (s['Fecha emitida'] - s['Fecha solicitud']) / 86400000;
  return d >= 0 ? d : null;
}

function media_(xs) { return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null; }
function fmtDias_(d) {
  if (d == null) return '—';
  if (d < 1) return Math.max(1, Math.round(d * 24)) + ' h';
  const r = Math.round(d * 10) / 10;
  return r + (r === 1 ? ' día' : ' días');
}
function contar_(lista, f) {
  const o = {};
  lista.forEach(x => { const k = f(x) || '(sin dato)'; o[k] = (o[k] || 0) + 1; });
  return Object.keys(o).map(k => [k, o[k]]).sort((a, b) => b[1] - a[1]);
}

function statsMes_(todas, y, m) { return statsRango_(todas, new Date(y, m, 1), new Date(y, m + 1, 1)); }

/** Estadísticas de las solicitudes recibidas entre desde (incluido) y hasta (excluido). */
function statsRango_(todas, desde, hasta) {
  const en = d => d instanceof Date && d >= desde && d < hasta;
  const rec = todas.filter(s => en(s['Fecha solicitud']));
  const dias = rec.map(diasHastaEnviar_).filter(d => d != null);
  return {
    recibidas: rec.length,
    enviadas: rec.filter(enviada_).length,
    rechazadas: rec.filter(s => s['Estado'] === 'Rechazada').length,
    pendientes: rec.filter(s => s['Estado'] === 'Pendiente').length,
    enviadasEnMes: todas.filter(s => en(s['Fecha emitida'])).length,
    diasMedios: media_(dias),
    web: rec.filter(s => s['Origen'] === 'Web').length,
    clientes: contar_(rec, s => s['Nombre o razón social'] || s['RFC']),
    pagos: contar_(rec, s => s['Forma de pago']),
    nuevos: leerTabla_(hoja_(HOJAS.CLIENTES)).filter(c => en(c['Fecha alta']) && c['Origen'] !== 'Migración').length
  };
}

function atrasadas_(todas) {
  const lim = Date.now() - DIAS_ATRASO_ * 86400000;
  return todas.filter(s => s['Estado'] === 'Pendiente' && s['Fecha solicitud'].getTime() < lim)
    .sort((a, b) => a['Fecha solicitud'] - b['Fecha solicitud']);
}

function fechaCorta_(d) { return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear(); }

/* ============================ PANEL (hoja) ============================ */

function actualizarPanel() {
  const ss = ss_();
  let sh = ss.getSheetByName('Panel');
  if (!sh) sh = ss.insertSheet('Panel');
  sh.clear();
  try { sh.getCharts().forEach(c => sh.removeChart(c)); } catch (e) { /* sin gráficos */ }

  const todas = solicitudesInforme_();
  const hoy = new Date();
  sh.getRange('A1').setValue('Panel de facturas — ' + cfg_().NOMBRE_NEGOCIO).setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue('Actualizado: ' + fechaCorta_(hoy) + ' (se actualiza solo cada mañana)').setFontColor('#666666');

  const pend = todas.filter(s => s['Estado'] === 'Pendiente');
  const atr = atrasadas_(todas);
  const esteMes = statsMes_(todas, hoy.getFullYear(), hoy.getMonth());
  const kpis = [
    ['Pendientes ahora', pend.length],
    ['Pendientes con más de ' + DIAS_ATRASO_ + ' días', atr.length],
    ['Recibidas este mes', esteMes.recibidas],
    ['Enviadas este mes', esteMes.enviadasEnMes],
    ['Tiempo medio hasta enviar (este mes)', fmtDias_(esteMes.diasMedios)]
  ];
  sh.getRange(4, 1, kpis.length, 2).setValues(kpis);
  sh.getRange(4, 1, kpis.length, 1).setFontWeight('bold');
  if (atr.length) sh.getRange(5, 2).setBackground('#f4cccc');

  const fila0 = 11;
  sh.getRange(fila0 - 1, 1).setValue('Últimos 12 meses').setFontWeight('bold');
  const cab = ['Mes', 'Recibidas', 'Enviadas', 'Rechazadas', 'Pendientes', 'Tiempo medio hasta enviar', 'Pedidas por la web'];
  const filas = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
    const st = statsMes_(todas, d.getFullYear(), d.getMonth());
    filas.push([MESES_[d.getMonth()] + ' ' + d.getFullYear(), st.recibidas, st.enviadas, st.rechazadas, st.pendientes,
      fmtDias_(st.diasMedios), st.recibidas ? Math.round(st.web / st.recibidas * 100) + '%' : '—']);
  }
  sh.getRange(fila0, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1f4e5f').setFontColor('#ffffff');
  sh.getRange(fila0 + 1, 1, filas.length, cab.length).setValues(filas);

  const desde = new Date(hoy.getFullYear(), hoy.getMonth() - 11, 1);
  const top = contar_(todas.filter(s => s['Fecha solicitud'] >= desde), s => s['Nombre o razón social'] || s['RFC']).slice(0, 10);
  const fT = fila0 + filas.length + 3;
  sh.getRange(fT - 1, 1).setValue('Clientes que más facturas piden (12 meses)').setFontWeight('bold');
  sh.getRange(fT, 1, 1, 2).setValues([['Cliente', 'Solicitudes']]).setFontWeight('bold');
  if (top.length) sh.getRange(fT + 1, 1, top.length, 2).setValues(top);

  const fA = fT + Math.max(top.length, 1) + 3;
  sh.getRange(fA - 1, 1).setValue('Pendientes con más de ' + DIAS_ATRASO_ + ' días').setFontWeight('bold');
  sh.getRange(fA, 1, 1, 4).setValues([['ID', 'Fecha', 'Cliente', 'Ticket']]).setFontWeight('bold');
  const la = atr.slice(0, 50).map(s => [s['ID'], fechaCorta_(s['Fecha solicitud']), s['Nombre o razón social'] || s['RFC'], s['Folio ticket']]);
  if (la.length) sh.getRange(fA + 1, 1, la.length, 4).setValues(la);
  else sh.getRange(fA + 1, 1).setValue('Ninguna 🎉');

  try {
    const graf = sh.newChart().asColumnChart()
      .addRange(sh.getRange(fila0, 1, filas.length + 1, 3))
      .setPosition(3, 9, 0, 0)
      .setOption('title', 'Facturas por mes').setOption('legend', { position: 'bottom' })
      .setOption('colors', ['#9fb8c1', '#1f4e5f'])
      .build();
    sh.insertChart(graf);
  } catch (e) { Logger.log('Sin gráfico: ' + e); }

  sh.setColumnWidth(1, 300);
}

/* ============================ CORREO ============================ */

function enviarInforme_(desde, hasta, desdeAnt, titulo, periodoAnt, tipo, nota) {
  const conf = cfg_();
  const todas = solicitudesInforme_();
  const st = statsRango_(todas, desde, hasta);
  const ant = statsRango_(todas, desdeAnt, desde);
  const atr = atrasadas_(todas);
  const dif = st.recibidas - ant.recibidas;
  const flecha = ant.recibidas ? (dif > 0 ? '▲ ' + dif + ' más' : dif < 0 ? '▼ ' + (-dif) + ' menos' : 'igual') + ' que ' + periodoAnt : '';

  const caja = (t, valor, color) => '<td style="padding:12px 14px;background:' + color + ';border-radius:8px;text-align:center">' +
    '<div style="font-size:26px;font-weight:700">' + esc_(valor) + '</div><div style="font-size:13px">' + esc_(t) + '</div></td>';
  const tabla = (filas, c1, c2) => '<table cellpadding="6" style="border-collapse:collapse;font-size:14px">' +
    '<tr style="background:#e8f0f2"><th align="left">' + c1 + '</th><th align="right">' + c2 + '</th></tr>' +
    filas.map(f => '<tr><td style="border-bottom:1px solid #eee">' + esc_(f[0]) + '</td><td align="right" style="border-bottom:1px solid #eee">' + esc_(f[1]) + '</td></tr>').join('') +
    '</table>';

  let html = '<div style="font-family:Arial,sans-serif;color:#222;max-width:640px">' +
    '<h2 style="color:#1f4e5f;margin:0 0 4px">' + esc_(tipo) + ' de facturas — ' + esc_(titulo) + '</h2>' +
    '<p style="margin:0 0 14px;color:#666">' + esc_(conf.NOMBRE_NEGOCIO) + (flecha ? ' · ' + esc_(flecha) : '') + '</p>' +
    '<table cellspacing="8"><tr>' +
    caja('solicitudes recibidas', st.recibidas, '#e8f0f2') + caja('enviadas', st.enviadas, '#d9ead3') +
    caja('rechazadas', st.rechazadas, '#f4cccc') + caja('siguen pendientes', st.pendientes, '#fff2cc') +
    '</tr></table>' +
    '<p><b>Tiempo medio hasta enviar la factura:</b> ' + esc_(fmtDias_(st.diasMedios)) + '<br>' +
    '<b>Pedidas por la web:</b> ' + st.web + ' de ' + st.recibidas + '<br>' +
    '<b>Clientes nuevos dados de alta:</b> ' + st.nuevos + '</p>';
  if (st.clientes.length) html += '<h3 style="color:#1f4e5f">Clientes que más facturas pidieron</h3>' + tabla(st.clientes.slice(0, 5), 'Cliente', 'Solicitudes');
  if (st.pagos.length) html += '<h3 style="color:#1f4e5f">Forma de pago</h3>' + tabla(st.pagos, 'Forma de pago', 'Solicitudes');
  html += '<h3 style="color:#1f4e5f">Pendientes con más de ' + DIAS_ATRASO_ + ' días (hoy)</h3>';
  html += atr.length
    ? tabla(atr.slice(0, 20).map(s => [s['ID'] + ' · ' + (s['Nombre o razón social'] || s['RFC']) + ' · ticket ' + s['Folio ticket'], fechaCorta_(s['Fecha solicitud'])]), 'Solicitud', 'Pedida el') +
      (atr.length > 20 ? '<p>… y ' + (atr.length - 20) + ' más.</p>' : '')
    : '<p>Ninguna. ¡Todo al día! 🎉</p>';
  html += '<p style="color:#666;font-size:13px;margin-top:20px">Más detalle en la hoja <i>Panel</i> del archivo de solicitudes. ' + esc_(nota || '') + '</p></div>';

  const para = String(conf.CORREOS_INFORME || conf.CORREOS_AVISO);
  enviarCorreo_(para, tipo + ' de facturas — ' + titulo, html);
  return { para: para, stats: st, atrasadas: atr.length };
}
