/**
 * AClímate — Google Apps Script Backend v3
 * Recibe los 6 formularios operativos, gestiona OTs e inventario.
 *
 * ── INSTRUCCIONES DE DESPLIEGUE ────────────────────────────────────────
 * 1. Abrí script.google.com → Nuevo proyecto → pegá este código.
 * 2. Cambiá SPREADSHEET_ID por el ID de tu Google Sheet.
 *    URL: https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit
 * 3. Las pestañas se crean automáticamente la primera vez que se usa cada form.
 * 4. Guardá el script (Ctrl+S).
 * 5. Menú Implementar → Nueva implementación → Tipo: Aplicación web.
 *    - Ejecutar como: Yo mismo
 *    - Quién tiene acceso: Cualquier persona
 * 6. Autorizá los permisos que Google solicita.
 * 7. Copiá el URL generado (termina en /exec) y pegalo en config.js → GAS_URL.
 * ──────────────────────────────────────────────────────────────────────
 */

// ─── CONFIGURACIÓN ─────────────────────────────────────────────────────
const SPREADSHEET_ID    = '1h1cNml-yMtj1lDk_XQYV3Qa5O9vBtGf-ks6OnhzK3Xo';
const DRIVE_FOLDER_NAME = 'AClímate — Fotos Operativas';
const EMAIL_DESTINATARIOS = 'aporras@avelectromecanica.com,aclimatecr@avelectromecanica.com,aalfaro@avelectromecanica.com';
// ───────────────────────────────────────────────────────────────────────

// ── GET: devuelve datos para los formularios ──────────────────────────
function doGet(e) {
  const action = e?.parameter?.action || '';

  if (action === 'ots') {
    try {
      const sheet = getSheet('OTs');
      const ots = leerOTs(sheet);
      return jsonOk(ots);
    } catch(err) {
      return jsonError(err.toString());
    }
  }

  if (action === 'inventario') {
    try {
      const sheet = getSheet('Inventario');
      if (sheet.getLastRow() <= 1) return jsonOk([]);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const items = data.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[String(h)] = row[i]; });
        return obj;
      });
      return jsonOk(items);
    } catch(err) {
      return jsonError(err.toString());
    }
  }

  if (action === 'dashboard') {
    try {
      return jsonOk(computarDashboard());
    } catch(err) {
      return jsonError(err.toString());
    }
  }

  if (action === 'movimientos') {
    try {
      const sheet = getSheet('Movimientos');
      if (sheet.getLastRow() <= 1) return jsonOk([]);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const limit = parseInt(e?.parameter?.limit || '300');
      const rows = data.slice(1).slice(-limit);
      const movs = rows.map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[String(h)] = row[i]; });
        return obj;
      });
      movs.reverse(); // newest first
      return jsonOk(movs);
    } catch(err) {
      return jsonError(err.toString());
    }
  }

  // Ruta raíz: devuelve estado del sistema
  return jsonOk({ sistema: 'AClimate Forms v3', estado: 'activo' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const formId = data.form;

    let resultado;
    switch (formId) {
      case 'F01': resultado = manejarF01(data); break;
      case 'F02': resultado = manejarF02(data); break;
      case 'F03': resultado = manejarF03(data); break;
      case 'F04': resultado = manejarF04(data); break;
      case 'F05': resultado = manejarF05(data); break;
      case 'F06': resultado = manejarF06(data); break;
      case 'OT_CREAR':              resultado = manejarOTCrear(data);             break;
      case 'OT_EDITAR':             resultado = manejarOTEditar(data);            break;
      case 'OT_CERRAR':             resultado = manejarOTCerrar(data);            break;
      case 'OT_ELIMINAR':          resultado = manejarOTEliminar(data);          break;
      case 'INVENTARIO_IMPORTAR':   resultado = manejarInventarioImportar(data);  break;
      case 'INVENTARIO_COMPRA':    resultado = manejarInventarioCompra(data);    break;
      default:    throw new Error('Formulario desconocido: ' + formId);
    }

    return jsonOk(resultado);
  } catch (err) {
    Logger.log('ERROR: ' + err.toString());
    return jsonError(err.toString());
  }
}

// ── Email helper ─────────────────────────────────────────────────────
function enviarEmail(asunto, lineas) {
  try {
    const cuerpo = lineas.join('\n');
    MailApp.sendEmail({
      to: EMAIL_DESTINATARIOS,
      subject: '[AClímate] ' + asunto,
      body: cuerpo + '\n\n—\nEnviado automáticamente por AClímate Forms'
    });
  } catch(e) {
    Logger.log('Email error: ' + e.toString());
  }
}

// ── OTs — Crear ─────────────────────────────────────────────────────
function manejarOTCrear(d) {
  const sheet = getSheet('OTs');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Cliente', 'Tipo', 'Técnico', 'Descripción', 'Fecha Asignación', 'Estado', 'Timestamp Creación', 'Timestamp Cierre', 'Contacto', 'Teléfono', 'Dirección', 'Presupuesto', 'Trabajos']);
    formatearEncabezados(sheet);
  } else {
    const hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    ['Contacto', 'Teléfono', 'Dirección', 'Presupuesto', 'Trabajos'].forEach(h => {
      if (!hdrs.includes(h)) { hdrs.push(h); sheet.getRange(1, hdrs.length).setValue(h); }
    });
    formatearEncabezados(sheet);
  }

  const otsActuales = leerOTs(sheet);
  if (otsActuales.find(o => o.id === d.id)) throw new Error('OT ' + d.id + ' ya existe.');

  const trabajos = Array.isArray(d.trabajos) ? d.trabajos : [];
  const tipoStr  = trabajos.length > 1
    ? trabajos.map(t => t.tipo).join(' / ')
    : trabajos.length === 1 ? trabajos[0].tipo : (d.tipo || '');
  const presStr  = trabajos.flatMap(t => (t.presupuesto || []).map(it => `${it.codigo} | ${it.nombre} | ${it.cantidad}`)).join('\n') || d.presupuesto_str || '';

  sheet.appendRow([
    d.id, d.cliente, tipoStr, d.tecnico, d.descripcion || '',
    d.fecha_asignacion, 'Abierta', d.timestamp, '',
    d.contacto_nombre || '', d.contacto_tel || '', d.direccion || '',
    presStr, trabajos.length ? JSON.stringify(trabajos) : ''
  ]);

  enviarEmail('Nueva OT creada: ' + d.id, [
    'OT:          ' + d.id,
    'Cliente:     ' + d.cliente,
    'Tipo:        ' + tipoStr,
    'Técnico:     ' + d.tecnico,
    'Fecha:       ' + d.fecha_asignacion,
    'Descripción: ' + (d.descripcion || '—'),
    trabajos.length > 1 ? ('Trabajos:    ' + trabajos.length) : '',
  ].filter(Boolean));

  return { id: d.id };
}

// ── OTs — Editar ─────────────────────────────────────────────────────
function manejarOTEditar(d) {
  const sheet = getSheet('OTs');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = h => headers.indexOf(h) + 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === d.id) {
      const trabajos = Array.isArray(d.trabajos) ? d.trabajos : [];
      const tipoStr  = trabajos.length > 1
        ? trabajos.map(t => t.tipo).join(' / ')
        : trabajos.length === 1 ? trabajos[0].tipo : (d.tipo || '');
      const presStr  = trabajos.flatMap(t => (t.presupuesto || []).map(it => `${it.codigo} | ${it.nombre} | ${it.cantidad}`)).join('\n') || d.presupuesto_str || '';

      sheet.getRange(i + 1, col('Cliente')).setValue(d.cliente || '');
      sheet.getRange(i + 1, col('Tipo')).setValue(tipoStr);
      sheet.getRange(i + 1, col('Técnico')).setValue(d.tecnico || '');
      sheet.getRange(i + 1, col('Descripción')).setValue(d.descripcion || '');
      sheet.getRange(i + 1, col('Fecha Asignación')).setValue(d.fecha_asignacion || '');
      if (col('Contacto')    > 0) sheet.getRange(i + 1, col('Contacto')).setValue(d.contacto_nombre || '');
      if (col('Teléfono')    > 0) sheet.getRange(i + 1, col('Teléfono')).setValue(d.contacto_tel    || '');
      if (col('Dirección')   > 0) sheet.getRange(i + 1, col('Dirección')).setValue(d.direccion      || '');
      if (col('Presupuesto') > 0) sheet.getRange(i + 1, col('Presupuesto')).setValue(presStr);
      if (col('Trabajos')    > 0) sheet.getRange(i + 1, col('Trabajos')).setValue(trabajos.length ? JSON.stringify(trabajos) : '');

      enviarEmail('OT editada: ' + d.id, [
        'OT:      ' + d.id,
        'Cliente: ' + (d.cliente || ''),
        'Tipo:    ' + tipoStr,
        'Técnico: ' + (d.tecnico || ''),
        d.direccion ? ('Dirección: ' + d.direccion) : '',
      ].filter(l => l !== ''));

      return { id: d.id, editada: true };
    }
  }
  throw new Error('OT ' + d.id + ' no encontrada para editar.');
}

// ── OTs — Cerrar ─────────────────────────────────────────────────────
function manejarOTCerrar(d) {
  const sheet = getSheet('OTs');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === d.id) {
      sheet.getRange(i + 1, 7).setValue('Cerrada');
      sheet.getRange(i + 1, 9).setValue(Utilities.formatDate(new Date(), 'America/Costa_Rica', 'dd/MM/yyyy hh:mm a'));

      enviarEmail('OT cerrada: ' + d.id, [
        'OT:      ' + d.id,
        'Estado:  Cerrada',
        'Cierre:  ' + Utilities.formatDate(new Date(), 'America/Costa_Rica', 'dd/MM/yyyy hh:mm a'),
      ]);

      return { id: d.id, estado: 'Cerrada' };
    }
  }
  throw new Error('OT ' + d.id + ' no encontrada.');
}

// ── Leer OTs desde sheet ─────────────────────────────────────────────
function leerOTs(sheet) {
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[String(h).toLowerCase().replace(/\s+/g, '_')] = row[i]; });
    if (typeof obj.trabajos === 'string' && obj.trabajos.trim()) {
      try { obj.trabajos = JSON.parse(obj.trabajos); } catch(e) { obj.trabajos = []; }
    } else {
      obj.trabajos = Array.isArray(obj.trabajos) ? obj.trabajos : [];
    }
    return obj;
  });
}

// ── OTs — Eliminar ───────────────────────────────────────────────────
function manejarOTEliminar(d) {
  const sheet = getSheet('OTs');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === d.id) {
      sheet.deleteRow(i + 1);
      enviarEmail('OT eliminada: ' + d.id, [
        'La OT ' + d.id + ' fue eliminada permanentemente por el PM.',
        'Timestamp: ' + Utilities.formatDate(new Date(), 'America/Costa_Rica', 'dd/MM/yyyy hh:mm a'),
      ]);
      return { id: d.id, eliminada: true };
    }
  }
  throw new Error('OT ' + d.id + ' no encontrada para eliminar.');
}

// ── F-01 Despacho de Materiales ─────────────────────────────────────
function manejarF01(d) {
  const sheet = getSheet('F01-Despacho');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Fecha', 'Técnico', 'OT', 'Cliente',
      'Ítems (código|nombre|cantidad)', 'Observaciones', 'Firma', 'Fotos'
    ]);
    formatearEncabezados(sheet);
  }

  const itemsStr = (d.items || [])
    .map(it => `${it.codigo} | ${it.nombre} | ${it.cantidad}`)
    .join('\n');

  const fotos = guardarFotos(d.fotos, `F01_${d.ot}_${d.fecha}`);

  sheet.appendRow([
    d.timestamp, d.fecha, d.tecnico, d.ot, d.cliente,
    itemsStr, d.observaciones || '', d.firma, fotos
  ]);

  // Registrar salida en inventario
  registrarMovimientos(d.items, 'Salida', d.ot, d.tecnico, d.fecha, 'F01');

  enviarEmail('F-01 Despacho — ' + d.ot + ' | ' + d.tecnico, [
    'Fecha:    ' + d.fecha,
    'Técnico:  ' + d.tecnico,
    'OT:       ' + d.ot,
    'Cliente:  ' + d.cliente,
    '',
    'Ítems despachados:',
    itemsStr || '—',
    '',
    'Observaciones: ' + (d.observaciones || '—'),
  ]);

  return { form: 'F01', ot: d.ot };
}

// ── F-02 Checklist Inicio ───────────────────────────────────────────
function manejarF02(d) {
  const sheet = getSheet('F02-ChecklistInicio');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Fecha', 'Técnico',
      'Combustible', 'Limpieza Interior', 'Limpieza Exterior',
      'Manifold', 'Bomba Vacío', 'Balanza', 'Detector Fugas',
      'Multímetro', 'Pinzas', 'Taladro', 'Escalera',
      'EPP Técnico', 'EPP Asistente', 'Documentos Camión',
      'Materiales Verificados', 'Problema', 'Fotos'
    ]);
    formatearEncabezados(sheet);
  }

  const fotos = guardarFotos(d.fotos, `F02_${d.fecha}`);

  sheet.appendRow([
    d.timestamp, d.fecha, d.tecnico,
    d.combustible, d.limpieza_interior, d.limpieza_exterior,
    d.manifold, d.bomba_vacio, d.balanza, d.detector_fugas,
    d.multimetro, d.pinzas, d.taladro, d.escalera,
    d.epp_tecnico, d.epp_asistente, d.documentos_camion,
    d.materiales_verificados, d.problema || '', fotos
  ]);

  const problemasF02 = d.problema ? ('Problema: ' + d.problema) : 'Sin problemas';
  enviarEmail('F-02 Checklist Inicio — ' + d.fecha + ' | ' + d.tecnico, [
    'Fecha:     ' + d.fecha,
    'Técnico:   ' + d.tecnico,
    '',
    'Combustible:          ' + d.combustible,
    'EPP técnico:          ' + d.epp_tecnico,
    'EPP asistente:        ' + d.epp_asistente,
    'Materiales verif.:    ' + d.materiales_verificados,
    '',
    problemasF02,
  ]);

  return { form: 'F02' };
}

// ── F-03 Reporte de Trabajo ─────────────────────────────────────────
function manejarF03(d) {
  const sheet = getSheet('F03-ReporteTrabajo');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Fecha', 'OT', 'Técnico', 'Cliente', 'Tipo Trabajo',
      'Hora Llegada', 'Hora Inicio', 'Hora Fin',
      'Ítems Utilizados', 'Descripción Trabajo',
      'Hubo Problema', 'Detalle Problema',
      'Estado Equipo', 'VoBo', 'VoBo Justif',
      'Observaciones', 'Fotos Trabajo', 'Foto VoBo'
    ]);
    formatearEncabezados(sheet);
  }

  const itemsStr = (d.items || [])
    .map(it => `${it.codigo} | ${it.nombre} | ${it.cantidad}`)
    .join('\n');

  const fotosT  = guardarFotos(d.fotos_trabajo, `F03_${d.ot}_${d.fecha}_trabajo`);
  const fotosV  = guardarFotos(d.fotos_vobo,    `F03_${d.ot}_${d.fecha}_vobo`);

  sheet.appendRow([
    d.timestamp, d.fecha, d.ot, d.tecnico, d.cliente, d.tipo_trabajo,
    d.hora_llegada, d.hora_inicio, d.hora_fin,
    itemsStr, d.descripcion,
    d.hubo_problema, d.problema_detalle || '',
    d.estado_equipo, d.vobo, d.vobo_justificacion || '',
    d.observaciones || '', fotosT, fotosV
  ]);

  // Registrar consumo (informativo, sin cambio de stock)
  registrarMovimientos(d.items, 'Consumo', d.ot, d.tecnico, d.fecha, 'F03');

  const problemasF03 = d.hubo_problema === 'Sí' ? ('Problema: ' + (d.problema_detalle || 'Sin detalle')) : 'Sin problemas';
  enviarEmail('F-03 Reporte Trabajo — ' + d.ot + ' | ' + d.tecnico, [
    'Fecha:        ' + d.fecha,
    'OT:           ' + d.ot,
    'Cliente:      ' + d.cliente,
    'Técnico:      ' + d.tecnico,
    'Tipo trabajo: ' + d.tipo_trabajo,
    '',
    'Llegada: ' + d.hora_llegada + '  |  Inicio: ' + d.hora_inicio + '  |  Fin: ' + d.hora_fin,
    '',
    'Descripción del trabajo:',
    d.descripcion || '—',
    '',
    'Ítems utilizados:',
    itemsStr || '—',
    '',
    'Estado del equipo: ' + d.estado_equipo,
    'VoBo cliente:      ' + d.vobo,
    problemasF03,
  ]);

  return { form: 'F03', ot: d.ot };
}

// ── F-04 Devolución ─────────────────────────────────────────────────
function manejarF04(d) {
  const sheet = getSheet('F04-Devolucion');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Fecha', 'OT Origen', 'Técnico',
      'Ítems Devueltos', 'Estado Ítems',
      'Material en Camión', 'Detalle Camión',
      'Firma', 'Fotos'
    ]);
    formatearEncabezados(sheet);
  }

  const itemsStr = (d.items || [])
    .map(it => `${it.codigo} | ${it.nombre} | ${it.cantidad}`)
    .join('\n');

  const fotos = guardarFotos(d.fotos, `F04_${d.ot}_${d.fecha}`);

  sheet.appendRow([
    d.timestamp, d.fecha, d.ot, d.tecnico,
    itemsStr, d.estado_items,
    d.material_en_camion, d.material_camion_detalle || '',
    d.firma, fotos
  ]);

  // Registrar devolución: suma stock de vuelta al inventario
  registrarMovimientos(d.items, 'Devolución', d.ot, d.tecnico, d.fecha, 'F04');

  enviarEmail('F-04 Devolución — ' + d.ot + ' | ' + d.tecnico, [
    'Fecha:    ' + d.fecha,
    'OT:       ' + d.ot,
    'Técnico:  ' + d.tecnico,
    '',
    'Ítems devueltos:',
    itemsStr || '—',
    '',
    'Estado ítems:       ' + d.estado_items,
    'Material en camión: ' + d.material_en_camion,
    d.material_camion_detalle ? ('Detalle camión: ' + d.material_camion_detalle) : '',
  ].filter(l => l !== ''));

  return { form: 'F04', ot: d.ot };
}

// ── F-05 Checklist Cierre ───────────────────────────────────────────
function manejarF05(d) {
  const sheet = getSheet('F05-ChecklistCierre');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Fecha', 'Técnico',
      'Sobrantes Devueltos', 'Justif. Sobrantes',
      'Herramientas', 'Detalle Herramientas',
      'Estado Camión', 'Combustible Cierre',
      'Incidente', 'Detalle Incidente',
      'Requiere Mantenimiento', 'Detalle Mantenimiento',
      'Resumen Día', 'Fotos'
    ]);
    formatearEncabezados(sheet);
  }

  const fotos = guardarFotos(d.fotos, `F05_${d.fecha}`);

  sheet.appendRow([
    d.timestamp, d.fecha, d.tecnico,
    d.sobrantes_devueltos, d.sobrantes_justificacion || '',
    d.herramientas, d.herramientas_detalle || '',
    d.estado_camion, d.combustible_cierre,
    d.incidente_camion, d.incidente_detalle || '',
    d.requiere_mantenimiento, d.mantenimiento_detalle || '',
    d.resumen_dia, fotos
  ]);

  const incidenteF05 = d.incidente_camion === 'Sí' ? ('Incidente: ' + (d.incidente_detalle || 'Sin detalle')) : 'Sin incidentes';
  const mantoF05 = d.requiere_mantenimiento === 'Sí' ? ('Mantenimiento: ' + (d.mantenimiento_detalle || 'Sin detalle')) : 'No requiere mantenimiento';
  enviarEmail('F-05 Checklist Cierre — ' + d.fecha + ' | ' + d.tecnico, [
    'Fecha:              ' + d.fecha,
    'Técnico:            ' + d.tecnico,
    '',
    'Estado camión:      ' + d.estado_camion,
    'Combustible cierre: ' + d.combustible_cierre,
    'Sobrantes devueltos:' + d.sobrantes_devueltos,
    'Herramientas OK:    ' + d.herramientas,
    '',
    incidenteF05,
    mantoF05,
    '',
    'Resumen del día:',
    d.resumen_dia || '—',
  ]);

  return { form: 'F05' };
}

// ── F-06 Viáticos ───────────────────────────────────────────────────
function manejarF06(d) {
  const sheet = getSheet('F06-Viaticos');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Fecha', 'Técnico', 'OT',
      'Tipo Gasto', 'Descripción', 'Monto (₡)',
      'Aprobado PM', 'Justif. Sin Aprobación',
      'Hay Vuelto', 'Monto Vuelto (₡)',
      'Observaciones', 'Foto Recibo'
    ]);
    formatearEncabezados(sheet);
  }

  const fotoR = guardarFotos(d.foto_recibo ? [d.foto_recibo] : [], `F06_${d.ot}_${d.fecha}`);

  sheet.appendRow([
    d.timestamp, d.fecha, d.tecnico, d.ot,
    d.tipo_gasto, d.descripcion, d.monto,
    d.aprobado_pm, d.justificacion_no_aprobado || '',
    d.hay_vuelto, d.monto_vuelto || '',
    d.observaciones || '', fotoR
  ]);

  const aprobF06 = d.aprobado_pm === 'No' ? ('Sin aprobación PM: ' + (d.justificacion_no_aprobado || '—')) : 'Aprobado por PM';
  enviarEmail('F-06 Viático — ₡' + d.monto + ' | ' + d.tecnico, [
    'Fecha:      ' + d.fecha,
    'Técnico:    ' + d.tecnico,
    'OT:         ' + d.ot,
    '',
    'Tipo gasto: ' + d.tipo_gasto,
    'Descripción:' + d.descripcion,
    'Monto:      ₡' + d.monto,
    aprobF06,
    d.hay_vuelto === 'Sí' ? ('Vuelto: ₡' + (d.monto_vuelto || '0')) : '',
    '',
    'Observaciones: ' + (d.observaciones || '—'),
  ].filter(l => l !== ''));

  return { form: 'F06' };
}

// ── Dashboard — Cálculo de KPIs ──────────────────────────────────────
function computarDashboard() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ahora = new Date();
  const ts  = Utilities.formatDate(ahora, 'America/Costa_Rica', 'dd/MM/yyyy hh:mm a');

  function sheetRows(nombre) {
    const sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() <= 1) return [];
    const data = sh.getDataRange().getValues();
    const h = data[0];
    return data.slice(1).map(r => {
      const o = {}; h.forEach((k, i) => { o[String(k)] = r[i]; }); return o;
    });
  }
  function sheetCount(nombre) {
    const sh = ss.getSheetByName(nombre);
    return sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  }

  // ── OTs ──
  const ots = sheetRows('OTs');
  const otsAbiertas  = ots.filter(o => o['Estado'] === 'Abierta');
  const otsCerradas  = ots.filter(o => o['Estado'] === 'Cerrada');
  const tiposOT      = {};
  ots.forEach(o => { const t = o['Tipo'] || '—'; tiposOT[t] = (tiposOT[t]||0)+1; });

  // ── Inventario ──
  const inv = sheetRows('Inventario');
  const invConStock  = inv.filter(i => Number(i['Existencia']) > 0).length;
  const invSinStock  = inv.filter(i => Number(i['Existencia']) <= 0).length;
  const invStockBajo = inv.filter(i => {
    const e = Number(i['Existencia']), m = Number(i['Stock Mín']);
    return e > 0 && m > 0 && e < m;
  }).length;
  const valorTotal   = inv.reduce((s, i) => s + (Number(i['Existencia'])||0) * (Number(i['Costo Unit (₡)'])||0), 0);
  const invCamionOK  = inv.filter(i => i['Ubicación']==='CAMIÓN' && i['Estado']==='OK').length;
  const invCamionDan = inv.filter(i => i['Ubicación']==='CAMIÓN' && i['Estado'] && i['Estado']!=='OK').length;

  // ── Movimientos (últimos 30 días) ──
  const movs = sheetRows('Movimientos');
  const hace30d = new Date(ahora.getTime() - 30*24*60*60*1000);
  const parseFecha = f => { try { const [d,m,y]=String(f).split('/'); return new Date(y,m-1,d); } catch(e){return new Date(0);} };
  const movs30   = movs.filter(m => parseFecha(m['Fecha']) >= hace30d);
  const salidas30  = movs30.filter(m => m['Tipo']==='Salida').length;
  const devol30    = movs30.filter(m => m['Tipo']==='Devolución').length;
  const compras30  = movs30.filter(m => m['Tipo']==='Compra').length;
  const consumo30  = movs30.filter(m => m['Tipo']==='Consumo').length;

  // ── F06 Viáticos ──
  const viat = sheetRows('F06-Viaticos');
  const totalViat   = viat.reduce((s, v) => s + (Number(v['Monto (₡)'])||0), 0);
  const sinAprob    = viat.filter(v => v['Aprobado PM']==='No').length;
  const totalVuelto = viat.reduce((s, v) => s + (Number(v['Monto Vuelto (₡)'])||0), 0);

  // ── F03 Incidencias de trabajo ──
  const f03 = sheetRows('F03-ReporteTrabajo');
  const problemas  = f03.filter(r => r['Hubo Problema']==='Sí').length;
  const voboNeg    = f03.filter(r => r['VoBo']==='No').length;

  // ── F05 Incidencias camión ──
  const f05 = sheetRows('F05-ChecklistCierre');
  const incCamion  = f05.filter(r => r['Incidente']==='Sí').length;
  const mantoReq   = f05.filter(r => r['Requiere Mantenimiento']==='Sí').length;

  // ── Formularios enviados (total acumulado) ──
  const forms = {
    'F01-Despacho':       sheetCount('F01-Despacho'),
    'F02-Inicio':         sheetCount('F02-ChecklistInicio'),
    'F03-Trabajo':        sheetCount('F03-ReporteTrabajo'),
    'F04-Devolución':     sheetCount('F04-Devolucion'),
    'F05-Cierre':         sheetCount('F05-ChecklistCierre'),
    'F06-Viáticos':       sheetCount('F06-Viaticos'),
  };
  const totalForms = Object.values(forms).reduce((s, v) => s + v, 0);

  // ── Control OT: diferencia F01 vs F03+F04 ──
  const otsConDif = [];
  if (movs.length) {
    const porOT = {};
    movs.forEach(m => {
      const ot = m['OT'] || '—';
      if (ot === '—') return;
      if (!porOT[ot]) porOT[ot] = { salida: 0, consumo: 0, devolucion: 0 };
      const qty = Number(m['Cantidad']) || 0;
      if (m['Tipo'] === 'Salida')     porOT[ot].salida     += qty;
      if (m['Tipo'] === 'Consumo')    porOT[ot].consumo    += qty;
      if (m['Tipo'] === 'Devolución') porOT[ot].devolucion += qty;
    });
    Object.entries(porOT).forEach(([ot, v]) => {
      const dif = v.salida - v.consumo - v.devolucion;
      if (Math.abs(dif) > 0) otsConDif.push({ ot, salida: v.salida, consumo: v.consumo, devolucion: v.devolucion, diferencia: dif });
    });
  }

  return {
    timestamp: ts,
    ots: {
      total: ots.length, abiertas: otsAbiertas.length, cerradas: otsCerradas.length,
      tipos: tiposOT,
      recientes: otsAbiertas.slice(0, 8).map(o => ({
        id: o['ID'] || o['id'] || '', cliente: o['Cliente'] || o['cliente'] || '',
        tipo: o['Tipo'] || o['tipo'] || '', tecnico: o['Técnico'] || o['técnico'] || '',
        fecha: o['Fecha Asignación'] || o['fecha_asignación'] || ''
      }))
    },
    inventario: {
      total: inv.length, con_stock: invConStock, sin_stock: invSinStock, stock_bajo: invStockBajo,
      valor_estimado: Math.round(valorTotal),
      camion_ok: invCamionOK, camion_dañado: invCamionDan
    },
    movimientos: { salidas_30d: salidas30, devoluciones_30d: devol30, compras_30d: compras30, consumo_30d: consumo30 },
    viaticos: { total_acumulado: totalViat, sin_aprobacion: sinAprob, vuelto_total: totalVuelto, registros: viat.length },
    incidencias: { trabajo: problemas, camion: incCamion, vobo_negativo: voboNeg, mantenimiento_pendiente: mantoReq },
    formularios: { detalle: forms, total: totalForms },
    control_ot: { ots_con_diferencia: otsConDif.length, detalle: otsConDif.slice(0, 10) }
  };
}

// ── Inventario — Importar (seed inicial desde Excel Maestro) ─────────
function manejarInventarioImportar(d) {
  const sheet = getSheet('Inventario');
  const INV_HEADERS = [
    'Código', 'Ubicación', 'Categoría', 'Descripción', 'Marca',
    'Proveedor', 'Ubic. Camión', 'Stock Mín', 'Stock Máx',
    'Existencia', 'Costo Unit (₡)', 'CTD Inicial', 'Estado', 'Última Actualización'
  ];

  // Crear encabezados si la hoja está vacía
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(INV_HEADERS);
    formatearEncabezados(sheet);
  }

  const items = d.items || [];
  if (!items.length) return { insertados: 0, actualizados: 0 };

  // Leer mapa existente para decidir insertar vs actualizar
  let existingData = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [INV_HEADERS];
  const codigoIdx = existingData[0].indexOf('Código');
  const existMap = {};
  for (let i = 1; i < existingData.length; i++) {
    existMap[String(existingData[i][codigoIdx])] = i + 1; // 1-indexed row
  }

  const ts = Utilities.formatDate(new Date(), 'America/Costa_Rica', 'dd/MM/yyyy hh:mm a');
  let insertados = 0, actualizados = 0;

  items.forEach(it => {
    const row = [
      it.codigo, it.ubicacion, it.categoria, it.descripcion,
      it.marca || '', it.proveedor || '', it.ubicacion_camion || '',
      Number(it.stock_min) || 0, Number(it.stock_max) || 0,
      Number(it.existencia) || 0, Number(it.costo_unit) || 0,
      Number(it.ctd_inicial) || 0, it.estado || '', ts
    ];
    if (existMap[it.codigo]) {
      sheet.getRange(existMap[it.codigo], 1, 1, row.length).setValues([row]);
      actualizados++;
    } else {
      sheet.appendRow(row);
      insertados++;
    }
  });

  return { insertados, actualizados, total: items.length };
}

// ── Inventario — Compra / ingreso ────────────────────────────────────
function manejarInventarioCompra(d) {
  const sheetInv = getSheet('Inventario');
  const sheetMov = getSheet('Movimientos');

  if (sheetMov.getLastRow() === 0) {
    sheetMov.appendRow([
      'Timestamp', 'Fecha', 'OT', 'Técnico', 'Código', 'Descripción',
      'Tipo', 'Cantidad', 'Stock Ant.', 'Stock Nuevo', 'Formulario', 'Observación'
    ]);
    formatearEncabezados(sheetMov);
  }

  const qty = Number(d.cantidad) || 0;
  if (qty <= 0) throw new Error('Cantidad debe ser mayor a 0');
  if (!d.codigo) throw new Error('Código de ítem requerido');

  const ts = Utilities.formatDate(new Date(), 'America/Costa_Rica', 'dd/MM/yyyy hh:mm a');

  // Buscar el ítem en Inventario y actualizar existencia
  let stockAnt = 0, stockNuevo = qty, rowNum = -1;
  if (sheetInv.getLastRow() > 1) {
    const invData    = sheetInv.getDataRange().getValues();
    const invHeaders = invData[0];
    const codigoCol  = invHeaders.indexOf('Código');
    const existCol   = invHeaders.indexOf('Existencia');

    for (let i = 1; i < invData.length; i++) {
      if (String(invData[i][codigoCol]) === d.codigo) {
        stockAnt   = Number(invData[i][existCol]) || 0;
        stockNuevo = stockAnt + qty;
        rowNum     = i + 1;
        sheetInv.getRange(rowNum, existCol + 1).setValue(stockNuevo);
        // Actualizar última actualización si existe la columna
        const ultimaCol = invHeaders.indexOf('Última Actualización');
        if (ultimaCol >= 0) sheetInv.getRange(rowNum, ultimaCol + 1).setValue(ts);
        break;
      }
    }
  }

  // Registrar en Movimientos
  const obs = [d.proveedor, d.factura ? ('Factura: ' + d.factura) : '', d.observaciones]
    .filter(x => x).join(' · ');

  sheetMov.appendRow([
    ts, d.fecha, '—', d.registrado_por || 'PM',
    d.codigo, d.descripcion || '',
    'Compra', qty, stockAnt, stockNuevo,
    'COMPRA', obs
  ]);

  enviarEmail('Ingreso de compra — ' + d.codigo, [
    'Fecha:       ' + d.fecha,
    'Registrado:  ' + (d.registrado_por || 'PM'),
    'Ítem:        ' + d.codigo + ' — ' + (d.descripcion || ''),
    'Cantidad:    +' + qty,
    'Stock ant.:  ' + stockAnt,
    'Stock nuevo: ' + stockNuevo,
    d.proveedor   ? ('Proveedor:   ' + d.proveedor)  : '',
    d.factura     ? ('Factura:     ' + d.factura)     : '',
    d.observaciones ? ('Obs:         ' + d.observaciones) : '',
  ].filter(l => l !== ''));

  return { codigo: d.codigo, cantidad: qty, stock_anterior: stockAnt, nueva_existencia: stockNuevo };
}

// ── Inventario — Registrar movimientos y actualizar stock ────────────
function registrarMovimientos(items, tipo, ot, tecnico, fecha, formId) {
  if (!items || !items.length) return;

  const sheetInv = getSheet('Inventario');
  const sheetMov = getSheet('Movimientos');

  if (sheetMov.getLastRow() === 0) {
    sheetMov.appendRow([
      'Timestamp', 'Fecha', 'OT', 'Técnico', 'Código', 'Descripción',
      'Tipo', 'Cantidad', 'Stock Ant.', 'Stock Nuevo', 'Formulario', 'Observación'
    ]);
    formatearEncabezados(sheetMov);
  }

  const ts = Utilities.formatDate(new Date(), 'America/Costa_Rica', 'dd/MM/yyyy hh:mm a');

  // Construir mapa de inventario: codigo → {rowNum, existencia, existColIdx}
  const invData    = sheetInv.getLastRow() > 1 ? sheetInv.getDataRange().getValues() : null;
  const invHeaders = invData ? invData[0] : [];
  const codigoCol  = invHeaders.indexOf('Código');
  const existCol   = invHeaders.indexOf('Existencia'); // 0-indexed

  const invMap = {};
  if (invData) {
    for (let i = 1; i < invData.length; i++) {
      const cod = String(invData[i][codigoCol]);
      invMap[cod] = { rowNum: i + 1, existencia: Number(invData[i][existCol]) || 0 };
    }
  }

  items.forEach(it => {
    const qty = Number(it.cantidad) || 0;
    let stockAnt = '—', stockNuevo = '—';

    const inv = invMap[it.codigo];
    // Solo Salida (F01) baja stock; Devolución (F04) sube; Consumo (F03) es informativo
    if (inv && existCol >= 0 && (tipo === 'Salida' || tipo === 'Devolución')) {
      stockAnt = inv.existencia;
      const delta = tipo === 'Salida' ? -qty : qty;
      stockNuevo = Math.max(0, inv.existencia + delta);
      sheetInv.getRange(inv.rowNum, existCol + 1).setValue(stockNuevo); // 1-indexed col
      inv.existencia = stockNuevo; // update local map in case same item appears twice
    }

    sheetMov.appendRow([
      ts, fecha, ot || '—', tecnico,
      it.codigo, it.nombre || it.descripcion || '',
      tipo, qty, stockAnt, stockNuevo,
      formId, ''
    ]);
  });
}

// ── Helpers ─────────────────────────────────────────────────────────

function getSheet(nombre) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) sheet = ss.insertSheet(nombre);
  return sheet;
}

function formatearEncabezados(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground('#2C3E50');
  range.setFontColor('#FFFFFF');
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function getDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function guardarFotos(fotos, prefix) {
  if (!fotos || !fotos.length) return '';
  const folder = getDriveFolder();
  const urls = [];

  fotos.forEach((foto, i) => {
    try {
      const base64 = typeof foto === 'string' ? foto : foto.data;
      const nombre = (typeof foto === 'object' && foto.nombre)
        ? foto.nombre
        : `${prefix}_${i + 1}.jpg`;

      const clean = base64.replace(/^data:image\/\w+;base64,/, '');
      const decoded = Utilities.base64Decode(clean);
      const blob = Utilities.newBlob(decoded, 'image/jpeg', nombre);

      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls.push(file.getUrl());
    } catch (e) {
      Logger.log('Error guardando foto: ' + e.toString());
    }
  });

  return urls.join('\n');
}

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Test helper (ejecutar desde editor para autorizar MailApp) ───────
function testEmail() {
  MailApp.sendEmail({
    to: EMAIL_DESTINATARIOS,
    subject: '[AClímate] Test de configuración',
    body: 'Si recibís este email, la configuración de notificaciones está correcta.'
  });
}

// ── Limpieza de datos de prueba (ejecutar UNA VEZ antes de producción) ─
// Borra todos los registros operativos manteniendo cabeceras e inventario.
function limpiarDatosPrueba() {
  const HOJAS_A_LIMPIAR = [
    'OTs',
    'F01-Despacho',
    'F02-ChecklistInicio',
    'F03-ReporteTrabajo',
    'F04-Devolucion',
    'F05-ChecklistCierre',
    'F06-Viaticos',
    'Movimientos'
  ];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resumen = [];

  HOJAS_A_LIMPIAR.forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) { resumen.push(`${nombre}: no existe (ok)`); return; }
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) { resumen.push(`${nombre}: ya vacía`); return; }
    sh.deleteRows(2, lastRow - 1);
    resumen.push(`${nombre}: ${lastRow - 1} fila(s) eliminada(s)`);
  });

  Logger.log('── Limpieza completada ──\n' + resumen.join('\n'));
  return resumen;
}
