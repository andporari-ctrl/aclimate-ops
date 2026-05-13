/**
 * AClímate — Google Apps Script Backend v2
 * Recibe los 6 formularios operativos y gestiona OTs.
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

// ── GET: devuelve OTs abiertas para los formularios ──────────────────
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

  // Ruta raíz: devuelve estado del sistema
  return jsonOk({ sistema: 'AClimate Forms v2', estado: 'activo' });
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
      case 'OT_CREAR':  resultado = manejarOTCrear(data);  break;
      case 'OT_EDITAR': resultado = manejarOTEditar(data); break;
      case 'OT_CERRAR': resultado = manejarOTCerrar(data); break;
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
    sheet.appendRow(['ID', 'Cliente', 'Tipo', 'Técnico', 'Descripción', 'Fecha Asignación', 'Estado', 'Timestamp Creación', 'Timestamp Cierre', 'Contacto', 'Teléfono', 'Dirección', 'Presupuesto']);
    formatearEncabezados(sheet);
  } else if (sheet.getLastColumn() < 13) {
    // Migrar hoja existente: agregar columnas nuevas
    const encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nuevos = ['Contacto', 'Teléfono', 'Dirección', 'Presupuesto'];
    nuevos.forEach((h, i) => {
      if (!encabezados.includes(h)) {
        sheet.getRange(1, encabezados.length + i + 1).setValue(h);
      }
    });
    formatearEncabezados(sheet);
  }

  const existentes = leerOTs(sheet);
  if (existentes.find(o => o.id === d.id)) throw new Error('OT ' + d.id + ' ya existe.');

  sheet.appendRow([
    d.id, d.cliente, d.tipo, d.tecnico, d.descripcion || '',
    d.fecha_asignacion, 'Abierta', d.timestamp, '',
    d.contacto_nombre || '', d.contacto_tel || '', d.direccion || '', d.presupuesto_str || ''
  ]);

  enviarEmail('Nueva OT creada: ' + d.id, [
    'OT:       ' + d.id,
    'Cliente:  ' + d.cliente,
    'Tipo:     ' + d.tipo,
    'Técnico:  ' + d.tecnico,
    'Fecha:    ' + d.fecha_asignacion,
    'Descripción: ' + (d.descripcion || '—'),
  ]);

  return { id: d.id };
}

// ── OTs — Editar ─────────────────────────────────────────────────────
function manejarOTEditar(d) {
  const sheet = getSheet('OTs');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  const col = h => headers.indexOf(h) + 1; // 1-indexed, 0 si no existe

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === d.id) {
      sheet.getRange(i + 1, col('Cliente')).setValue(d.cliente || '');
      sheet.getRange(i + 1, col('Tipo')).setValue(d.tipo || '');
      sheet.getRange(i + 1, col('Técnico')).setValue(d.tecnico || '');
      sheet.getRange(i + 1, col('Descripción')).setValue(d.descripcion || '');
      sheet.getRange(i + 1, col('Fecha Asignación')).setValue(d.fecha_asignacion || '');
      if (col('Contacto')   > 0) sheet.getRange(i + 1, col('Contacto')).setValue(d.contacto_nombre || '');
      if (col('Teléfono')   > 0) sheet.getRange(i + 1, col('Teléfono')).setValue(d.contacto_tel    || '');
      if (col('Dirección')  > 0) sheet.getRange(i + 1, col('Dirección')).setValue(d.direccion      || '');
      if (col('Presupuesto') > 0) sheet.getRange(i + 1, col('Presupuesto')).setValue(d.presupuesto_str || '');

      enviarEmail('OT editada: ' + d.id, [
        'OT:      ' + d.id,
        'Cliente: ' + (d.cliente || ''),
        'Tipo:    ' + (d.tipo    || ''),
        'Técnico: ' + (d.tecnico || ''),
        d.direccion ? ('📍 ' + d.direccion) : '',
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
        'Cierre:  ' + new Date().toLocaleString('es-CR'),
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
    return obj;
  });
}

// ── F-01 Despacho de Materiales ─────────────────────────────────────
function manejarF01(d) {
  const sheet = getSheet('F01-Despacho');

  // Encabezados si la hoja está vacía
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

  const problemasF02 = d.problema ? ('⚠️ Problema: ' + d.problema) : 'Sin problemas';
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

  const problemasF03 = d.hubo_problema === 'Sí' ? ('⚠️ ' + (d.problema_detalle || 'Sin detalle')) : 'Sin problemas';
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

  const incidenteF05 = d.incidente_camion === 'Sí' ? ('⚠️ Incidente: ' + (d.incidente_detalle || 'Sin detalle')) : 'Sin incidentes';
  const mantoF05 = d.requiere_mantenimiento === 'Sí' ? ('🔧 Mantenimiento: ' + (d.mantenimiento_detalle || 'Sin detalle')) : 'No requiere mantenimiento';
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

  const aprobF06 = d.aprobado_pm === 'No' ? ('⚠️ Sin aprobación PM: ' + (d.justificacion_no_aprobado || '—')) : 'Aprobado por PM';
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

/**
 * Guarda fotos en Google Drive y devuelve los URLs separados por coma.
 * @param {Array} fotos - Array de {nombre, data} o strings base64
 * @param {string} prefix - Prefijo del nombre de archivo
 */
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

      // Decodificar base64 (remover prefijo data:image/...;base64,)
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
