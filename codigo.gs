// ═══════════════════════════════════════════════════════════════
// ERP POS LITE — BACKEND SIMPLIFICADO
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = "1_UIN1IB5SdDqUmN8DgGnZYxFa-HSsSuoYOaKoCJ7T6w"; // ACTUALIZADO

const H_CATEGORIAS   = "Categorias";
const H_PRODUCTOS    = "Productos";
const H_VENTAS       = "Ventas";
const H_VENTA_DET    = "VentaDetalle";
const H_PAPELERA     = "Papelera";
const H_ACTIVIDAD    = "Actividad";
const H_RESPONSABLES = "Responsables";
const H_CAMBIOS      = "Cambios"; 
const H_COMPRAS      = "Compras"; // NUEVA HOJA

const HDR = {
  Categorias:    ["id","nombre","emoji","activo"],
  Productos:     ["id","nombre","código","categoría","precio_compra","precio_venta","stock","stock_minimo","imagen_url","favorito","activo","fecha_creado"],
  Ventas:        ["id","fecha","cliente_nombre","subtotal","descuento","impuesto","total","pago_con","cambio","usuario","estado"],
  VentaDetalle:  ["id","venta_id","producto_id","producto_nombre","cantidad","precio_unitario","descuento_linea","subtotal_linea"],
  Papelera:      ["id","tipo","datos_originales","fecha_eliminado","eliminado_por"],
  Actividad:     ["id","fecha","usuario","accion","detalle"],
  Responsables:  ["id","email","nombre","reportes","umbral_venta_grande","activo","ultimo_envio_mensual","ultimo_envio_anual","ultimo_envio_stock","fecha_creado","ultimo_envio_backup"],
  Cambios:       ["id","venta_id","fecha","cliente_nombre","monto","estado","fecha_pagado","usuario"],
  Compras:       ["id","producto_id","producto_nombre","cantidad","costo_total","fecha","usuario"] // NUEVA CABECERA
};

function ss()  { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function uid() { return 'id-' + (Date.now().toString(36) + Math.random().toString(36).substring(2,9)).toUpperCase(); }
function sh(n) { return ss().getSheetByName(n); }
function resp(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function log(usuario, accion, detalle) { try { sh(H_ACTIVIDAD).appendRow([uid(), new Date(), usuario||'Sistema', accion, detalle||'']); } catch(e){} }

function getData(nombre) {
  const hoja = sh(nombre);
  if (!hoja || hoja.getLastRow() < 2) return { status:'error', message:`'${nombre}' vacía.` };
  const vals = hoja.getDataRange().getValues();
  const heads = vals[0];
  const rows = vals.slice(1).map(r => {
    const o = {}; heads.forEach((h, i) => { o[h] = (r[i] instanceof Date) ? r[i].toISOString() : (r[i] === '' || r[i] == null ? '' : r[i]); });
    return o;
  }).filter(r => Object.values(r).some(v => v !== ''));
  return { status:'success', data: rows };
}

function findRow(hoja, id) {
  const vals = hoja.getDataRange().getValues();
  const idStr = String(id).toLowerCase();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === idStr) return { row: vals[i], idx: i };
  }
  return { row: null, idx: -1 };
}

function crearHoja(nombre) {
  const spreadsheet = ss();
  if (!HDR[nombre]) return;
  let hoja = spreadsheet.getSheetByName(nombre);
  if (!hoja) hoja = spreadsheet.insertSheet(nombre);
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1,1,1,HDR[nombre].length).setValues([HDR[nombre]]);
    hoja.setFrozenRows(1);
  }
}

function doGet(e) {
  const p = e.parameter; let r;
  try {
    switch(p.action) {
      case 'iniciar':   r = iniciarBD(); break;
      case 'getCategorias': r = getData(H_CATEGORIAS); break;
      case 'getInventario': r = {status:'success', data: (getData(H_PRODUCTOS).data || []).filter(x => x.activo === true || x.activo === 'true' || x.activo === 1)}; break;
      case 'getVentas': r = getVentasConDetalle(); break;
      case 'getPapelera': r = getData(H_PAPELERA); break;
      case 'getCambios': r = getData(H_CAMBIOS); break; 
      case 'getCompras': r = getData(H_COMPRAS); break; // Obtener historial de compras
      case 'getResponsables': r = getData(H_RESPONSABLES); break;
      case 'exportarDatos': r = exportarDatosCompletos(); break; // Copia de seguridad manual (descarga)
      default: r = {status:'error', message:`Acción '${p.action}' no válida.`};
    }
  } catch(ex) { r = {status:'error', message:ex.message}; }
  return resp(r);
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents); let r;
    switch(req.action) {
      case 'agregarCategoria': sh(H_CATEGORIAS).appendRow([uid(), req.nombre, req.emoji||'📦', true]); r = {status:'success', message:'Categoría agregada.'}; break;
      case 'eliminarCategoria': r = eliminarEntidad(H_CATEGORIAS, req, 'Categoría'); break;
      case 'subirImagenProducto': r = subirImagenDrive(req); break;
      case 'agregarProducto': {
        const pv = Math.max(0, parseFloat(req.precio_venta)||0);
        // Si el precio de compra viene vacío/0, se usa el mismo precio de venta
        const pc = Math.max(0, parseFloat(req.precio_compra) || pv);
        const stockInicial = Math.max(0, parseInt(req.stock)||0);
        const stockMin = Math.max(0, parseInt(req.stock_minimo)||5);
        sh(H_PRODUCTOS).appendRow([uid(), req.nombre, String(req.codigo||''), req.categoria, pc, pv, stockInicial, stockMin, req.imagen_url||'', false, true, new Date()]);
        r = {status:'success', message:'Producto registrado.'};
        break;
      }
      case 'editarProducto': r = editarProducto(req); break;
      case 'eliminarProducto': r = eliminarProducto(req); break;
      case 'restaurarProducto': r = restaurarProducto(req); break;
      case 'registrarVenta': r = registrarVenta(req); break;
      case 'anularVenta': r = anularVenta(req); break;
      case 'marcarCambioPagado': r = marcarCambioPagado(req); break;
      case 'registrarCompra': r = registrarCompra(req); break; // Registro de gastos y stock
      case 'agregarResponsable': r = agregarResponsable(req); break;
      case 'editarResponsable': r = editarResponsable(req); break;
      case 'eliminarResponsable': r = eliminarEntidad(H_RESPONSABLES, req, 'Responsable'); break;
      case 'probarEnvioResponsable': r = probarEnvioResponsable(req); break;
      case 'importarDatos': r = importarDatosCompletos(req); break; // Restaurar copia de seguridad
      case 'enviarBackupAhora': r = enviarBackupManual(); break; // Forzar envío de backup a responsables
      default: r = {status:'error', message:'Acción no reconocida'};
    }
    return resp(r);
  } catch(ex) { return resp({status:'error', message:ex.message}); }
}

function subirImagenDrive(data) {
  try {
    const carpetas = DriveApp.getFoldersByName("POS_Imagenes_Productos");
    let carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder("POS_Imagenes_Productos");
    carpeta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const blob = Utilities.newBlob(Utilities.base64Decode(data.base64), data.mimeType || 'image/jpeg', data.filename || 'img.jpg');
    const url = "https://drive.google.com/uc?export=view&id=" + carpeta.createFile(blob).getId();
    return { status: 'success', url: url };
  } catch (e) { return { status: 'error', message: 'No se pudo subir la imagen.' }; }
}

function registrarVenta(data) {
  const shV = sh(H_VENTAS); const shD = sh(H_VENTA_DET); const shP = sh(H_PRODUCTOS); const shC = sh(H_CAMBIOS);
  let subtotal = 0;
  // Primero se valida TODO el carrito contra el stock real (incluyendo cantidades repetidas
  // del mismo producto) antes de escribir nada, para nunca dejar una venta a medias.
  const acumulado = {}; // producto_id -> cantidad total pedida en este carrito
  for (const item of data.items) {
    const cant = parseInt(item.cantidad) || 0;
    if (cant <= 0) return {status:'error', message:`Cantidad inválida para ${item.producto_id}.`};
    const {row, idx} = findRow(shP, item.producto_id);
    if (!row) return {status:'error', message:`Producto no encontrado.`};
    if (row[10] === false || row[10] === 'false') return {status:'error', message:`El producto "${row[1]}" fue eliminado del inventario.`};
    acumulado[item.producto_id] = (acumulado[item.producto_id] || 0) + cant;
    const stockDisponible = parseInt(row[6]) || 0;
    if (stockDisponible < acumulado[item.producto_id]) {
      return {status:'warning', message:`Stock insuficiente de "${row[1]}". Disponible: ${stockDisponible}, solicitado: ${acumulado[item.producto_id]}.`};
    }
    item._rowIdx = idx; item._nombre = row[1]; item._subtotal = (parseFloat(item.precio_unitario)*cant);
    subtotal += item._subtotal;
  }
  const total = subtotal;
  const ventaId = uid();
  const nombreCliente = data.cliente_nombre && String(data.cliente_nombre).trim() !== '' ? String(data.cliente_nombre).trim() : 'N/A';

  shV.appendRow([ventaId, new Date(), nombreCliente, subtotal, 0, 0, total, data.pago_con||0, data.cambio||0, data.usuario||'Sistema', 'completada']);
  
  for (const item of data.items) {
    shD.appendRow([uid(), ventaId, item.producto_id, item._nombre, parseInt(item.cantidad), parseFloat(item.precio_unitario), 0, item._subtotal]);
    shP.getRange(item._rowIdx+1, 7).setValue((parseInt(shP.getRange(item._rowIdx+1, 7).getValue())||0) - parseInt(item.cantidad));
  }

  if (data.pendiente && data.cambio > 0) {
    shC.appendRow([uid(), ventaId, new Date(), nombreCliente, data.cambio, 'Pendiente', '', data.usuario||'Sistema']);
  }

  try { avisarVentaGrande_(total, nombreCliente, ventaId); } catch(e) {}

  return { status:'success', message:'Venta registrada con éxito.' };
}

// Anula (elimina) una venta: guarda copia completa + justificación en la Papelera,
// devuelve al inventario el stock que se había descontado, borra cualquier "cambio
// pendiente" ligado a esa venta, y elimina la venta de Ventas/VentaDetalle para que
// deje de contarse en Reportes, Resúmenes y el total recaudado.
function anularVenta(data) {
  const shV = sh(H_VENTAS); const shD = sh(H_VENTA_DET); const shP = sh(H_PRODUCTOS); const shPap = sh(H_PAPELERA); const shC = sh(H_CAMBIOS);
  const { row: rowV } = findRow(shV, data.id);
  if (!rowV) return { status:'error', message:'Venta no encontrada.' };
  if (!data.motivo || !String(data.motivo).trim()) return { status:'error', message:'Debes indicar un motivo de la anulación.' };

  // Recolectar las líneas de detalle asociadas a esta venta
  const detalleVals = shD.getDataRange().getValues();
  const itemsVenta = [];
  const filasDetalleABorrar = [];
  for (let i = 1; i < detalleVals.length; i++) {
    if (String(detalleVals[i][1]) === String(data.id)) {
      itemsVenta.push({
        producto_id: detalleVals[i][2],
        producto_nombre: detalleVals[i][3],
        cantidad: detalleVals[i][4],
        precio_unitario: detalleVals[i][5],
        subtotal_linea: detalleVals[i][7]
      });
      filasDetalleABorrar.push(i + 1);
    }
  }

  // Devolver el stock vendido a cada producto
  itemsVenta.forEach(it => {
    const { idx: idxP } = findRow(shP, it.producto_id);
    if (idxP > -1) {
      const stockActual = parseInt(shP.getRange(idxP + 1, 7).getValue()) || 0;
      shP.getRange(idxP + 1, 7).setValue(stockActual + (parseInt(it.cantidad) || 0));
    }
  });

  // Guardar copia completa (venta + items + motivo) en la Papelera antes de borrar
  const datosOriginales = {
    venta: { id: rowV[0], fecha: rowV[1], cliente_nombre: rowV[2], subtotal: rowV[3], descuento: rowV[4], impuesto: rowV[5], total: rowV[6], pago_con: rowV[7], cambio: rowV[8], usuario: rowV[9], estado: rowV[10] },
    items: itemsVenta,
    motivo: String(data.motivo).trim()
  };
  shPap.appendRow([uid(), 'Venta Anulada', JSON.stringify(datosOriginales), new Date(), data.usuario || 'Sistema']);

  // Eliminar cualquier "cambio pendiente" ligado a esta venta, ya no aplica
  const cambiosVals = shC.getDataRange().getValues();
  for (let i = cambiosVals.length - 1; i >= 1; i--) {
    if (String(cambiosVals[i][1]) === String(data.id)) shC.deleteRow(i + 1);
  }

  // Eliminar las líneas de detalle (de mayor a menor índice para no desordenar filas)
  filasDetalleABorrar.sort((a, b) => b - a).forEach(i => shD.deleteRow(i));

  // Eliminar la venta (así deja de sumar en Reportes/Resúmenes/total recaudado)
  const { idx: idxVentaFinal } = findRow(shV, data.id);
  if (idxVentaFinal > -1) shV.deleteRow(idxVentaFinal + 1);

  log(data.usuario, 'Venta Anulada', `Motivo: ${data.motivo}. Cliente: ${rowV[2]}, Total: $${(parseFloat(rowV[6])||0).toFixed(2)}`);

  return { status:'success', message:'Venta eliminada. El stock fue devuelto al inventario.' };
}

function registrarCompra(data) {
  const shP = sh(H_PRODUCTOS);
  const shC = sh(H_COMPRAS);
  
  const {row, idx} = findRow(shP, data.producto_id);
  if (!row) return {status:'error', message:'Producto no encontrado.'};
  
  const cantidad = parseInt(data.cantidad) || 0;
  const costo = parseFloat(data.costo) || 0;
  if (cantidad <= 0) return {status:'error', message:'La cantidad comprada debe ser mayor a 0.'};
  if (costo < 0) return {status:'error', message:'El costo no puede ser negativo.'};
  
  // Actualizar Stock sumando lo comprado
  const stockActual = parseInt(row[6]) || 0;
  shP.getRange(idx+1, 7).setValue(stockActual + cantidad);
  
  // Registrar la compra en la tabla Compras
  shC.appendRow([uid(), data.producto_id, row[1], cantidad, costo, new Date().toISOString(), data.usuario || 'Sistema']);
  log(data.usuario, 'Compra Stock', `Producto: ${row[1]} (+${cantidad}), Costo: $${costo}`);
  
  return {status:'success', message:'Compra registrada y stock actualizado con éxito.'};
}

function marcarCambioPagado(data) {
  const hoja = sh(H_CAMBIOS); const {idx, row} = findRow(hoja, data.id);
  if (idx > -1) {
    hoja.getRange(idx+1, 6).setValue('Pagado'); 
    hoja.getRange(idx+1, 7).setValue(new Date().toISOString());
    return {status: 'success', message: 'Cambio marcado como pagado.'};
  }
  return {status: 'error', message: 'Registro no encontrado.'};
}

function editarProducto(data) {
  const hoja = sh(H_PRODUCTOS); const {idx} = findRow(hoja, data.id);
  if (idx < 0) return {status:'error', message:'No encontrado.'};
  if (data.nombre) hoja.getRange(idx+1, 2).setValue(data.nombre);
  const pvEdit = data.precio_venta !== undefined ? (parseFloat(data.precio_venta)||0) : undefined;
  if (data.precio_compra !== undefined) hoja.getRange(idx+1, 5).setValue(Math.max(0, parseFloat(data.precio_compra) || pvEdit || 0));
  if (data.precio_venta !== undefined) hoja.getRange(idx+1, 6).setValue(Math.max(0, pvEdit));
  if (data.stock !== undefined) hoja.getRange(idx+1, 7).setValue(Math.max(0, parseInt(data.stock) || 0));
  if (data.stock_minimo !== undefined) hoja.getRange(idx+1, 8).setValue(Math.max(0, parseInt(data.stock_minimo) || 5));
  if (data.categoria) hoja.getRange(idx+1, 4).setValue(data.categoria);
  if (data.imagen_url) hoja.getRange(idx+1, 9).setValue(data.imagen_url);
  return {status:'success', message:'Producto actualizado.'};
}

function eliminarProducto(data) {
  const hoja = sh(H_PRODUCTOS); const papelera = sh(H_PAPELERA); const {row, idx} = findRow(hoja, data.id);
  if (idx > -1) {
    papelera.appendRow([uid(), 'Producto', JSON.stringify({id:row[0], nombre:row[1], código:row[2], categoría:row[3], precio_compra:row[4], precio_venta:row[5], stock:row[6]}), new Date(), data.usuario||'Sistema']);
    hoja.getRange(idx+1, 11).setValue(false);
    return {status:'success', message:`Producto eliminado.`};
  } return {status:'error', message:'No encontrado.'};
}

function restaurarProducto(data) {
  const shPap = sh(H_PAPELERA); const shProd = sh(H_PRODUCTOS);
  const rows = shPap.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.papelera_id)) {
      const orig = JSON.parse(rows[i][2]);
      const {idx} = findRow(shProd, orig.id);
      if (idx > -1) shProd.getRange(idx+1, 11).setValue(true);
      shPap.deleteRow(i+1);
      return {status:'success', message:'Producto restaurado.'};
    }
  } return {status:'error', message:'No encontrado.'};
}

function eliminarEntidad(hoja, data, tipo) {
  const hojaObj = sh(hoja); const {idx} = findRow(hojaObj, data.id);
  if (idx > -1) { hojaObj.deleteRow(idx+1); return {status:'success', message:`${tipo} eliminada.`}; }
  return {status:'error', message:'No encontrado.'};
}

function agregarResponsable(data) {
  sh(H_RESPONSABLES).appendRow([uid(), data.email.trim(), data.nombre || '', data.reportes || '', parseFloat(data.umbral_venta_grande)||100, data.activo !== false, '', '', '', new Date(), '']);
  return {status:'success', message:'Responsable agregado con éxito.'};
}

function editarResponsable(data) {
  const hoja = sh(H_RESPONSABLES); const {idx} = findRow(hoja, data.id);
  if (idx < 0) return {status:'error', message:'Responsable no encontrado.'};
  if (data.email) hoja.getRange(idx+1, 2).setValue(data.email.trim());
  if (data.nombre !== undefined) hoja.getRange(idx+1, 3).setValue(data.nombre);
  if (data.reportes !== undefined) hoja.getRange(idx+1, 4).setValue(data.reportes);
  if (data.umbral_venta_grande !== undefined) hoja.getRange(idx+1, 5).setValue(parseFloat(data.umbral_venta_grande)||0);
  if (data.activo !== undefined) hoja.getRange(idx+1, 6).setValue(data.activo);
  return {status:'success', message:'Responsable actualizado con éxito.'};
}

function getVentasConDetalle() {
  const ventas = getData(H_VENTAS).data || [];
  const detalles = getData(H_VENTA_DET).data || [];
  return {status:'success', data: ventas.map(v => ({...v, items: detalles.filter(d => String(d.venta_id) === String(v.id))}))};
}

function iniciarBD() {
  [H_CATEGORIAS,H_PRODUCTOS,H_VENTAS,H_VENTA_DET,H_PAPELERA,H_ACTIVIDAD,H_RESPONSABLES,H_CAMBIOS,H_COMPRAS].forEach(crearHoja);
  return {status:'success', message:'BD inicializada correctamente.'};
}

// ═══════════════════════════════════════════════════════════════
// EXPORTAR / IMPORTAR DATOS (Copia de seguridad manual desde Ajustes)
// Usa exactamente el mismo formato que la copia de seguridad que se
// envía automáticamente por correo a los Responsables (construirBlobJSON_),
// para que el archivo descargado y el adjunto por correo sean intercambiables.
// ═══════════════════════════════════════════════════════════════
function exportarDatosCompletos() {
  try {
    const productos = getData(H_PRODUCTOS).data || [];
    const categorias = getData(H_CATEGORIAS).data || [];
    const ventas = getData(H_VENTAS).data || [];
    const detalle = getData(H_VENTA_DET).data || [];
    return { status:'success', data: { productos, categorias, ventas, detalle_ventas: detalle, fecha_respaldo: new Date().toISOString() } };
  } catch(e) {
    return { status:'error', message:'No se pudo generar la copia de seguridad: ' + e.message };
  }
}

// Reemplaza todo el contenido (sin encabezado) de una hoja con nuevas filas,
// construidas a partir de objetos {campo: valor} según el orden de `headers`.
function reemplazarHoja_(nombreHoja, filasObjs, headers) {
  const hoja = sh(nombreHoja);
  if (!hoja) return 0;
  const lastRow = hoja.getLastRow();
  if (lastRow > 1) hoja.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (!filasObjs || !filasObjs.length) return 0;
  const filas = filasObjs.map(o => headers.map(h => (o[h] !== undefined && o[h] !== null) ? o[h] : ''));
  hoja.getRange(2, 1, filas.length, headers.length).setValues(filas);
  return filas.length;
}

function importarDatosCompletos(data) {
  try {
    const backup = data.backup;
    if (!backup || typeof backup !== 'object') return { status:'error', message:'Archivo de copia de seguridad no válido.' };

    let resumen = [];
    if (Array.isArray(backup.categorias)) resumen.push(`${reemplazarHoja_(H_CATEGORIAS, backup.categorias, HDR.Categorias)} categorías`);
    if (Array.isArray(backup.productos)) resumen.push(`${reemplazarHoja_(H_PRODUCTOS, backup.productos, HDR.Productos)} productos`);
    if (Array.isArray(backup.ventas)) resumen.push(`${reemplazarHoja_(H_VENTAS, backup.ventas, HDR.Ventas)} ventas`);
    if (Array.isArray(backup.detalle_ventas)) resumen.push(`${reemplazarHoja_(H_VENTA_DET, backup.detalle_ventas, HDR.VentaDetalle)} líneas de detalle`);

    if (!resumen.length) return { status:'error', message:'El archivo no contiene datos reconocibles (productos, categorías, ventas o detalle_ventas).' };

    log('Sistema', 'Importar Copia de Seguridad', resumen.join(', '));
    return { status:'success', message:'Datos restaurados: ' + resumen.join(', ') + '.' };
  } catch(e) {
    return { status:'error', message:'Error al importar los datos: ' + e.message };
  }
}

// Envía la copia de seguridad AHORA MISMO por correo a los responsables que
// la tengan activada, sin esperar el activador diario ni revisar si ya se
// envió hoy (botón manual en Ajustes).
function enviarBackupManual() {
  try {
    const enviados = respaldoDiario_(new Date(), true);
    if (!enviados) return { status:'warning', message:'No hay responsables activos con "Copia de seguridad" (JSON/CSV) habilitada. Actívala en la sección Responsables.' };
    const plural = enviados === 1 ? '1 responsable' : `${enviados} responsables`;
    return { status:'success', message:`Copia de seguridad enviada por correo a ${plural}.` };
  } catch(e) {
    return { status:'error', message:'Error al enviar la copia de seguridad: ' + e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// SCRIPT DE INICIALIZACIÓN RÁPIDA (Ejecutar desde el editor)
// ═══════════════════════════════════════════════════════════════
function configuracionInicial() {
  iniciarBD();
  const hojaCat = ss().getSheetByName(H_CATEGORIAS);
  if (hojaCat && hojaCat.getLastRow() === 1) {
    hojaCat.appendRow([uid(), "Bebidas", "🥤", true]);
    hojaCat.appendRow([uid(), "Snacks", "🍪", true]);
  }
}

// ═══════════════════════════════════════════════════════════════
// CORREOS AUTOMÁTICOS PARA RESPONSABLES
// ═══════════════════════════════════════════════════════════════
// Columnas de la hoja "Responsables" (1-based, para usar con getRange):
// 1=id  2=email  3=nombre  4=reportes  5=umbral_venta_grande  6=activo
// 7=ultimo_envio_mensual  8=ultimo_envio_anual  9=ultimo_envio_stock
// 10=fecha_creado  11=ultimo_envio_backup
const RESP_COL = { ultimo_envio_mensual: 7, ultimo_envio_anual: 8, ultimo_envio_stock: 9, ultimo_envio_backup: 11 };

function activoResp_(r)          { return r.activo === true || r.activo === 'true'; }
function tieneReporte_(r, tipo)  { return String(r.reportes || '').split(',').map(s => s.trim()).includes(tipo); }
function yaEnviadoHoy_(valorGuardado, hoy) {
  if (!valorGuardado) return false;
  const d = new Date(valorGuardado);
  return d.getFullYear() === hoy.getFullYear() && d.getMonth() === hoy.getMonth() && d.getDate() === hoy.getDate();
}
function marcarEnvioResponsable_(id, campo, fecha) {
  const hoja = sh(H_RESPONSABLES); const {idx} = findRow(hoja, id);
  if (idx > -1 && RESP_COL[campo]) hoja.getRange(idx + 1, RESP_COL[campo]).setValue(fecha.toISOString());
}
function formatoFecha_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/El_Salvador', 'yyyy-MM-dd');
}

// ── Punto de entrada único: crea UN solo activador de tiempo diario
// (ver instructivo) que llame a esta función. Ella decide internamente
// qué correos le tocan hoy a cada responsable. ──
function tareasDiarias() {
  const hoy = new Date();
  try { enviarAlertaStockBajo_(hoy); } catch(e) { log('Sistema','Error alerta stock', e.message); }
  try { respaldoDiario_(hoy); } catch(e) { log('Sistema','Error respaldo', e.message); }
  if (hoy.getDate() === 1) {
    try { enviarResumenMensual_(hoy); } catch(e) { log('Sistema','Error resumen mensual', e.message); }
  }
  if (hoy.getMonth() === 0 && hoy.getDate() === 1) {
    try { enviarResumenAnual_(hoy); } catch(e) { log('Sistema','Error resumen anual', e.message); }
  }
}

// 💰 Aviso INMEDIATO de venta grande — se llama desde registrarVenta().
function avisarVentaGrande_(total, cliente, ventaId) {
  const responsables = (getData(H_RESPONSABLES).data || []).filter(r => activoResp_(r) && tieneReporte_(r, 'ventas_grandes'));
  responsables.forEach(r => {
    const umbral = parseFloat(r.umbral_venta_grande) || 0;
    if (umbral > 0 && total >= umbral) {
      try { MailApp.sendEmail({ to: r.email, subject: `💰 Venta grande registrada: $${total.toFixed(2)}`, htmlBody: contenidoVentaGrande_(total, cliente, ventaId, umbral) }); } catch(e) {}
    }
  });
}
function contenidoVentaGrande_(total, cliente, ventaId, umbral) {
  return `<p>Se registró una venta de <b>$${total.toFixed(2)}</b> a <b>${cliente || 'N/A'}</b>, superando el umbral de $${umbral.toFixed(2)}.</p><p>Fecha: ${new Date().toLocaleString('es-SV')}</p><p style="color:#888;font-size:12px;">ID de venta: ${ventaId}</p>`;
}

// ⚠️ Alerta diaria de stock bajo.
function contenidoStockBajo_() {
  const productos = (getData(H_PRODUCTOS).data || []).filter(p =>
    (p.activo === true || p.activo === 'true') && (parseInt(p.stock) || 0) <= (parseInt(p.stock_minimo) || 5));
  if (!productos.length) return null;
  const filas = productos.map(p => `<tr><td style="padding:4px 10px;">${p.nombre}</td><td style="padding:4px 10px;color:#dc3545;font-weight:bold;">${p.stock}</td><td style="padding:4px 10px;">${p.stock_minimo}</td></tr>`).join('');
  return {
    asunto: `⚠️ Alerta de stock bajo (${productos.length} producto(s))`,
    html: `<h3>⚠️ Alerta de Stock Bajo</h3><p>${productos.length} producto(s) necesitan reabastecerse:</p>
      <table border="1" cellspacing="0" style="border-collapse:collapse;"><tr style="background:#f4f4f4;"><th style="padding:4px 10px;">Producto</th><th style="padding:4px 10px;">Stock actual</th><th style="padding:4px 10px;">Stock mínimo</th></tr>${filas}</table>`
  };
}
function enviarAlertaStockBajo_(hoy) {
  const responsables = (getData(H_RESPONSABLES).data || []).filter(r =>
    activoResp_(r) && tieneReporte_(r, 'stock_bajo') && !yaEnviadoHoy_(r.ultimo_envio_stock, hoy));
  if (!responsables.length) return;
  const contenido = contenidoStockBajo_();
  if (!contenido) return;
  responsables.forEach(r => {
    try {
      MailApp.sendEmail({ to: r.email, subject: contenido.asunto, htmlBody: contenido.html });
      marcarEnvioResponsable_(r.id, 'ultimo_envio_stock', hoy);
    } catch(e) {}
  });
}

// 📊 Resumen mensual (se envía el día 1, resume el mes que acaba de terminar).
function contenidoResumenMensual_(mesRef) {
  const ventas = (getData(H_VENTAS).data || []).filter(v => {
    const d = new Date(v.fecha);
    return d.getFullYear() === mesRef.getFullYear() && d.getMonth() === mesRef.getMonth();
  });
  const totalMes = ventas.reduce((s, v) => s + (parseFloat(v.total) || 0), 0);
  const nombreMes = mesRef.toLocaleDateString('es-SV', { month: 'long', year: 'numeric' });
  return {
    asunto: `📊 Resumen mensual — ${nombreMes}`,
    html: `<h3>📊 Resumen de ${nombreMes}</h3><p>Ventas registradas: <b>${ventas.length}</b></p><p>Total recaudado: <b>$${totalMes.toFixed(2)}</b></p>`
  };
}
function enviarResumenMensual_(hoy) {
  const responsables = (getData(H_RESPONSABLES).data || []).filter(r =>
    activoResp_(r) && tieneReporte_(r, 'mensual') && !yaEnviadoHoy_(r.ultimo_envio_mensual, hoy));
  if (!responsables.length) return;
  const mesRef = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  const contenido = contenidoResumenMensual_(mesRef);
  responsables.forEach(r => {
    try {
      MailApp.sendEmail({ to: r.email, subject: contenido.asunto, htmlBody: contenido.html });
      marcarEnvioResponsable_(r.id, 'ultimo_envio_mensual', hoy);
    } catch(e) {}
  });
}

// 📅 Resumen anual (se envía el 1 de enero, resume el año que acaba de terminar).
function contenidoResumenAnual_(anioRef) {
  const ventas = (getData(H_VENTAS).data || []).filter(v => new Date(v.fecha).getFullYear() === anioRef);
  const totalAnio = ventas.reduce((s, v) => s + (parseFloat(v.total) || 0), 0);
  return {
    asunto: `📅 Resumen anual — ${anioRef}`,
    html: `<h3>📅 Resumen anual ${anioRef}</h3><p>Ventas registradas: <b>${ventas.length}</b></p><p>Total recaudado: <b>$${totalAnio.toFixed(2)}</b></p>`
  };
}
function enviarResumenAnual_(hoy) {
  const responsables = (getData(H_RESPONSABLES).data || []).filter(r =>
    activoResp_(r) && tieneReporte_(r, 'anual') && !yaEnviadoHoy_(r.ultimo_envio_anual, hoy));
  if (!responsables.length) return;
  const anioRef = hoy.getFullYear() - 1;
  const contenido = contenidoResumenAnual_(anioRef);
  responsables.forEach(r => {
    try {
      MailApp.sendEmail({ to: r.email, subject: contenido.asunto, htmlBody: contenido.html });
      marcarEnvioResponsable_(r.id, 'ultimo_envio_anual', hoy);
    } catch(e) {}
  });
}

// 📅📊 Copia de seguridad diaria (JSON y/o CSV, según lo que cada responsable
// haya marcado). Se manda en un solo correo por responsable con los
// adjuntos que le correspondan, para no duplicar el registro de envío.
function construirBlobJSON_(fechaTxt) {
  const productos = getData(H_PRODUCTOS).data || [];
  const ventas = getData(H_VENTAS).data || [];
  const detalle = getData(H_VENTA_DET).data || [];
  const categorias = getData(H_CATEGORIAS).data || [];
  return Utilities.newBlob(
    JSON.stringify({ productos, ventas, detalle_ventas: detalle, categorias, fecha_respaldo: new Date().toISOString() }, null, 2),
    'application/json', `respaldo_${fechaTxt}.json`);
}
function construirBlobCSV_(fechaTxt) {
  const ventas = getData(H_VENTAS).data || [];
  return Utilities.newBlob(construirCSVVentas_(ventas), 'text/csv', `respaldo_ventas_${fechaTxt}.csv`);
}
// `forzar` (opcional): si es true, ignora el control de "ya enviado hoy" y
// manda de inmediato a todos los responsables con backup habilitado (lo usa
// el botón manual "Enviar Copia de Seguridad a Responsables Ahora"). Devuelve
// la cantidad de responsables a los que se les envió el correo.
function respaldoDiario_(hoy, forzar) {
  const responsables = (getData(H_RESPONSABLES).data || []).filter(r =>
    activoResp_(r) && (tieneReporte_(r, 'backup_json') || tieneReporte_(r, 'backup_csv')) && (forzar || !yaEnviadoHoy_(r.ultimo_envio_backup, hoy)));
  if (!responsables.length) return 0;

  const fechaTxt = formatoFecha_(hoy);
  const jsonBlob = construirBlobJSON_(fechaTxt);
  const csvBlob = construirBlobCSV_(fechaTxt);
  let enviados = 0;

  responsables.forEach(r => {
    const adjuntos = [];
    if (tieneReporte_(r, 'backup_json')) adjuntos.push(jsonBlob);
    if (tieneReporte_(r, 'backup_csv')) adjuntos.push(csvBlob);
    if (!adjuntos.length) return;
    try {
      MailApp.sendEmail({ to: r.email, subject: `📅 Copia de seguridad — ${fechaTxt}`, body: 'Adjunto la copia de seguridad diaria de la base de datos del sistema.', attachments: adjuntos });
      marcarEnvioResponsable_(r.id, 'ultimo_envio_backup', hoy);
      enviados++;
    } catch(e) {}
  });
  return enviados;
}

function construirCSVVentas_(ventas) {
  const headers = ['id','fecha','cliente_nombre','subtotal','total','pago_con','cambio','usuario','estado'];
  const filas = [headers.join(',')];
  ventas.forEach(v => filas.push(headers.map(h => `"${String(v[h] || '').replace(/"/g, '""')}"`).join(',')));
  return filas.join('\n');
}

const TIPOS_REPORTE_TODOS = ['mensual', 'anual', 'ventas_grandes', 'stock_bajo', 'backup_json', 'backup_csv'];

// Construye asunto/html/adjuntos de prueba para UN tipo de reporte. Devuelve
// null si el tipo no existe.
function construirPruebaTipo_(tipo, row, hoy) {
  const umbral = parseFloat(row[4]) || 0;
  switch (tipo) {
    case 'mensual': {
      const c = contenidoResumenMensual_(new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1));
      return { asunto: c.asunto, html: c.html, adjuntos: [] };
    }
    case 'anual': {
      const c = contenidoResumenAnual_(hoy.getFullYear() - 1);
      return { asunto: c.asunto, html: c.html, adjuntos: [] };
    }
    case 'ventas_grandes':
      return {
        asunto: '💰 Aviso de venta grande',
        html: `<p>Este es un correo de <b>prueba</b>. Así se vería un aviso real cuando se registre una venta igual o mayor a tu umbral actual de <b>$${umbral.toFixed(2)}</b>.</p>`,
        adjuntos: []
      };
    case 'stock_bajo': {
      const c = contenidoStockBajo_();
      if (!c) return { asunto: '⚠️ Alerta de stock bajo', html: '<p>Este es un correo de <b>prueba</b>. Ahora mismo no hay productos con stock bajo, pero así se vería la alerta cuando sí los haya.</p>', adjuntos: [] };
      return { asunto: c.asunto, html: c.html, adjuntos: [] };
    }
    case 'backup_json': {
      const fechaTxt = formatoFecha_(hoy);
      return { asunto: `📅 Copia de seguridad JSON — ${fechaTxt}`, html: '<p>Este es un correo de <b>prueba</b>. Adjunto la copia de seguridad en formato JSON.</p>', adjuntos: [construirBlobJSON_(fechaTxt)] };
    }
    case 'backup_csv': {
      const fechaTxt = formatoFecha_(hoy);
      return { asunto: `📊 Copia de seguridad CSV — ${fechaTxt}`, html: '<p>Este es un correo de <b>prueba</b>. Adjunto la copia de seguridad de ventas en formato CSV.</p>', adjuntos: [construirBlobCSV_(fechaTxt)] };
    }
    default:
      return null;
  }
}

// ── Botón "Probar Envío" en Responsables: manda AHORA MISMO, sin esperar
// la fecha programada ni revisar si ya se envió hoy, el tipo de reporte que
// el usuario elija en el selector (o TODOS los tipos si elige "todos"),
// solo a ese responsable. ──
function probarEnvioResponsable(data) {
  const hoja = sh(H_RESPONSABLES); const { row } = findRow(hoja, data.id);
  if (!row) return { status:'error', message:'Responsable no encontrado.' };
  const email = row[1];
  const hoy = new Date();

  const esTodos = data.tipo === 'todos';
  const tiposAProbar = esTodos ? TIPOS_REPORTE_TODOS : [data.tipo];

  let enviados = 0;
  for (const tipo of tiposAProbar) {
    const contenido = construirPruebaTipo_(tipo, row, hoy);
    if (!contenido) {
      if (!esTodos) return { status:'error', message:'Tipo de reporte no válido.' };
      continue;
    }
    try {
      MailApp.sendEmail({ to: email, subject: '[PRUEBA] ' + contenido.asunto, htmlBody: contenido.html, attachments: contenido.adjuntos });
      enviados++;
    } catch(e) { /* seguir intentando con los demás tipos */ }
  }

  if (!enviados) return { status:'error', message:'No se pudo enviar ningún correo de prueba.' };
  const plural = enviados === 1 ? 'Correo de prueba enviado' : `${enviados} correos de prueba enviados`;
  return { status:'success', message:`${plural} a ${email}.` };
}

// ── Función de prueba: ejecútala UNA VEZ manualmente desde el editor
// (▶ Ejecutar, eligiendo "probarConexionCorreo") para autorizar el envío
// de correos y confirmar que llega el mensaje de prueba a tu cuenta. ──
function probarConexionCorreo() {
  const destino = Session.getActiveUser().getEmail();
  MailApp.sendEmail(destino, '✅ Prueba ERP POS LITE', 'Si recibiste este correo, el envío automático de reportes está funcionando correctamente.');
}