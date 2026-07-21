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
const H_CAMBIOS_MOV  = "CambiosMovimientos"; // NUEVA HOJA — historial de "mini cambios" (entregas parciales/totales)
const H_JORNADAS     = "Jornadas"; // NUEVA HOJA — Jornadas de Venta (Stock Temporal / Stock de Maleta)
const H_JORNADA_DET  = "DetalleJornada"; // NUEVA HOJA — detalle por producto de cada Jornada

const HDR = {
  Categorias:    ["id","nombre","emoji","activo"],
  Productos:     ["id","nombre","código","categoría","precio_compra","precio_venta","stock","stock_minimo","imagen_url","favorito","activo","fecha_creado"],
  Ventas:        ["id","fecha","cliente_nombre","subtotal","descuento","impuesto","total","pago_con","cambio","usuario","estado","jornada_id"],
  VentaDetalle:  ["id","venta_id","producto_id","producto_nombre","cantidad","precio_unitario","descuento_linea","subtotal_linea"],
  Papelera:      ["id","tipo","datos_originales","fecha_eliminado","eliminado_por"],
  Actividad:     ["id","fecha","usuario","accion","detalle"],
  Responsables:  ["id","email","nombre","reportes","umbral_venta_grande","activo","ultimo_envio_mensual","ultimo_envio_anual","ultimo_envio_stock","fecha_creado","ultimo_envio_backup","ultimo_envio_cierre"],
  Cambios:       ["id","venta_id","fecha","cliente_nombre","monto","estado","fecha_pagado","usuario","monto_original"],
  Compras:       ["id","producto_id","producto_nombre","cantidad","costo_total","fecha","usuario"], // NUEVA CABECERA
  CambiosMovimientos: ["id","cambio_id","fecha","cliente_nombre","monto","tipo","usuario"], // NUEVA CABECERA — tipo: Generado / Entrega parcial / Entrega total
  Jornadas:      ["id","fecha_inicio","hora_inicio","fecha_fin","hora_fin","usuario","estado","ventas_totales","productos_vendidos","productos_cargados","productos_devueltos","ganancia","valorado"], // NUEVA CABECERA — "valorado" = valor en $ de lo cargado (cantidad × precio de venta)
  DetalleJornada: ["id","jornada_id","producto_id","producto_nombre","cantidad_cargada","cantidad_vendida","cantidad_devuelta","precio_compra","precio_venta","ganancia"] // NUEVA CABECERA
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

// Agrega, al final del encabezado, cualquier columna nueva que todavía no exista
// en una hoja ya creada anteriormente — sin tocar ni borrar las columnas ni las
// filas existentes. Así, actualizaciones como esta (que agregan "monto_original"
// a Cambios, "ultimo_envio_cierre" a Responsables o "jornada_id" a Ventas) no
// requieren recrear la hoja.
function asegurarColumnas_(hoja, columnas) {
  if (!hoja) return;
  const lastCol = Math.max(1, hoja.getLastColumn());
  const headers = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
  let cambiado = false;
  columnas.forEach(col => {
    if (headers.indexOf(col) === -1) { headers.push(col); cambiado = true; }
  });
  if (cambiado) hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
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
      case 'getCambiosMovimientos': r = getData(H_CAMBIOS_MOV); break; // Histórico de mini-cambios (entregas parciales/totales)
      case 'getCompras': r = getData(H_COMPRAS); break; // Obtener historial de compras
      case 'getResponsables': r = getData(H_RESPONSABLES); break;
      case 'getJornadaActiva': r = getJornadaActivaConDetalle(); break; // Jornada Activa (si existe) + su detalle por producto
      case 'getHistorialJornadas': r = getHistorialJornadasConValorado_(); break; // Historial de Jornadas (todas, activas y cerradas), con "valorado" recalculado y autocorregido
      case 'getDetalleJornada': r = getDetalleJornadaPorId_(p.jornada_id); break; // Detalle producto-por-producto de UNA jornada
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
      case 'agregarCambioManual': r = agregarCambioManual(req); break; // Ingresar un cambio pendiente manualmente
      case 'registrarEntregaCambio': r = registrarEntregaCambio(req); break; // Entregar un cambio (total o parcial / "mini cambio")
      case 'editarMovimientoCambio': r = editarMovimientoCambio(req); break; // Corregir fecha/cliente/monto/tipo de un registro del Histórico de Cambios
      case 'registrarCompra': r = registrarCompra(req); break; // Registro de gastos y stock
      case 'agregarResponsable': r = agregarResponsable(req); break;
      case 'editarResponsable': r = editarResponsable(req); break;
      case 'eliminarResponsable': r = eliminarEntidad(H_RESPONSABLES, req, 'Responsable'); break;
      case 'probarEnvioResponsable': r = probarEnvioResponsable(req); break;
      case 'importarDatos': r = importarDatosCompletos(req); break; // Restaurar copia de seguridad
      case 'enviarBackupAhora': r = enviarBackupManual(); break; // Forzar envío de backup a responsables
      case 'iniciarJornada': r = iniciarJornada(req); break; // Iniciar una nueva Jornada de Venta (carga Stock de Jornada)
      case 'cerrarJornada': r = cerrarJornada(req); break; // Cerrar la Jornada activa (devuelve sobrante al inventario global)
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

// ═══════════════════════════════════════════════════════════════
// JORNADAS DE VENTA (Stock Temporal / Stock de Maleta)
// ═══════════════════════════════════════════════════════════════
// A partir de esta actualización, TODAS las ventas se descuentan del
// "Stock de Jornada" (DetalleJornada.cantidad_vendida) y NUNCA directamente
// del inventario global (Productos.stock). El inventario global solo se
// mueve en dos momentos: (1) al iniciar una jornada, cuando se "carga" la
// maleta (se resta del inventario global), y (2) al cerrar la jornada,
// cuando el sobrante regresa automáticamente al inventario global.
// ═══════════════════════════════════════════════════════════════

// Devuelve el objeto (formato getData) de la Jornada con estado "Activa", o null si no existe.
// Nunca puede haber más de una jornada activa al mismo tiempo.
function obtenerJornadaActiva_() {
  const data = getData(H_JORNADAS).data || [];
  return data.find(j => j.estado === 'Activa') || null;
}

function getJornadaActivaConDetalle() {
  const jornada = obtenerJornadaActiva_();
  if (!jornada) return { status:'success', data: { jornada: null, detalle: [] } };
  const detalle = (getData(H_JORNADA_DET).data || []).filter(d => String(d.jornada_id) === String(jornada.id));
  return { status:'success', data: { jornada, detalle } };
}

function getDetalleJornadaPorId_(jornadaId) {
  if (!jornadaId) return { status:'error', message:'Falta el ID de la jornada.' };
  const detalleData = getData(H_JORNADA_DET);
  const detalle = (detalleData.data || []).filter(d => String(d.jornada_id) === String(jornadaId));
  return { status:'success', data: detalle };
}

// Suma cantidad_cargada × precio_venta (precio "congelado" al momento de cargar
// la jornada) de todas las líneas de DetalleJornada que pertenecen a una Jornada.
// Este es el "Valorado $": cuánto dinero representa, en precio de venta, todo
// lo que se cargó a la maleta — sin importar si ya se vendió o no.
function calcularValoradoDesdeDetalle_(jornadaId, detalleTodos) {
  let total = 0;
  detalleTodos.forEach(d => {
    if (String(d.jornada_id) === String(jornadaId)) {
      total += (parseFloat(d.cantidad_cargada) || 0) * (parseFloat(d.precio_venta) || 0);
    }
  });
  return total;
}

// Historial de Jornadas para la pantalla: en vez de confiar ciegamente en la
// columna "valorado" ya guardada en la hoja (que puede haber quedado vacía o
// en 0 en jornadas antiguas creadas antes de este cálculo), la recalcula
// siempre a partir de DetalleJornada — la fuente real de verdad — y de paso
// autocorrige la hoja si el valor guardado no coincide, para que la próxima
// lectura ya sea directa.
function getHistorialJornadasConValorado_() {
  const jornadasRes = getData(H_JORNADAS);
  if (jornadasRes.status !== 'success') return jornadasRes;

  const detalleTodos = (getData(H_JORNADA_DET).data) || [];
  const shJ = sh(H_JORNADAS);
  const headers = shJ.getRange(1, 1, 1, Math.max(1, shJ.getLastColumn())).getValues()[0];
  const idxValorado = headers.indexOf('valorado');

  const jornadas = jornadasRes.data.map((j, i) => {
    const valoradoCalculado = calcularValoradoDesdeDetalle_(j.id, detalleTodos);
    const valoradoGuardado = parseFloat(j.valorado) || 0;
    if (idxValorado > -1 && Math.abs(valoradoGuardado - valoradoCalculado) > 0.001) {
      shJ.getRange(i + 2, idxValorado + 1).setValue(valoradoCalculado); // autocorrección
    }
    j.valorado = valoradoCalculado;
    return j;
  });

  return { status:'success', data: jornadas };
}

// Inicia una nueva Jornada: valida que no exista otra activa, valida el stock
// en bodega de cada producto solicitado, crea la Jornada y su DetalleJornada
// (una fila por producto cargado, con cantidad_vendida/devuelta en 0 y el
// precio_compra/precio_venta "congelados" al momento de cargar), y descuenta
// de inmediato esa cantidad del inventario global.
function iniciarJornada(data) {
  if (obtenerJornadaActiva_()) {
    return { status:'error', message:'Ya existe una Jornada activa. Debes cerrarla antes de iniciar una nueva.' };
  }

  const solicitados = Array.isArray(data.productos) ? data.productos : [];
  const itemsValidos = solicitados
    .map(p => ({ producto_id: p.producto_id, cantidad: parseInt(p.cantidad) || 0 }))
    .filter(p => p.producto_id && p.cantidad > 0);

  if (!itemsValidos.length) {
    return { status:'error', message:'Debes indicar al menos un producto con una cantidad mayor a 0 para llevar a la jornada.' };
  }

  const shP = sh(H_PRODUCTOS);

  // Validar TODO antes de escribir nada (igual que registrarVenta), para nunca
  // dejar una jornada a medias si algún producto no tiene stock suficiente.
  const filasProducto = {};
  for (const it of itemsValidos) {
    const { row, idx } = findRow(shP, it.producto_id);
    if (!row) return { status:'error', message:'Uno de los productos seleccionados ya no existe.' };
    if (row[10] === false || row[10] === 'false') return { status:'error', message:`El producto "${row[1]}" fue eliminado del inventario.` };
    const stockBodega = parseInt(row[6]) || 0;
    if (it.cantidad > stockBodega) {
      return { status:'error', message:`Stock insuficiente en bodega de "${row[1]}". Disponible: ${stockBodega}, solicitado: ${it.cantidad}.` };
    }
    filasProducto[it.producto_id] = { row, idx };
  }

  const ahora = new Date();
  const jornadaId = uid();
  const shDJ = sh(H_JORNADA_DET);
  let totalCargado = 0;
  let totalValorado = 0; // valor en $ de todo lo cargado (cantidad × precio de venta)

  itemsValidos.forEach(it => {
    const { row, idx } = filasProducto[it.producto_id];
    const precioCompra = parseFloat(row[4]) || 0;
    const precioVenta = parseFloat(row[5]) || 0;
    // id, jornada_id, producto_id, producto_nombre, cantidad_cargada, cantidad_vendida, cantidad_devuelta, precio_compra, precio_venta, ganancia
    shDJ.appendRow([uid(), jornadaId, it.producto_id, row[1], it.cantidad, 0, 0, precioCompra, precioVenta, 0]);
    // Descontar de inmediato del inventario global (la maleta se "carga")
    const stockActual = parseInt(shP.getRange(idx + 1, 7).getValue()) || 0;
    shP.getRange(idx + 1, 7).setValue(stockActual - it.cantidad);
    totalCargado += it.cantidad;
    totalValorado += it.cantidad * precioVenta;
  });

  const shJ = sh(H_JORNADAS);
  // id, fecha_inicio, hora_inicio, fecha_fin, hora_fin, usuario, estado, ventas_totales, productos_vendidos, productos_cargados, productos_devueltos, ganancia, valorado
  shJ.appendRow([jornadaId, ahora, ahora, '', '', data.usuario || 'Sistema', 'Activa', 0, 0, totalCargado, 0, 0, totalValorado]);

  log(data.usuario, 'Jornada Iniciada', `${itemsValidos.length} producto(s), ${totalCargado} unidad(es) cargadas.`);

  return { status:'success', message:'Jornada iniciada con éxito. El stock cargado ya está disponible para la venta.', jornada_id: jornadaId };
}

// Cierra la Jornada activa: por cada producto de su DetalleJornada calcula lo
// devuelto (cargado - vendido), lo regresa al inventario global, calcula la
// ganancia de esa línea, y actualiza los totales de la Jornada. No permite
// cerrar una jornada que ya está cerrada. Devuelve un resumen completo para
// mostrar en pantalla e imprimir en PDF.
function cerrarJornada(data) {
  const shJ = sh(H_JORNADAS); const shDJ = sh(H_JORNADA_DET); const shP = sh(H_PRODUCTOS);
  const { row: rowJ, idx: idxJ } = findRow(shJ, data.id);
  if (!rowJ) return { status:'error', message:'Jornada no encontrada.' };
  if (String(rowJ[6]) !== 'Activa') return { status:'error', message:'Esta jornada ya fue cerrada anteriormente.' };

  const detVals = shDJ.getDataRange().getValues();
  let totalCargado = 0, totalVendido = 0, totalDevuelto = 0, montoVendido = 0, gananciaTotal = 0;
  const detalleResumen = [];

  for (let i = 1; i < detVals.length; i++) {
    if (String(detVals[i][1]) !== String(data.id)) continue;

    const productoId   = detVals[i][2];
    const productoNom  = detVals[i][3];
    const cargada      = parseInt(detVals[i][4]) || 0;
    const vendida       = parseInt(detVals[i][5]) || 0;
    const devuelta      = Math.max(0, cargada - vendida);
    const precioCompra  = parseFloat(detVals[i][7]) || 0;
    const precioVenta   = parseFloat(detVals[i][8]) || 0;
    const gananciaLinea = vendida * (precioVenta - precioCompra);

    shDJ.getRange(i + 1, 7).setValue(devuelta);       // cantidad_devuelta
    shDJ.getRange(i + 1, 10).setValue(gananciaLinea); // ganancia

    // Regresar el sobrante al inventario global
    if (devuelta > 0) {
      const { idx: idxP } = findRow(shP, productoId);
      if (idxP > -1) {
        const stockActual = parseInt(shP.getRange(idxP + 1, 7).getValue()) || 0;
        shP.getRange(idxP + 1, 7).setValue(stockActual + devuelta);
      }
    }

    totalCargado += cargada; totalVendido += vendida; totalDevuelto += devuelta;
    montoVendido += vendida * precioVenta; gananciaTotal += gananciaLinea;

    detalleResumen.push({
      producto_id: productoId, producto_nombre: productoNom,
      cantidad_cargada: cargada, cantidad_vendida: vendida, cantidad_devuelta: devuelta,
      precio_venta: precioVenta, ingresos: vendida * precioVenta, ganancia: gananciaLinea
    });
  }

  const ahora = new Date();
  shJ.getRange(idxJ + 1, 4).setValue(ahora);          // fecha_fin
  shJ.getRange(idxJ + 1, 5).setValue(ahora);          // hora_fin
  shJ.getRange(idxJ + 1, 7).setValue('Cerrada');      // estado
  shJ.getRange(idxJ + 1, 8).setValue(montoVendido);   // ventas_totales
  shJ.getRange(idxJ + 1, 9).setValue(totalVendido);   // productos_vendidos
  shJ.getRange(idxJ + 1, 10).setValue(totalCargado);  // productos_cargados
  shJ.getRange(idxJ + 1, 11).setValue(totalDevuelto); // productos_devueltos
  shJ.getRange(idxJ + 1, 12).setValue(gananciaTotal); // ganancia

  log(data.usuario, 'Jornada Cerrada', `Cargado: ${totalCargado}, Vendido: ${totalVendido}, Devuelto: ${totalDevuelto}, Ventas: $${montoVendido.toFixed(2)}, Ganancia: $${gananciaTotal.toFixed(2)}`);

  return {
    status:'success',
    message:'Jornada cerrada con éxito. El stock sobrante fue devuelto al inventario global.',
    resumen: {
      id: data.id,
      fecha_inicio: rowJ[1] ? new Date(rowJ[1]).toISOString() : '',
      hora_inicio: rowJ[2] ? new Date(rowJ[2]).toISOString() : '',
      fecha_fin: ahora.toISOString(),
      hora_fin: ahora.toISOString(),
      usuario: rowJ[5],
      productos_cargados: totalCargado,
      productos_vendidos: totalVendido,
      productos_devueltos: totalDevuelto,
      ventas_totales: montoVendido,
      ganancia: gananciaTotal,
      detalle: detalleResumen
    }
  };
}

function registrarVenta(data) {
  const shV = sh(H_VENTAS); const shD = sh(H_VENTA_DET); const shC = sh(H_CAMBIOS);
  const shDJ = sh(H_JORNADA_DET);

  // Regla clave de Jornadas: NUNCA se vende directamente del inventario global.
  // Debe existir una Jornada activa; todas las ventas descuentan del Stock de Jornada.
  const jornadaActiva = obtenerJornadaActiva_();
  if (!jornadaActiva) {
    return { status:'error', message:'No hay una Jornada activa. Debes iniciar una jornada antes de registrar ventas.' };
  }

  let subtotal = 0;
  // Primero se valida TODO el carrito contra el Stock de Jornada real (incluyendo
  // cantidades repetidas del mismo producto) antes de escribir nada, para nunca
  // dejar una venta a medias.
  const acumulado = {}; // producto_id -> cantidad total pedida en este carrito
  const detalleJornadaVals = shDJ.getDataRange().getValues();
  const filaDetalleJornadaPorProducto = {}; // producto_id -> índice de fila (0-based, alineado con detalleJornadaVals)
  for (let i = 1; i < detalleJornadaVals.length; i++) {
    if (String(detalleJornadaVals[i][1]) === String(jornadaActiva.id)) {
      filaDetalleJornadaPorProducto[String(detalleJornadaVals[i][2])] = i;
    }
  }

  for (const item of data.items) {
    const cant = parseInt(item.cantidad) || 0;
    if (cant <= 0) return {status:'error', message:`Cantidad inválida para ${item.producto_id}.`};
    const filaIdx = filaDetalleJornadaPorProducto[String(item.producto_id)];
    if (filaIdx === undefined) return {status:'error', message:`Ese producto no fue cargado en la jornada actual, así que no se puede vender.`};
    const filaDJ = detalleJornadaVals[filaIdx];
    const nombreProd = filaDJ[3];
    acumulado[item.producto_id] = (acumulado[item.producto_id] || 0) + cant;
    const cargada = parseInt(filaDJ[4]) || 0;
    const vendidaPrevia = parseInt(filaDJ[5]) || 0;
    const disponibleJornada = cargada - vendidaPrevia;
    if (disponibleJornada < acumulado[item.producto_id]) {
      return {status:'warning', message:`Stock de jornada insuficiente de "${nombreProd}". Disponible: ${disponibleJornada}, solicitado: ${acumulado[item.producto_id]}.`};
    }
    item._rowIdxJornada = filaIdx; item._nombre = nombreProd; item._subtotal = (parseFloat(item.precio_unitario)*cant);
    subtotal += item._subtotal;
  }
  const total = subtotal;
  const ventaId = uid();
  const nombreCliente = data.cliente_nombre && String(data.cliente_nombre).trim() !== '' ? String(data.cliente_nombre).trim() : 'N/A';

  shV.appendRow([ventaId, new Date(), nombreCliente, subtotal, 0, 0, total, data.pago_con||0, data.cambio||0, data.usuario||'Sistema', 'completada', jornadaActiva.id]);

  for (const item of data.items) {
    shD.appendRow([uid(), ventaId, item.producto_id, item._nombre, parseInt(item.cantidad), parseFloat(item.precio_unitario), 0, item._subtotal]);
    // Descontar del Stock de Jornada — el inventario global NO se toca aquí.
    const filaSheetIdx = item._rowIdxJornada;
    const vendidaActual = parseInt(shDJ.getRange(filaSheetIdx + 1, 6).getValue()) || 0;
    shDJ.getRange(filaSheetIdx + 1, 6).setValue(vendidaActual + parseInt(item.cantidad));
  }

  if (data.pendiente && data.cambio > 0) {
    const cambioId = uid();
    shC.appendRow([cambioId, ventaId, new Date(), nombreCliente, data.cambio, 'Pendiente', '', data.usuario||'Sistema', data.cambio]);
    try { sh(H_CAMBIOS_MOV).appendRow([uid(), cambioId, new Date(), nombreCliente, data.cambio, 'Generado', data.usuario||'Sistema']); } catch(e) {}
  }

  try { avisarVentaGrande_(total, nombreCliente, ventaId); } catch(e) {}

  return { status:'success', message:'Venta registrada con éxito.' };
}

// Anula (elimina) una venta: guarda copia completa + justificación en la Papelera,
// devuelve el stock que se había descontado (al Stock de Jornada si la jornada de
// esa venta sigue activa, o al inventario global si no hay jornada asociada o esa
// jornada ya fue cerrada), borra cualquier "cambio pendiente" ligado a esa venta, y
// elimina la venta de Ventas/VentaDetalle para que deje de contarse en Reportes,
// Resúmenes y el total recaudado.
function anularVenta(data) {
  const shV = sh(H_VENTAS); const shD = sh(H_VENTA_DET); const shP = sh(H_PRODUCTOS); const shPap = sh(H_PAPELERA); const shC = sh(H_CAMBIOS);
  const shDJ = sh(H_JORNADA_DET);
  const { row: rowV } = findRow(shV, data.id);
  if (!rowV) return { status:'error', message:'Venta no encontrada.' };
  if (!data.motivo || !String(data.motivo).trim()) return { status:'error', message:'Debes indicar un motivo de la anulación.' };

  // ¿A qué Jornada pertenecía esta venta (si a alguna)? La columna jornada_id
  // se agregó de forma aditiva al final de Ventas, así que puede no existir en
  // filas muy antiguas — en ese caso se trata igual que antes de esta actualización.
  const jornadaIdVenta = rowV.length > 11 ? String(rowV[11] || '') : '';
  let jornadaDestinoActiva = null;
  if (jornadaIdVenta) {
    const jornadaData = (getData(H_JORNADAS).data || []).find(j => String(j.id) === jornadaIdVenta);
    if (jornadaData && jornadaData.estado === 'Activa') jornadaDestinoActiva = jornadaData;
  }

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

  // Devolver el stock vendido: si la Jornada de esta venta SIGUE ACTIVA, se
  // devuelve al Stock de Jornada (se resta de cantidad_vendida, sin tocar el
  // inventario global, que nunca se tocó al vender). En cualquier otro caso
  // (venta sin jornada asociada, o cuya jornada ya fue cerrada) se devuelve
  // directamente al inventario global — mismo comportamiento que antes de
  // esta actualización.
  if (jornadaDestinoActiva) {
    const detJornadaVals = shDJ.getDataRange().getValues();
    itemsVenta.forEach(it => {
      for (let i = 1; i < detJornadaVals.length; i++) {
        if (String(detJornadaVals[i][1]) === jornadaIdVenta && String(detJornadaVals[i][2]) === String(it.producto_id)) {
          const vendidaActual = parseInt(shDJ.getRange(i + 1, 6).getValue()) || 0;
          const nuevaVendida = Math.max(0, vendidaActual - (parseInt(it.cantidad) || 0));
          shDJ.getRange(i + 1, 6).setValue(nuevaVendida);
          break;
        }
      }
    });
  } else {
    itemsVenta.forEach(it => {
      const { idx: idxP } = findRow(shP, it.producto_id);
      if (idxP > -1) {
        const stockActual = parseInt(shP.getRange(idxP + 1, 7).getValue()) || 0;
        shP.getRange(idxP + 1, 7).setValue(stockActual + (parseInt(it.cantidad) || 0));
      }
    });
  }

  // Guardar copia completa (venta + items + motivo) en la Papelera antes de borrar
  const datosOriginales = {
    venta: { id: rowV[0], fecha: rowV[1], cliente_nombre: rowV[2], subtotal: rowV[3], descuento: rowV[4], impuesto: rowV[5], total: rowV[6], pago_con: rowV[7], cambio: rowV[8], usuario: rowV[9], estado: rowV[10] },
    items: itemsVenta,
    motivo: String(data.motivo).trim()
  };
  shPap.appendRow([uid(), 'Venta Anulada', JSON.stringify(datosOriginales), new Date(), data.usuario || 'Sistema']);

  // Eliminar cualquier "cambio pendiente" ligado a esta venta, ya no aplica
  // (y su historial de mini-cambios/entregas, si tenía alguno).
  const cambiosVals = shC.getDataRange().getValues();
  const idsCambiosABorrar = [];
  for (let i = cambiosVals.length - 1; i >= 1; i--) {
    if (String(cambiosVals[i][1]) === String(data.id)) { idsCambiosABorrar.push(String(cambiosVals[i][0])); shC.deleteRow(i + 1); }
  }
  if (idsCambiosABorrar.length) {
    try {
      const shCMov = sh(H_CAMBIOS_MOV);
      const movVals = shCMov.getDataRange().getValues();
      for (let i = movVals.length - 1; i >= 1; i--) {
        if (idsCambiosABorrar.indexOf(String(movVals[i][1])) > -1) shCMov.deleteRow(i + 1);
      }
    } catch(e) {}
  }

  // Eliminar las líneas de detalle (de mayor a menor índice para no desordenar filas)
  filasDetalleABorrar.sort((a, b) => b - a).forEach(i => shD.deleteRow(i));

  // Eliminar la venta (así deja de sumar en Reportes/Resúmenes/total recaudado)
  const { idx: idxVentaFinal } = findRow(shV, data.id);
  if (idxVentaFinal > -1) shV.deleteRow(idxVentaFinal + 1);

  log(data.usuario, 'Venta Anulada', `Motivo: ${data.motivo}. Cliente: ${rowV[2]}, Total: $${(parseFloat(rowV[6])||0).toFixed(2)}`);

  return { status:'success', message:'Venta eliminada. El stock fue devuelto correctamente.' };
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

// Marca un cambio como entregado POR COMPLETO de una sola vez (botón rápido
// "Entregado" cuando no se necesita hacer una entrega parcial).
function marcarCambioPagado(data) {
  const hoja = sh(H_CAMBIOS); const {idx, row} = findRow(hoja, data.id);
  if (idx < 0) return {status: 'error', message: 'Registro no encontrado.'};
  if (String(row[5]) !== 'Pendiente') return {status:'error', message:'Este cambio ya fue entregado.'};
  const pendienteActual = parseFloat(row[4]) || 0;
  hoja.getRange(idx+1, 5).setValue(0);
  hoja.getRange(idx+1, 6).setValue('Pagado');
  hoja.getRange(idx+1, 7).setValue(new Date().toISOString());
  if (pendienteActual > 0) {
    try { sh(H_CAMBIOS_MOV).appendRow([uid(), data.id, new Date(), row[3], pendienteActual, 'Entrega total', data.usuario||'Sistema']); } catch(e) {}
  }
  log(data.usuario, 'Cambio Entregado (Total)', `Cliente: ${row[3]}, Entregado: $${pendienteActual.toFixed(2)}`);
  return {status: 'success', message: 'Cambio marcado como pagado.'};
}

// Ingresa manualmente un nuevo "cambio pendiente" que no viene ligado a una
// venta (por ejemplo, un vuelto que se quedó a deber en un ajuste de caja).
function agregarCambioManual(data) {
  const cliente = data.cliente_nombre && String(data.cliente_nombre).trim() !== '' ? String(data.cliente_nombre).trim() : '';
  if (!cliente) return {status:'error', message:'Indica el nombre del cliente.'};
  const monto = parseFloat(data.monto);
  if (isNaN(monto) || monto <= 0) return {status:'error', message:'El monto debe ser mayor a 0.'};

  const id = uid();
  sh(H_CAMBIOS).appendRow([id, '', new Date(), cliente, monto, 'Pendiente', '', data.usuario||'Sistema', monto]);
  try { sh(H_CAMBIOS_MOV).appendRow([uid(), id, new Date(), cliente, monto, 'Generado', data.usuario||'Sistema']); } catch(e) {}
  log(data.usuario, 'Cambio Manual Agregado', `Cliente: ${cliente}, Monto: $${monto.toFixed(2)}`);
  return {status:'success', message:'Cambio pendiente agregado con éxito.'};
}

// Registra la entrega de un cambio pendiente, ya sea TOTAL o PARCIAL ("mini
// cambio"). Si el monto entregado es menor al pendiente, el registro sigue
// "Pendiente" con el monto restante, y queda un "mini cambio" para seguir
// entregando después; si cubre todo lo pendiente, se marca "Pagado".
function registrarEntregaCambio(data) {
  const hoja = sh(H_CAMBIOS); const {row, idx} = findRow(hoja, data.id);
  if (!row) return {status:'error', message:'Registro no encontrado.'};
  if (String(row[5]) !== 'Pendiente') return {status:'error', message:'Este cambio ya fue entregado por completo.'};

  const pendienteActual = parseFloat(row[4]) || 0;
  const montoEntrega = parseFloat(data.monto);
  if (isNaN(montoEntrega) || montoEntrega <= 0) return {status:'error', message:'Indica un monto de entrega válido (mayor a 0).'};
  if (montoEntrega > pendienteActual + 0.009) return {status:'error', message:`El monto no puede superar lo pendiente ($${pendienteActual.toFixed(2)}).`};

  const nuevoPendiente = Math.max(0, pendienteActual - montoEntrega);
  const esTotal = nuevoPendiente <= 0.009;

  hoja.getRange(idx+1, 5).setValue(esTotal ? 0 : nuevoPendiente);
  if (esTotal) {
    hoja.getRange(idx+1, 6).setValue('Pagado');
    hoja.getRange(idx+1, 7).setValue(new Date().toISOString());
  }

  try { sh(H_CAMBIOS_MOV).appendRow([uid(), data.id, new Date(), row[3], montoEntrega, esTotal ? 'Entrega total' : 'Entrega parcial', data.usuario||'Sistema']); } catch(e) {}
  log(data.usuario, esTotal ? 'Cambio Entregado (Total)' : 'Cambio Entregado (Parcial)', `Cliente: ${row[3]}, Entregado: $${montoEntrega.toFixed(2)}, Restante: $${(esTotal?0:nuevoPendiente).toFixed(2)}`);

  return {status:'success', message: esTotal ? 'Cambio entregado por completo.' : `Entrega parcial (mini cambio) registrada. Queda pendiente $${nuevoPendiente.toFixed(2)}.`};
}

// Corrige un registro ya existente del Histórico de Cambios (CambiosMovimientos):
// fecha, cliente, monto y/o tipo. Es solo para arreglar datos mal ingresados
// (p.ej. la fecha quedó mal, o en realidad sí se entregó el dinero) — NO
// recalcula el "Cambio Pendiente" original (hoja Cambios) al que pertenece,
// así que si se corrige aquí el monto de una entrega, revisa también el
// cambio pendiente correspondiente si hiciera falta.
function editarMovimientoCambio(data) {
  const hoja = sh(H_CAMBIOS_MOV); const {idx} = findRow(hoja, data.id);
  if (idx < 0) return {status:'error', message:'Movimiento no encontrado.'};

  if (data.cliente_nombre !== undefined) {
    const cliente = String(data.cliente_nombre).trim();
    if (!cliente) return {status:'error', message:'Indica el nombre del cliente.'};
    hoja.getRange(idx+1, 4).setValue(cliente);
  }
  if (data.fecha !== undefined) {
    const f = new Date(data.fecha);
    if (isNaN(f.getTime())) return {status:'error', message:'Fecha inválida.'};
    hoja.getRange(idx+1, 3).setValue(f);
  }
  if (data.monto !== undefined) {
    const m = parseFloat(data.monto);
    if (isNaN(m) || m < 0) return {status:'error', message:'El monto debe ser un número mayor o igual a 0.'};
    hoja.getRange(idx+1, 5).setValue(m);
  }
  if (data.tipo !== undefined) {
    if (!['Generado','Entrega parcial','Entrega total'].includes(data.tipo)) return {status:'error', message:'Tipo de movimiento inválido.'};
    hoja.getRange(idx+1, 6).setValue(data.tipo);
  }

  log(data.usuario, 'Movimiento de Cambio Editado', `Id: ${data.id}`);
  return {status:'success', message:'Movimiento actualizado con éxito.'};
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
  sh(H_RESPONSABLES).appendRow([uid(), data.email.trim(), data.nombre || '', data.reportes || '', parseFloat(data.umbral_venta_grande)||100, data.activo !== false, '', '', '', new Date(), '', '']);
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
  [H_CATEGORIAS,H_PRODUCTOS,H_VENTAS,H_VENTA_DET,H_PAPELERA,H_ACTIVIDAD,H_RESPONSABLES,H_CAMBIOS,H_COMPRAS,H_CAMBIOS_MOV,H_JORNADAS,H_JORNADA_DET].forEach(crearHoja);
  // Migración de columnas nuevas en hojas que ya existían antes de esta actualización
  // (cierre de caja, mini-cambios/entregas parciales, y Jornadas de Venta), sin borrar
  // datos existentes.
  try { asegurarColumnas_(sh(H_RESPONSABLES), ["ultimo_envio_cierre"]); } catch(e) {}
  try { asegurarColumnas_(sh(H_CAMBIOS), ["monto_original"]); } catch(e) {}
  try { asegurarColumnas_(sh(H_VENTAS), ["jornada_id"]); } catch(e) {}
  try { asegurarColumnas_(sh(H_JORNADAS), ["valorado"]); } catch(e) {}
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
    const jornadas = getData(H_JORNADAS).data || [];
    const detalleJornadas = getData(H_JORNADA_DET).data || [];
    return { status:'success', data: { productos, categorias, ventas, detalle_ventas: detalle, jornadas, detalle_jornadas: detalleJornadas, fecha_respaldo: new Date().toISOString() } };
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
    if (Array.isArray(backup.jornadas)) resumen.push(`${reemplazarHoja_(H_JORNADAS, backup.jornadas, HDR.Jornadas)} jornadas`);
    if (Array.isArray(backup.detalle_jornadas)) resumen.push(`${reemplazarHoja_(H_JORNADA_DET, backup.detalle_jornadas, HDR.DetalleJornada)} líneas de detalle de jornada`);

    if (!resumen.length) return { status:'error', message:'El archivo no contiene datos reconocibles (productos, categorías, ventas, detalle_ventas, jornadas o detalle_jornadas).' };

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
// 10=fecha_creado  11=ultimo_envio_backup  12=ultimo_envio_cierre
const RESP_COL = { ultimo_envio_mensual: 7, ultimo_envio_anual: 8, ultimo_envio_stock: 9, ultimo_envio_backup: 11, ultimo_envio_cierre: 12 };

// Compara si dos fechas caen en el mismo día calendario.
function mismoDia_(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
}

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
  try { enviarCierreCaja_(hoy); } catch(e) { log('Sistema','Error cierre de caja', e.message); }
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
  const jornadas = getData(H_JORNADAS).data || [];
  const detalleJornadas = getData(H_JORNADA_DET).data || [];
  return Utilities.newBlob(
    JSON.stringify({ productos, ventas, detalle_ventas: detalle, categorias, jornadas, detalle_jornadas: detalleJornadas, fecha_respaldo: new Date().toISOString() }, null, 2),
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

// 🧾 Cierre de caja diario — se envía a la misma hora que todo lo demás
// (junto con la copia de seguridad y las alertas de tareasDiarias). Incluye
// un listado en texto de las ventas del día, los cambios que se entregaron
// (descritos uno por uno) y los totales de ventas, cambios entregados y
// ganancia estimada del día.
function contenidoCierreCaja_(hoy) {
  const fechaTxt = formatoFecha_(hoy);
  const ventasHoy = (getData(H_VENTAS).data || []).filter(v => mismoDia_(new Date(v.fecha), hoy));
  const idsVentasHoy = {}; ventasHoy.forEach(v => idsVentasHoy[String(v.id)] = true);
  const lineasHoy = (getData(H_VENTA_DET).data || []).filter(d => idsVentasHoy[String(d.venta_id)]);

  const costoPorProducto = {};
  (getData(H_PRODUCTOS).data || []).forEach(p => { costoPorProducto[p.id] = parseFloat(p.precio_compra) || 0; });

  let totalVentas = 0, totalGanancia = 0;
  ventasHoy.forEach(v => totalVentas += parseFloat(v.total) || 0);
  lineasHoy.forEach(l => {
    const cant = parseInt(l.cantidad) || 0;
    const precioVenta = parseFloat(l.precio_unitario) || 0;
    const precioCompra = costoPorProducto[l.producto_id] !== undefined ? costoPorProducto[l.producto_id] : precioVenta;
    totalGanancia += (precioVenta - precioCompra) * cant;
  });

  // Cambios entregados hoy (movimientos de tipo "Entrega parcial" o "Entrega total")
  const entregasHoy = (getData(H_CAMBIOS_MOV).data || []).filter(m =>
    mismoDia_(new Date(m.fecha), hoy) && (m.tipo === 'Entrega parcial' || m.tipo === 'Entrega total'));
  const totalCambiosEntregados = entregasHoy.reduce((s, m) => s + (parseFloat(m.monto) || 0), 0);

  const hora_ = d => new Date(d).toLocaleTimeString('es-SV', { hour: '2-digit', minute: '2-digit' });

  const listadoVentasTexto = ventasHoy.length
    ? ventasHoy.map(v => `${hora_(v.fecha)} — ${v.cliente_nombre || 'N/A'} — $${(parseFloat(v.total) || 0).toFixed(2)}`).join('\n')
    : 'Sin ventas registradas hoy.';
  const listadoCambiosTexto = entregasHoy.length
    ? entregasHoy.map(m => `${hora_(m.fecha)} — ${m.cliente_nombre || 'N/A'} — $${(parseFloat(m.monto) || 0).toFixed(2)} (${m.tipo})`).join('\n')
    : 'No se entregaron cambios hoy.';

  const texto = `CIERRE DE CAJA — ${fechaTxt}
========================================
VENTAS DEL DÍA (${ventasHoy.length})
${listadoVentasTexto}

CAMBIOS ENTREGADOS (${entregasHoy.length})
${listadoCambiosTexto}

----------------------------------------
Total vendido:             $${totalVentas.toFixed(2)}
Total cambios entregados:  $${totalCambiosEntregados.toFixed(2)}
Ganancia estimada del día: $${totalGanancia.toFixed(2)}
========================================`;

  const filasVentasHtml = ventasHoy.map(v =>
    `<tr><td style="padding:4px 8px;">${hora_(v.fecha)}</td><td style="padding:4px 8px;">${v.cliente_nombre || 'N/A'}</td><td style="padding:4px 8px;">$${(parseFloat(v.total) || 0).toFixed(2)}</td></tr>`
  ).join('') || '<tr><td colspan="3" style="padding:4px 8px;color:#888;">Sin ventas registradas hoy.</td></tr>';

  const filasCambiosHtml = entregasHoy.map(m =>
    `<tr><td style="padding:4px 8px;">${hora_(m.fecha)}</td><td style="padding:4px 8px;">${m.cliente_nombre || 'N/A'}</td><td style="padding:4px 8px;">$${(parseFloat(m.monto) || 0).toFixed(2)}</td><td style="padding:4px 8px;">${m.tipo}</td></tr>`
  ).join('') || '<tr><td colspan="4" style="padding:4px 8px;color:#888;">No se entregaron cambios hoy.</td></tr>';

  const html = `
    <h3>🧾 Cierre de Caja — ${fechaTxt}</h3>
    <p>Ventas: <b>${ventasHoy.length}</b> &nbsp;|&nbsp; Total vendido: <b>$${totalVentas.toFixed(2)}</b> &nbsp;|&nbsp; Ganancia estimada: <b>$${totalGanancia.toFixed(2)}</b> &nbsp;|&nbsp; Cambios entregados: <b>$${totalCambiosEntregados.toFixed(2)}</b></p>
    <h4>Ventas del día</h4>
    <table border="1" cellspacing="0" style="border-collapse:collapse;width:100%;max-width:520px;"><tr style="background:#f4f4f4;"><th style="padding:4px 8px;">Hora</th><th style="padding:4px 8px;">Cliente</th><th style="padding:4px 8px;">Total</th></tr>${filasVentasHtml}</table>
    <h4>Cambios entregados (descritos)</h4>
    <table border="1" cellspacing="0" style="border-collapse:collapse;width:100%;max-width:520px;"><tr style="background:#f4f4f4;"><th style="padding:4px 8px;">Hora</th><th style="padding:4px 8px;">Cliente</th><th style="padding:4px 8px;">Monto</th><th style="padding:4px 8px;">Tipo</th></tr>${filasCambiosHtml}</table>
    <p><b>Resumen del día:</b><br>Total vendido: $${totalVentas.toFixed(2)}<br>Total cambios entregados: $${totalCambiosEntregados.toFixed(2)}<br>Ganancia estimada: $${totalGanancia.toFixed(2)}</p>
    <p style="color:#888;font-size:12px;">Listado en texto plano (adjunto abajo para copiar/imprimir):</p>
    <pre style="background:#f7f7f7;padding:10px;border-radius:6px;white-space:pre-wrap;font-size:12px;">${texto}</pre>
  `;

  return { asunto: `🧾 Cierre de Caja — ${fechaTxt}`, html, texto };
}

function enviarCierreCaja_(hoy) {
  const responsables = (getData(H_RESPONSABLES).data || []).filter(r =>
    activoResp_(r) && tieneReporte_(r, 'cierre_caja') && !yaEnviadoHoy_(r.ultimo_envio_cierre, hoy));
  if (!responsables.length) return;
  const contenido = contenidoCierreCaja_(hoy);
  responsables.forEach(r => {
    try {
      MailApp.sendEmail({ to: r.email, subject: contenido.asunto, htmlBody: contenido.html, body: contenido.texto });
      marcarEnvioResponsable_(r.id, 'ultimo_envio_cierre', hoy);
    } catch(e) {}
  });
}

function construirCSVVentas_(ventas) {
  const headers = ['id','fecha','cliente_nombre','subtotal','total','pago_con','cambio','usuario','estado'];
  const filas = [headers.join(',')];
  ventas.forEach(v => filas.push(headers.map(h => `"${String(v[h] || '').replace(/"/g, '""')}"`).join(',')));
  return filas.join('\n');
}

const TIPOS_REPORTE_TODOS = ['mensual', 'anual', 'ventas_grandes', 'stock_bajo', 'backup_json', 'backup_csv', 'cierre_caja'];

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
    case 'cierre_caja': {
      const c = contenidoCierreCaja_(hoy);
      return { asunto: c.asunto, html: c.html, adjuntos: [] };
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
