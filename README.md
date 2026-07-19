# ERP POS LITE — Manual de Usuario y Documentación Técnica

Sistema de Punto de Venta e Inventario · Google Apps Script + Google Sheets + HTML/JavaScript

> Versión: julio 2026 · Incluye: validación de stock en 3 colores, modo sin conexión mejorado y nueva columna **P.Compra** en Inventario.

## Índice

1. [Introducción](#1-introducción)
2. [Arquitectura del sistema](#2-arquitectura-del-sistema)
3. [Manual de uso](#3-manual-de-uso)
   - [3.1 Punto de Venta](#31-punto-de-venta)
   - [3.2 Inventario y Stock](#32-inventario-y-stock)
   - [3.3 Registrar Producto y Categorías](#33-registrar-producto-y-categorías)
   - [3.4 Compras (Reabastecimiento)](#34-compras-reabastecimiento)
   - [3.5 Cambios Pendientes](#35-cambios-pendientes)
   - [3.6 Reportes y Resúmenes](#36-reportes-y-resúmenes)
   - [3.7 Responsables](#37-responsables)
   - [3.8 Papelera](#38-papelera)
   - [3.9 Ajustes](#39-ajustes)
4. [Modo sin conexión (offline)](#4-modo-sin-conexión-offline)
5. [Estructura de datos (Google Sheets)](#5-estructura-de-datos-google-sheets)
6. [Funciones del backend (codigo.gs)](#6-funciones-del-backend-codigogs)
7. [Funciones clave del frontend (index.html)](#7-funciones-clave-del-frontend-indexhtml)
8. [Validaciones y correcciones aplicadas](#8-validaciones-y-correcciones-aplicadas)
9. [Despliegue de esta actualización](#9-despliegue-de-esta-actualización)

---

## 1. Introducción

ERP POS LITE es un sistema de Punto de Venta e Inventario pensado para negocios pequeños y medianos. Funciona completamente en el navegador (sin instalar nada) y usa una Google Sheet como base de datos, administrada por un backend en Google Apps Script.

Este manual cubre dos cosas a la vez: cómo usar el sistema día a día (con capturas de pantalla reales de la aplicación) y cómo funciona por dentro (arquitectura, funciones del backend y del frontend, y la estructura de las hojas de cálculo que actúan como base de datos).

**Módulos del sistema:**

- **Punto de Venta**: registrar ventas con carrito, control de vuelto y validación de stock en tiempo real.
- **Inventario y Stock**: listado de productos con estado de stock, precios de compra/venta y valor total.
- **Registro de Productos y Categorías**: alta de nuevos productos y categorías.
- **Compras (Reabastecimiento)**: registrar entradas de mercadería y su costo.
- **Cambios Pendientes**: vueltos que quedaron a deber al cliente.
- **Reportes y Resúmenes**: historial de ventas, productos más vendidos y márgenes de ganancia.
- **Responsables**: correos que reciben alertas automáticas (stock bajo, ventas grandes, resúmenes).
- **Papelera**: respaldo de lo eliminado (productos, categorías, ventas anuladas).
- **Ajustes**: modo sin conexión, copias de seguridad e inicialización de la base de datos.

---

## 2. Arquitectura del sistema

El sistema tiene tres capas: lo que el usuario ve en el navegador, el backend en Google Apps Script que procesa las solicitudes, y Google Sheets como base de datos. Cuando no hay internet, el navegador guarda temporalmente la información en su propio almacenamiento local (`localStorage`) hasta poder sincronizar.

<p align="center"><img src="manual_assets/arquitectura.png" width="700" alt="Arquitectura general del sistema"></p>
<p align="center"><i>Figura 1. Arquitectura general del sistema.</i></p>

**Cómo viaja una solicitud:**

- El navegador carga `index.html`, que contiene toda la interfaz (HTML, CSS y JavaScript) en un solo archivo.
- Para leer datos (ver inventario, ventas, etc.) el frontend hace una petición **GET** a la URL del Web App de Apps Script, indicando una acción (por ejemplo `?action=getInventario`).
- Para crear o modificar datos (registrar una venta, editar un producto) el frontend hace una petición **POST** con un JSON que incluye la acción y los datos.
- `codigo.gs` recibe la solicitud en `doGet()` o `doPost()`, la enruta según la acción, lee/escribe en la hoja de cálculo correspondiente y responde en formato JSON.
- Si el navegador no tiene conexión y el modo offline está activado, la venta se guarda localmente y se sincroniza más tarde (ver sección 4).

---

## 3. Manual de uso

### 3.1 Punto de Venta

Es la pantalla principal. Del lado izquierdo se muestra la cuadrícula de productos disponibles; del lado derecho, el carrito de la venta actual.

<p align="center"><img src="manual_assets/pos.png" width="750" alt="Punto de Venta con semáforo de stock"></p>
<p align="center"><i>Figura 2. Punto de Venta: cada producto muestra su nivel de stock en 3 colores, y el carrito muestra la disponibilidad de cada línea.</i></p>

**Semáforo de stock:** cada tarjeta de producto incluye una etiqueta de color entre la imagen y el precio:

| Color | Significado |
|---|---|
| 🟢 Verde | Stock normal, por encima del mínimo configurado. |
| 🟠 Amarillo | Stock bajo (igual o menor al mínimo de alerta). |
| 🔴 Rojo | Sin stock. La tarjeta se ve atenuada y no se puede agregar al carrito. |

Dentro del carrito, cada producto muestra "Stock disponible" con el mismo color. Si intentas agregar más unidades de las que hay disponibles (ya sea tocando el producto otra vez o presionando "+" en el carrito), el sistema muestra un aviso **"Stock insuficiente"** y no permite continuar — esto aplica tanto con internet como sin conexión.

- **Cliente obligatorio**: no se puede finalizar una venta sin escribir el nombre del cliente.
- **Registrar Cobro Exitoso**: registra la venta como completada.
- **Queda Pendiente**: registra la venta pero deja el vuelto como una deuda pendiente en el módulo Cambios.

### 3.2 Inventario y Stock

Lista todos los productos activos con su información clave. La tabla ahora incluye la columna **P.Compra** junto a P.Venta, para poder comparar de un vistazo el costo y el precio de venta de cada producto.

<p align="center"><img src="manual_assets/inventario.png" width="750" alt="Inventario y Stock con columna P.Compra"></p>
<p align="center"><i>Figura 3. Inventario y Stock, con la nueva columna P.Compra.</i></p>

| Columna | Descripción |
|---|---|
| Img | Foto del producto (o un ícono genérico si no tiene imagen). |
| Nombre | Nombre del producto. |
| Categoría | Categoría a la que pertenece. |
| Stock | Unidades disponibles, junto al mínimo configurado. |
| Estado | Etiqueta OK / Stock bajo / Sin stock, con color. |
| **P.Compra** | Precio al que se compró el producto (costo). |
| P.Venta | Precio al que se vende al cliente. |
| Valor en $ | Stock actual multiplicado por el precio de venta. |
| Acciones | Editar o eliminar el producto. |

### 3.3 Registrar Producto y Categorías

Desde "Registrar Producto" se da de alta un producto nuevo: nombre, categoría, precio de compra (si se deja vacío, se usa el mismo precio de venta), precio de venta, stock inicial, alerta de stock mínimo e imagen. El sistema calcula y muestra el margen de ganancia en tiempo real mientras se escriben los precios. "Editar Producto" usa el mismo tipo de validación: no permite guardar campos vacíos ni precios o stock negativos.

Las categorías se administran en su propia sección: se agregan por nombre (con un emoji opcional) y se pueden eliminar.

### 3.4 Compras (Reabastecimiento)

Registra una entrada de mercadería: se elige el producto, la cantidad comprada y el costo total. El sistema suma automáticamente esa cantidad al stock del producto y guarda el gasto en el historial de compras. La cantidad debe ser mayor a 0 y el costo no puede ser negativo.

### 3.5 Cambios Pendientes

Cuando una venta se registra como "Queda Pendiente", aparece aquí con el monto que se le debe al cliente. Al entregarlo, se marca como "Entregado" y sale de la lista de pendientes.

### 3.6 Reportes y Resúmenes

Reportes permite filtrar el historial de ventas (por fecha, cliente, etc.). Resúmenes muestra un panel con productos más vendidos, ingresos y el margen de ganancia (precio de venta menos precio de compra) por producto — otra razón por la que el precio de compra necesitaba ser visible también en el Inventario.

### 3.7 Responsables

Aquí se configuran los correos que reciben alertas automáticas: resumen mensual, resumen anual, aviso de venta grande (según un umbral en dólares) y alerta de stock bajo. Cada responsable se puede activar/desactivar y probar de forma manual con "Probar Envío de Correo", sin esperar a la fecha programada.

### 3.8 Papelera

Guarda una copia de lo que se elimina (productos, categorías, ventas anuladas) junto con la fecha y el motivo, como respaldo ante errores.

### 3.9 Ajustes

Incluye la activación del modo sin conexión (ver sección 4), la inicialización de la base de datos (crea las hojas necesarias sin borrar datos existentes) y las copias de seguridad manuales: exportar todo a un archivo JSON, importar un respaldo, o enviar una copia por correo a los Responsables de inmediato.

---

## 4. Modo sin conexión (offline)

Este modo permite seguir vendiendo aunque se corte el internet. Se activa manualmente desde Ajustes con la casilla "Modo Offline". Cuando está activo y el navegador detecta que no hay conexión, aparece un aviso rojo arriba del Punto de Venta.

<p align="center"><img src="manual_assets/offline_banner.png" width="750" alt="Aviso de sin conexión"></p>
<p align="center"><i>Figura 4. Aviso de "sin conexión" y aviso de ventas pendientes de sincronizar.</i></p>

**Qué pasa exactamente al vender sin internet:**

- La venta se guarda en una cola dentro del propio navegador (`localStorage`), no en Google Sheets todavía.
- El stock de los productos vendidos se descuenta en la copia local (caché) del inventario, para que el Punto de Venta siga mostrando cifras realistas mientras se sigue vendiendo sin conexión.
- Todas las validaciones de stock (semáforo de colores, bloqueo al agregar de más) funcionan igual sin internet, porque se comparan contra esa copia local — este es precisamente el cambio que evita vender más unidades de las que realmente hay.
- Al recuperar la conexión, aparece el aviso "tienes N ventas pendientes de subir a la nube" con el botón "Sincronizar Ahora".

<p align="center"><img src="manual_assets/flujo_venta.png" width="520" alt="Flujo de una venta online y offline"></p>
<p align="center"><i>Figura 5. Flujo completo de una venta, en línea y sin conexión.</i></p>

**Conflictos de stock al sincronizar (y cómo resolverlos):**

Puede ocurrir que, entre que una venta se guardó sin conexión y el momento de sincronizar, el stock real ya haya bajado (por ejemplo, otro dispositivo vendió el mismo producto). En ese caso el backend rechaza esa venta puntual con un mensaje de stock insuficiente.

> ⚠️ **Antes esto obligaba a borrar los datos del navegador.** Ahora el sistema distingue ese tipo de conflicto de un simple error de red, y ofrece el botón "Ver pendientes" para revisarlo y resolverlo sin perder el resto de la información guardada en el dispositivo.

<p align="center"><img src="manual_assets/ventas_pendientes.png" width="750" alt="Modal Ver pendientes"></p>
<p align="center"><i>Figura 6. Modal "Ver pendientes", mostrando una venta con conflicto de stock.</i></p>

Desde ese listado, cada venta pendiente puede descartarse manualmente (si ya no es válida) — al hacerlo, el stock que se le había restado a la caché local se devuelve automáticamente, y la venta se quita de la cola.

---

## 5. Estructura de datos (Google Sheets)

Cada módulo de la aplicación corresponde a una hoja dentro de la misma Google Sheet.

### Productos

| Columna | Descripción |
|---|---|
| `id` | Identificador único del producto. |
| `nombre` | Nombre del producto. |
| `código` | Código o SKU (opcional). |
| `categoría` | Categoría a la que pertenece. |
| `precio_compra` | Costo de compra por unidad. |
| `precio_venta` | Precio de venta al público. |
| `stock` | Unidades disponibles actualmente. |
| `stock_minimo` | Umbral para la alerta de stock bajo. |
| `imagen_url` | Enlace a la imagen del producto. |
| `favorito` | Marca de producto favorito (uso futuro). |
| `activo` | `false` si el producto fue eliminado (pasa a Papelera). |
| `fecha_creado` | Fecha de alta del producto. |

### Ventas (cabecera) y VentaDetalle (líneas)

| Columna (Ventas) | Descripción |
|---|---|
| `id` | Identificador único de la venta. |
| `fecha` | Fecha y hora de la venta. |
| `cliente_nombre` | Nombre del cliente (obligatorio). |
| `subtotal` / `total` | Monto de la venta. |
| `pago_con` / `cambio` | Con cuánto pagó el cliente y el vuelto. |
| `usuario` | Quién registró la venta. |
| `estado` | Ej. completada. |

| Columna (VentaDetalle) | Descripción |
|---|---|
| `venta_id` | A qué venta pertenece esta línea. |
| `producto_id` / `producto_nombre` | Producto vendido. |
| `cantidad` | Unidades vendidas de ese producto. |
| `precio_unitario` / `subtotal_linea` | Precio y subtotal de esa línea. |

### Otras hojas

| Hoja | Para qué sirve |
|---|---|
| `Categorias` | Nombre, emoji y estado de cada categoría. |
| `Compras` | Historial de reabastecimiento: producto, cantidad, costo total y fecha. |
| `Cambios` | Vueltos pendientes de entregar, ligados a una venta. |
| `Papelera` | Copia de productos, categorías o ventas eliminadas, con motivo y fecha. |
| `Actividad` | Bitácora de acciones importantes del sistema. |
| `Responsables` | Correos configurados para recibir alertas automáticas. |

---

## 6. Funciones del backend (codigo.gs)

`codigo.gs` es el programa de Google Apps Script que actúa como servidor: recibe las solicitudes del navegador, valida la información y lee/escribe en la Google Sheet.

**Enrutamiento de solicitudes**

| Función | Qué hace |
|---|---|
| `doGet(e)` | Atiende solicitudes de lectura (consultar inventario, ventas, categorías, etc.) según el parámetro `action`. |
| `doPost(e)` | Atiende solicitudes de escritura (registrar venta, crear/editar/eliminar producto, etc.) según el campo `action` del cuerpo JSON. |

**Ventas e inventario**

| Función | Qué hace |
|---|---|
| `registrarVenta(data)` | Valida el stock real de cada producto del carrito (sumando cantidades repetidas), y solo si alcanza para todos, registra la venta y descuenta el inventario. |
| `anularVenta(data)` | Elimina una venta, devuelve el stock al inventario, borra cambios pendientes ligados a ella y guarda una copia completa en la Papelera. |
| `registrarCompra(data)` | Suma stock comprado al producto y registra el gasto. Rechaza cantidades ≤ 0 o costos negativos. |
| `editarProducto(data)` | Actualiza nombre, precios, stock o categoría de un producto, evitando guardar valores negativos o inválidos. |
| `eliminarProducto` / `restaurarProducto` | Mueve un producto a la Papelera (o lo restaura desde ahí). |

**Cambios, responsables y respaldo**

| Función | Qué hace |
|---|---|
| `marcarCambioPagado(data)` | Marca un vuelto pendiente como entregado. |
| `agregarResponsable` / `editarResponsable` | Alta y edición de correos que reciben alertas. |
| `exportarDatosCompletos()` | Arma el archivo JSON de respaldo manual para descargar. |
| `importarDatosCompletos(data)` | Reemplaza los datos actuales por los de un archivo de respaldo. |
| `iniciarBD()` | Crea las hojas necesarias en la Google Sheet si no existen. |

**Alertas automáticas** (disparador diario `tareasDiarias()`)

| Función | Qué hace |
|---|---|
| `avisarVentaGrande_` | Envía un correo si una venta superó el umbral configurado. |
| `enviarAlertaStockBajo_` | Envía la lista de productos en stock bajo o agotados. |
| `enviarResumenMensual_` / `enviarResumenAnual_` | Envían el resumen de ventas del mes o del año a los Responsables que lo tengan activado. |
| `respaldoDiario_` | Genera y envía la copia de seguridad diaria (JSON y/o CSV). |
| `probarEnvioResponsable(data)` | Envía cualquiera de los reportes anteriores de inmediato, sin esperar la fecha programada (botón "Probar Envío"). |

---

## 7. Funciones clave del frontend (index.html)

**Catálogo y carrito**

| Función | Qué hace |
|---|---|
| `loadInventario()` | Descarga el catálogo actual; si falla por falta de red y el modo offline está activo, usa la copia guardada en el dispositivo. |
| `renderGridPOS()` | Dibuja las tarjetas de producto en el Punto de Venta, con el semáforo de stock. |
| `addToCart(id)` | Agrega un producto al carrito, validando que no se exceda el stock disponible. |
| `cambiarQty(idx, delta)` | Aumenta o reduce la cantidad de una línea del carrito; el aumento también se valida contra el stock. |
| `renderCarrito()` | Dibuja el carrito, mostrando el stock disponible de cada producto. |

**Registrar y sincronizar ventas**

| Función | Qué hace |
|---|---|
| `finalizarVenta(quedaPendiente)` | Valida cliente, calcula el vuelto, hace una validación final de stock y decide si envía la venta al servidor o la guarda localmente. |
| `guardarVentaLocal(payload)` | Guarda una venta en la cola offline y descuenta el stock en la caché local. |
| `sincronizarVentas()` | Reenvía la cola de ventas pendientes al servidor; distingue conflictos de stock de errores de red. |
| `verVentasPendientes()` | Muestra el listado de ventas guardadas localmente, con el motivo si alguna no pudo sincronizarse. |
| `descartarVentaPendiente(idx)` | Elimina una venta pendiente de la cola y devuelve el stock que se le había descontado en la caché. |

**Inventario, compras y otros módulos**

| Función | Qué hace |
|---|---|
| `renderInventario()` | Dibuja la tabla de Inventario, incluyendo la columna P.Compra. |
| `guardarProducto()` / `guardarEdicionProducto()` | Crean o editan un producto, validando campos obligatorios y que no haya precios o stock negativos. |
| `guardarCompra()` | Registra una compra de reabastecimiento, validando cantidad y costo. |
| `loadCambios()` / `marcarPagado()` | Listan y liquidan los vueltos pendientes. |
| `loadReportes()` / `loadResumenes()` | Cargan el historial de ventas y el panel de resúmenes. |
| `exportarDatos()` / `importarDatos()` | Descargan o restauran una copia de seguridad en JSON desde el navegador. |

---

## 8. Validaciones y correcciones aplicadas

| Problema encontrado | Corrección aplicada |
|---|---|
| Se podía vender más unidades de las que había en stock, especialmente sin conexión. | Validación de stock en 3 niveles: al agregar al carrito, al aumentar cantidad, y una verificación final antes de registrar la venta (online y offline). |
| Al sincronizar una venta rechazada por falta de stock, el sistema reintentaba para siempre, obligando a borrar los datos del navegador. | Se distingue un conflicto de stock de un error de red; nuevo panel "Ver pendientes" para revisar y descartar manualmente esas ventas. |
| El Inventario no mostraba el precio de compra junto al de venta. | Se agregó la columna P.Compra en la tabla de Inventario y Stock. |
| `editarProducto` podía guardar stock o precios como `NaN` o negativos si el formulario enviaba datos inválidos. | El backend ahora normaliza y limita esos valores a un mínimo de 0. |
| `registrarCompra` aceptaba cantidades o costos negativos. | Se rechazan cantidades ≤ 0 y costos negativos, tanto en el formulario como en el backend. |
| Si fallaba la red al guardar un producto, el sistema no avisaba nada (fallo silencioso). | Se agregó el mensaje de error correspondiente. |
| El formulario de Editar Producto no validaba campos vacíos ni valores negativos. | Ahora valida igual que el formulario de Registrar Producto. |
| Si fallaba la red al agregar/eliminar una categoría o eliminar un responsable, el ícono de carga quedaba girando para siempre. | Se agregó manejo de errores en las tres funciones, con aviso al usuario. |

---

## 9. Despliegue de esta actualización

**Backend (`codigo.gs`)**
1. Abre el proyecto de Google Apps Script vinculado a tu Google Sheet.
2. Reemplaza el contenido del archivo `.gs` actual por el de `codigo.gs`.
3. Guarda los cambios.
4. Ve a **Implementar → Administrar implementaciones**, edita la implementación web existente (ícono de lápiz) y en "Versión" elige **Nueva versión**.
5. Haz clic en **Implementar**.

> ⚠️ Si en vez de editar la implementación existente creas una "Nueva implementación", la URL del script cambiará y habrá que actualizar `SCRIPT_URL` en `index.html`.

**Frontend (`index.html`)**
- Si el HTML se sirve desde el propio proyecto de Apps Script, reemplaza el archivo HTML del proyecto por este `index.html` y vuelve a implementar como en el paso anterior.
- Si se aloja aparte (Drive, Sites, otro hosting), sube este archivo reemplazando el anterior.
- Verifica que la constante `SCRIPT_URL` siga apuntando a la URL correcta de tu Web App.

**Después de desplegar**
- Abre la consola del navegador (F12) y confirma que no aparezcan errores al cargar.
- Vende un producto con poco stock hasta agotarlo, y confirma que el sistema bloquea la venta al llegar a 0.
- Activa el Modo Offline en Ajustes, desconecta el wifi y repite la prueba anterior para confirmar que también se bloquea sin conexión.
- Si ya tenías ventas pendientes de sincronizar de antes de esta actualización, revísalas en "Ver pendientes" por si alguna necesita descartarse manualmente.
