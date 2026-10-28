# Alta de productos

Programa para crear los productos de la tienda, imprimir sus etiquetas con BarTender y generar el Excel para importarlos en Siigo Nube.

Funciona sin internet en Windows 10 u 11. Trae su propio Node.js, así que no hay que instalar nada más.

## Qué hace

- **Alta de productos**: referencia, descripción, código de barras, grupo, costo, precio detal, precio mayorista, impuesto (IVA 19 %, 5 %, exento, excluido, no gravado) e INC.
- **Precio sugerido a partir del costo**: costo × (1 + ganancia) × (1 + IVA), redondeado a terminación 900. Ejemplo: 76.000 → 117.900.
- **Códigos de barras propios de 6 dígitos** (007760, 007761…). Continúan la numeración de los productos que ya tiene la tienda y nunca se repiten, aunque se borre el producto.
- **Etiquetas con BarTender**: imprime los productos de la sesión, con cantidad y precio editables, historial y reimpresión.
- **Excel para Siigo**: llena la plantilla oficial de importación sin dañar su formato. Revisa los campos obligatorios y las listas, y arma lotes de 500 productos.
- **Guardado automático** en una base SQLite, con copias de seguridad con fecha.

## Instalación

1. Descarga el `.zip` y extráelo en una carpeta fija, por ejemplo `C:\AltaProductos`. No lo uses desde dentro del `.zip`.
2. Doble clic en `instalar.bat`. Crea el acceso directo **Alta de productos** en el escritorio y en el menú Inicio.
3. Abre el acceso directo. El programa se abre en el navegador, en `http://127.0.0.1:3789`.

Si Windows muestra «Windows protegió su PC», elige **Más información → Ejecutar de todas formas**.

## Uso diario

1. Crea el producto y guárdalo. Si dejas el código de barras vacío, se genera solo.
2. Botón verde **Imprimir**: marca los productos, pon la cantidad y dale **Imprimir**.
3. Al final del día ve a **Excel para Siigo → Exportar a Excel**. Por defecto exporta solo los productos que aún no se han subido a Siigo. Luego sube el archivo en Siigo: Configuración → Productos y servicios → Importación.

## Etiquetas con BarTender

El programa genera un CSV, separado por punto y coma, con una fila por etiqueta. Luego ejecuta:

```
bartend.exe /F="plantilla.btw" /D="datos.csv" /PRN="impresora" /C=1 /P /X
```

Las rutas de `bartend.exe` y de la plantilla, la impresora y las copias se configuran en **Configuración de BarTender** o en la pestaña Imprimir. La impresora se elige de la lista de impresoras instaladas en el computador, así que sirve cualquiera que haya en la tienda. Si se agrega una impresora nueva, solo hay que instalarla en Windows y elegirla ahí.

Para preparar la plantilla `.btw`:

1. En BarTender, ve a Archivo → Configuración de base de datos y agrega un **archivo de texto** que apunte a `datos\etiquetas\etiquetas.csv`. Debe estar delimitado por **punto y coma**, con la primera fila como nombres de campo.
2. Enlaza cada objeto de la etiqueta a su campo:

| Objeto de la etiqueta | Campo | Ejemplo |
|---|---|---|
| Descripción + referencia | `referencia_descripcion` | Jean VFk11017 |
| Código de barras (Code 128) | `codigo_barras` | 007754 |
| Código interno | `codigo_interno` | 007754 |
| Precio al mayor | `precio_mayor_codigo` | 2\*1079VY |
| Precio detal (el `$` va fijo en la plantilla) | `precio_detal_txt` | 140900 |

3. Para rollos de 3 etiquetas a lo ancho, configúralo en Archivo → Configurar página → Diseño: 1 fila, 3 columnas y la separación del rollo.

Para que BarTender imprima desde el programa, debe estar **cerrado**. Si queda abierto, o con procesos `bartend.exe` ocultos, el programa dice que imprimió pero no sale nada.

## Códigos de barras

- Los códigos generados son consecutivos de 6 dígitos con ceros adelante.
- El siguiente código es el mayor entre el último producto con código de 6 dígitos y el contador guardado. El contador se guarda en la base y en `configuracion\opciones.json`.
- En **Código de barras propio → Seguir después del número** se fija desde dónde continuar. También sirve para corregirlo; si el número es menor al actual, pide confirmación.
- También se aceptan códigos de fabricante (EAN/UPC de 8, 12, 13 o 14 dígitos) escritos a mano.

## Carpetas

```
AltaDeProductos/
├── app/
│   ├── server.js          servidor local (127.0.0.1), BarTender, archivos
│   ├── siigo.js           escribe el Excel sobre la plantilla de Siigo
│   └── public/            interfaz (index.html, imprimir.html, sql.js)
├── node/node.exe          Node.js incluido
├── datos/                 base, copias, CSV de etiquetas y logs   (no se sube al repo)
├── configuracion/         opciones.json                           (no se sube al repo)
├── plantillas/            plantilla de Siigo (.xlsx) y etiqueta (.btw)
├── exportados/            Excel generados para Siigo              (no se sube al repo)
├── iniciar.vbs / iniciar.bat
├── instalar.bat / instalar.ps1
└── LEEME.txt
```

## Datos y copias de seguridad

- La base es `datos\productos.sqlite`. Se puede mover a OneDrive desde el panel **Base de datos y copias de seguridad**.
- Se hacen copias automáticas al abrir el programa y cada hora de trabajo, en `datos\copias`. Se guardan las últimas 60.
- Para llevar los datos a otro computador, copia `productos.sqlite` y ábrelo con **Abrir otra base…**. Para llevar la configuración, copia la carpeta `configuracion`.
- No uses la misma base en dos computadores al mismo tiempo.

## Problemas comunes

| Problema | Solución |
|---|---|
| El programa no abre | Ciérralo y vuelve a abrirlo. Los errores quedan en `datos\registro.log`. |
| «La base se modificó en otra ventana» | Recarga la página con F5. Pasa cuando el programa está abierto en dos pestañas. |
| Dice «IMPRESO» pero no sale nada | Cierra BarTender y finaliza los `bartend.exe` en el Administrador de tareas. |
| `EPERM` al imprimir | Cierra `etiquetas.csv` si está abierto en Excel. |
| No sale el código de barras | Revisa que el objeto esté enlazado a `codigo_barras` con simbología Code 128. |
| Historial de impresiones | `datos\impresiones.log` (no se borra al cerrar el programa). |

## Mantenimiento

- El código vive en este repositorio privado de GitHub. Una sola persona lo mantiene.
- Para actualizar un computador de la tienda, reemplaza las carpetas `app/` y `plantillas/`, y los archivos `.bat`, `.vbs`, `.ps1` y `LEEME.txt`. No toques `datos/` ni `configuracion/`, porque ahí están la base y los ajustes de ese computador.
- Después de actualizar, cierra el programa con «Cerrar el programa» y ábrelo de nuevo.

## Desarrollo

- Interfaz en el navegador con [sql.js](https://github.com/sql-js/sql.js) (SQLite en WebAssembly). Servidor en Node.js sin frameworks; la única dependencia es `jszip`.
- Para arrancar en desarrollo: `node\node.exe app\server.js` y abrir `http://127.0.0.1:3789`.
- El servidor solo escucha en `127.0.0.1` y usa un token por sesión.

`.gitignore` recomendado:

```
datos/
configuracion/
exportados/
node/
plantillas/*.xlsx
```

---

Esta herramienta ayuda a evitar errores de captura, pero no reemplaza la asesoría contable. Confirma el tipo de impuesto de cada producto con tu contador.
