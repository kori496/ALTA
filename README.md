ALTA DE PRODUCTOS - instalación y uso
=====================================

Programa para crear los productos de la tienda, guardarlos en una base .sqlite,
imprimir etiquetas con BarTender y sacar el Excel para importar en Siigo Nube.
Funciona sin internet. Windows 10 u 11.

INSTALAR EN UN COMPUTADOR
1. Copia el archivo .zip al computador y descomprímelo (clic derecho > Extraer todo)
   en una carpeta fija, por ejemplo C:\AltaProductos. No lo uses desde dentro del .zip.
2. Entra a la carpeta y haz doble clic en  instalar.bat
   Crea el acceso directo "Alta de productos" en el escritorio y en el menú Inicio.
3. Abre "Alta de productos" desde el escritorio. Se abre en el navegador.

Si Windows muestra "Windows protegió su PC": clic en "Más información" > "Ejecutar de todas formas".

USO DIARIO
- Cada cambio se guarda solo en el disco (arriba dice "Guardado" y la hora).
- Los datos quedan en la carpeta  datos\productos.sqlite
- Copias de seguridad con fecha en  datos\copias  (al abrir el programa y cada hora de trabajo; se guardan las últimas 60).
- Para usar una base que ya tenías (.sqlite de la versión anterior): botón "Abrir otra base…".
  Antes de reemplazarla, el programa guarda una copia de la base actual.
- Para guardar la base en OneDrive: panel "Base de datos y copias de seguridad" > escribe la ruta > "Cambiar ubicación".
  Si dos computadores usan la misma base de OneDrive, no la editen al mismo tiempo.
- Para cerrar el programa del todo: panel "Base de datos y copias de seguridad" > "Cerrar el programa".
  Si cierras todas las pestañas, el programa se cierra solo a los 10 minutos (y se borra el
  historial de impresión). Si vuelves a abrirlo antes, sigue igual.

EXCEL PARA SIIGO
1. En Siigo: Configuración > Productos y servicios > Importación > Productos y Servicios > Crear plantilla.
2. En el programa: "Cargar plantilla de Siigo (.xlsx)…" (o copia la plantilla a la carpeta  plantillas).
3. Revisa "Columnas de la plantilla": qué dato de la app va en cada columna. Las obligatorias
   (encabezado azul oscuro) deben tener un dato o un valor fijo. Guarda.
4. "Revisar productos" muestra cuáles fallan y por qué. "Exportar a Excel" crea los archivos en  exportados
   (máximo 500 productos por archivo).
5. Sube el archivo en Siigo, en la misma pantalla de importación.

ETIQUETAS CON BARTENDER
1. Panel "Configuración de BarTender": ruta de bartend.exe, plantilla .btw, impresora y copias.
2. En la plantilla .btw, enlaza la base de datos al archivo  datos\etiquetas\etiquetas.csv
   (columnas: referencia, nombre, codigo_barras, precio_detal, grupo, impuesto).
   Cada fila del CSV es una etiqueta: si pides 3, el producto sale 3 veces.
3. Botón verde "Imprimir" (arriba): abre otra pestaña con SOLO los productos guardados desde que
   abriste el programa. Marca cuáles imprimir, cuántas etiquetas de cada uno y, si quieres,
   otro precio para la etiqueta (no cambia el producto guardado). Luego "Imprimir".
   "← Volver" regresa a la ventana principal.
4. Historial (en esa pestaña): cada impresión con hora, referencia, precio y cantidad.
   "Reimprimir" deja cambiar la cantidad y el precio. Se borra al cerrar el programa.
   Ctrl+F7 (en cualquiera de las dos ventanas) reinicia la sesión y empieza de cero:
   borra la lista de la sesión y el historial; los productos guardados y el log no se tocan.
5. Log que no se borra:  datos\impresiones.log  (botón "Ver log de impresiones").
   Sirve si se cierra algo por accidente: ahí está todo lo que se imprimió y los errores.
También están el botón "Etiqueta" de cada producto y "Imprimir con BarTender" del formulario.

OPCIONES GUARDADAS (carpeta configuracion)
- configuracion\opciones.json guarda: ruta de bartend.exe, plantilla .btw, impresora elegida, copias,
  parámetros, impresoras ya usadas, ubicación de la base, columnas de Siigo y la ÚLTIMA CARPETA usada
  para abrir cada cosa (base, plantilla de Siigo, .btw, bartend.exe, códigos).
- Los botones "Abrir otra base…", "Cargar plantilla de Siigo", "Cargar códigos existentes" y "Buscar…"
  abren la ventana de Windows directamente en esa última carpeta.
- La impresora se elige de una lista: impresoras instaladas en Windows y las que ya usaste.
- Para llevar las opciones a otro computador, copia la carpeta configuracion.

LLEVAR LOS DATOS A OTRO COMPUTADOR
Copia el archivo  datos\productos.sqlite  y ábrelo allá con "Abrir otra base…".

PROBLEMAS
- No abre: cierra el programa, vuelve a abrirlo. Errores en  datos\registro.log
- "La base se modificó en otra ventana": recarga la página (F5). Pasa si se abrió el programa en dos pestañas.
