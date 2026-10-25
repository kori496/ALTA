// Alta de productos: servidor local (solo escucha en 127.0.0.1).
// Sirve la interfaz, guarda la base .sqlite en disco con copias de seguridad,
// escribe el Excel de Siigo sobre la plantilla y manda las etiquetas a BarTender.
const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const os = require("os");
const { spawn } = require("child_process");
const siigo = require("./siigo");

const RAIZ = process.env.ALTA_RAIZ ? path.resolve(process.env.ALTA_RAIZ) : path.resolve(__dirname, "..");
const DATOS = path.join(RAIZ, "datos");
const COPIAS = path.join(DATOS, "copias");
const ETIQUETAS = path.join(DATOS, "etiquetas");
const EXPORTADOS = path.join(RAIZ, "exportados");
const PLANTILLAS = path.join(RAIZ, "plantillas");
const PUBLICO = path.join(__dirname, "public");
// Opciones guardadas (impresora, BarTender, últimas carpetas…): carpeta propia para poder respaldarla o copiarla a otro PC
const CONFIG_DIR = path.join(RAIZ, "configuracion");
const CONFIG = path.join(CONFIG_DIR, "opciones.json");
const CONFIG_VIEJO = path.join(DATOS, "config.json"); // donde se guardaba antes del 24 oct
const MAX_COPIAS = 60;
const TOKEN = crypto.randomBytes(24).toString("hex");
// Firma del código con el que arrancó este proceso: si los archivos cambian (actualización),
// el siguiente doble clic cierra este proceso y arranca el código nuevo
function firmaCodigo() {
  return ["server.js", "siigo.js"].map(f => { try { return fs.statSync(path.join(__dirname, f)).mtimeMs; } catch (e) { return 0; } }).join("-");
}
const FIRMA = firmaCodigo();
const ID_APP = "alta-de-productos";

for (const d of [DATOS, COPIAS, ETIQUETAS, EXPORTADOS, PLANTILLAS, CONFIG_DIR]) fs.mkdirSync(d, { recursive: true });

// ---------- configuración ----------
const BT_ARGS = '/F="{plantilla}" /D="{datos}" {impresora} /C={copias} /P /X';
const CONFIG_DEFECTO = {
  puerto: 3789,
  rutaBase: "",
  bartender: {
    exe: "C:\\Program Files\\Seagull\\BarTender Suite\\bartend.exe",
    plantilla: "",
    impresora: "",
    copias: 1,
    argumentos: BT_ARGS
  },
  siigo: { plantilla: "", hoja: "", columnas: {} },
  barras: { ultimoCorto: 0, ultimoCodigo: "" },   // último código de barras corto usado (respaldo fuera de la base)
  impresorasUsadas: [],   // las últimas impresoras elegidas (la primera es la más reciente)
  ultimasCarpetas: {}     // por tipo de archivo: base, plantillaSiigo, btw, bartender, codigos
};
function leerConfig() {
  let c = {};
  // Primera vez con la carpeta configuracion: se traen las opciones de datos\config.json
  const origen = fs.existsSync(CONFIG) ? CONFIG : (fs.existsSync(CONFIG_VIEJO) ? CONFIG_VIEJO : null);
  try { if (origen) c = JSON.parse(fs.readFileSync(origen, "utf8")); } catch (e) { /* archivo dañado: valores por defecto */ }
  return {
    ...CONFIG_DEFECTO, ...c,
    bartender: { ...CONFIG_DEFECTO.bartender, ...(c.bartender || {}) },
    siigo: { ...CONFIG_DEFECTO.siigo, ...(c.siigo || {}) },
    barras: { ...CONFIG_DEFECTO.barras, ...(c.barras || {}) },
    impresorasUsadas: Array.isArray(c.impresorasUsadas) ? c.impresorasUsadas : [],
    ultimasCarpetas: { ...(c.ultimasCarpetas || {}) }
  };
}
function guardarConfig(c) { escribirSeguro(CONFIG, Buffer.from(JSON.stringify(c, null, 2), "utf8")); }
let config = leerConfig();
if (!fs.existsSync(CONFIG)) {
  guardarConfig(config);
  if (fs.existsSync(CONFIG_VIEJO)) {
    try { fs.renameSync(CONFIG_VIEJO, CONFIG_VIEJO + ".migrado"); } catch (e) { /* se queda donde está; ya no se usa */ }
  }
  const leeme = path.join(CONFIG_DIR, "LEEME.txt");
  if (!fs.existsSync(leeme)) {
    fs.writeFileSync(leeme, "\uFEFFOpciones guardadas de Alta de productos (opciones.json):\r\n" +
      "- BarTender: ruta de bartend.exe, plantilla .btw, impresora, copias y parámetros\r\n" +
      "- impresoras ya usadas y última carpeta de cada tipo de archivo (base, plantilla de Siigo, .btw, bartend.exe, códigos)\r\n" +
      "- ubicación de la base y columnas de la plantilla de Siigo\r\n\r\n" +
      "Para llevar las opciones a otro computador, copia esta carpeta a la carpeta del programa de allá.\r\n");
  }
}
function recordarImpresora(nombre) {
  const n = String(nombre || "").trim();
  if (!n) return;
  config.impresorasUsadas = [n, ...config.impresorasUsadas.filter(x => x !== n)].slice(0, 10);
}

// ---------- ventana de Windows para elegir archivos (empieza en la última carpeta usada) ----------
const DIALOGOS = {
  base: { titulo: "Abrir base de productos", filtro: "Base de productos (*.sqlite;*.db)|*.sqlite;*.db|Todos los archivos (*.*)|*.*" },
  plantillaSiigo: { titulo: "Plantilla de Siigo (.xlsx)", filtro: "Excel (*.xlsx)|*.xlsx" },
  btw: { titulo: "Plantilla de etiqueta de BarTender (.btw)", filtro: "Plantilla BarTender (*.btw)|*.btw|Todos los archivos (*.*)|*.*" },
  bartender: { titulo: "Programa de BarTender (bartend.exe)", filtro: "BarTender (bartend.exe)|bartend.exe|Programas (*.exe)|*.exe" },
  codigos: { titulo: "Códigos de barras que ya existen (CSV o TXT)", filtro: "CSV o texto (*.csv;*.txt)|*.csv;*.txt|Todos los archivos (*.*)|*.*" }
};
const elegidos = {}; // tipo -> ruta elegida en la última ventana
function carpetaInicial(tipo) {
  const existe = d => { try { return !!d && fs.statSync(d).isDirectory(); } catch (e) { return false; } };
  const candidatas = [config.ultimasCarpetas[tipo]];
  if (tipo === "base") candidatas.push(path.dirname(rutaBase()));
  if (tipo === "btw" && config.bartender.plantilla) candidatas.push(path.dirname(config.bartender.plantilla));
  if (tipo === "bartender" && config.bartender.exe) candidatas.push(path.dirname(config.bartender.exe), "C:\\Program Files");
  if (tipo === "plantillaSiigo") candidatas.push(path.join(os.homedir(), "Downloads"));
  candidatas.push(path.join(os.homedir(), "Documents"));
  return candidatas.find(existe) || "";
}
// La ventana se abre encima del navegador (formulario invisible "siempre visible" como dueño)
const PS_DIALOGO = [
  "$ErrorActionPreference = 'Stop'",
  "[Console]::OutputEncoding = [Text.Encoding]::UTF8",
  "Add-Type -AssemblyName System.Windows.Forms, System.Drawing",
  "$d = New-Object System.Windows.Forms.OpenFileDialog",
  "$d.Title = $env:ALTA_TITULO",
  "$d.Filter = $env:ALTA_FILTRO",
  "if ($env:ALTA_CARPETA -and (Test-Path -LiteralPath $env:ALTA_CARPETA)) { $d.InitialDirectory = $env:ALTA_CARPETA }",
  "if ($env:ALTA_SOLO_PROBAR) { [Console]::Out.Write('PRUEBA ' + $d.InitialDirectory); exit 0 }",
  "$o = New-Object System.Windows.Forms.Form",
  "$o.TopMost = $true; $o.ShowInTaskbar = $false; $o.Opacity = 0",
  // La ventana de archivos se abre en la esquina del formulario dueño: se ubica para que quede centrada
  "$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea",
  "$o.StartPosition = 'Manual'; $o.Size = New-Object System.Drawing.Size(1, 1)",
  "$o.Location = New-Object System.Drawing.Point([int][Math]::Max($wa.Left, $wa.Left + $wa.Width / 2 - 450), [int][Math]::Max($wa.Top, $wa.Top + $wa.Height / 2 - 300))",
  "$o.Show(); $o.Activate()",
  "if ($d.ShowDialog($o) -eq [System.Windows.Forms.DialogResult]::OK) { [Console]::Out.Write($d.FileName) }",
  "$o.Close()"
].join("\n");
let dialogoAbierto = false;
function ventanaArchivo(tipo) {
  if (process.env.ALTA_DIALOGO_FALSO) { // para las pruebas automáticas
    const r = JSON.parse(process.env.ALTA_DIALOGO_FALSO)[tipo];
    return Promise.resolve(r || null);
  }
  if (dialogoAbierto) return Promise.reject(new Error("Ya hay una ventana para elegir archivo abierta (puede estar detrás del navegador)."));
  dialogoAbierto = true;
  return new Promise((resolve, reject) => {
    const p = spawn("powershell.exe", ["-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-EncodedCommand", Buffer.from(PS_DIALOGO, "utf16le").toString("base64")], {
      windowsHide: true,
      env: { ...process.env, ALTA_TITULO: DIALOGOS[tipo].titulo, ALTA_FILTRO: DIALOGOS[tipo].filtro, ALTA_CARPETA: carpetaInicial(tipo) }
    });
    let out = "", err = "";
    p.stdout.on("data", d => { out += d; });
    p.stderr.on("data", d => { err += d; });
    p.on("error", e => { dialogoAbierto = false; reject(e); });
    p.on("exit", codigo => {
      dialogoAbierto = false;
      if (codigo !== 0) return reject(new Error("No se pudo abrir la ventana de Windows: " + (err.trim().split("\n")[0] || "código " + codigo)));
      resolve(out.trim() || null);
    });
  });
}
async function elegirArchivo(tipo) {
  if (!DIALOGOS[tipo]) throw new Error("Tipo de archivo desconocido.");
  const ruta = await ventanaArchivo(tipo);
  if (!ruta) return null;
  elegidos[tipo] = ruta;
  config.ultimasCarpetas[tipo] = path.dirname(ruta);
  guardarConfig(config);
  return ruta;
}

// Impresoras instaladas en Windows (para la lista desplegable)
let cacheImpresoras = { hora: 0, lista: [] };
function impresorasInstaladas() {
  if (Date.now() - cacheImpresoras.hora < 60000) return Promise.resolve(cacheImpresoras.lista);
  return new Promise(resolve => {
    const ps = "[Console]::OutputEncoding = [Text.Encoding]::UTF8; Get-CimInstance Win32_Printer | ForEach-Object { $_.Name }";
    const p = spawn("powershell.exe", ["-NoProfile", "-EncodedCommand", Buffer.from(ps, "utf16le").toString("base64")], { windowsHide: true });
    let out = "";
    p.stdout.on("data", d => { out += d; });
    p.on("error", () => resolve([]));
    p.on("exit", () => {
      cacheImpresoras = { hora: Date.now(), lista: out.split(/\r?\n/).map(x => x.trim()).filter(Boolean) };
      resolve(cacheImpresoras.lista);
    });
  });
}

function rutaBase() { return config.rutaBase ? path.resolve(RAIZ, config.rutaBase) : path.join(DATOS, "productos.sqlite"); }

// ---------- utilidades ----------
function registrar(msg) {
  try { fs.appendFileSync(path.join(DATOS, "registro.log"), `${new Date().toISOString()} ${msg}\n`); } catch (e) { /* sin log */ }
}
function sello() {
  const d = new Date(), p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}
function esperar(ms) { Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms); }

// Escribe en un temporal y lo renombra; reintenta si OneDrive o el antivirus bloquean el archivo.
// Después verifica que en disco quedó exactamente lo escrito.
function escribirSeguro(destino, buf) {
  fs.mkdirSync(path.dirname(destino), { recursive: true });
  const tmp = `${destino}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, buf);
  for (let i = 0; ; i++) {
    try { fs.renameSync(tmp, destino); break; } catch (e) {
      if (i >= 8) { try { fs.unlinkSync(tmp); } catch (e2) { /* nada */ } throw e; }
      esperar(150 * (i + 1));
    }
  }
  const leido = fs.readFileSync(destino);
  if (!leido.equals(buf)) throw new Error(`El archivo ${destino} no quedó igual a lo guardado. Intenta de nuevo.`);
}

// ---------- base y copias ----------
let ultimaCopia = 0;

// Versión de la base = fecha y tamaño del archivo. La pestaña envía la versión que leyó;
// si el archivo cambió desde entonces (otra ventana u otro computador guardó), no se pisa.
// Como sale del archivo, sigue valiendo aunque el programa se reinicie.
function versionArchivo(ruta) {
  try { const st = fs.statSync(ruta); return `${st.mtimeMs}-${st.size}`; } catch (e) { return "0"; }
}

function hacerCopia(motivo) {
  const ruta = rutaBase();
  if (!fs.existsSync(ruta)) return null;
  const destino = path.join(COPIAS, `productos_${sello()}${motivo ? "_" + motivo : ""}.sqlite`);
  fs.copyFileSync(ruta, destino);
  ultimaCopia = Date.now();
  const viejas = fs.readdirSync(COPIAS).filter(f => f.endsWith(".sqlite")).sort().reverse().slice(MAX_COPIAS);
  for (const f of viejas) { try { fs.unlinkSync(path.join(COPIAS, f)); } catch (e) { /* en uso */ } }
  return destino;
}

function esSqlite(buf) { return buf.length >= 100 && buf.subarray(0, 16).toString("latin1") === "SQLite format 3\u0000"; }

function guardarBase(buf, versionCliente, reemplazo) {
  const ruta = rutaBase();
  if (!esSqlite(buf)) return { estado: 400, error: "Lo recibido no es una base SQLite." };
  if (String(versionCliente) !== versionArchivo(ruta)) {
    return { estado: 409, error: "La base se guardó desde otra ventana u otro computador. Recarga la página (F5) para ver los últimos datos." };
  }
  if (reemplazo) hacerCopia("antes-de-abrir");
  else if (Date.now() - ultimaCopia > 60 * 60 * 1000) hacerCopia("");
  escribirSeguro(ruta, buf);
  return { estado: 200, version: versionArchivo(ruta) };
}

// ---------- BarTender ----------
function csvEsc(v) {
  if (v === null || v === undefined) return "";
  const s = String(v);
  return /[";\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
let imprimiendo = Promise.resolve();
function imprimir(filas, copias) {
  const bt = config.bartender;
  if (!bt.exe || !fs.existsSync(bt.exe)) throw new Error(`No se encuentra BarTender en: ${bt.exe || "(sin ruta)"}. Revisa «Configuración de BarTender».`);
  if (!bt.plantilla || !fs.existsSync(bt.plantilla)) throw new Error(`No se encuentra la plantilla .btw: ${bt.plantilla || "(sin ruta)"}.`);
  // Siempre se imprime en la impresora elegida en el programa, nunca en la que trae guardada la plantilla
  if (!String(bt.impresora || "").trim()) throw new Error("No hay impresora elegida. Elígela y guárdala en la pestaña Imprimir (Configuración de impresión) o en «Configuración de BarTender».");
  // Columnas para la etiqueta V.YA (las 6 primeras se mantienen por compatibilidad)
  const cab = ["referencia", "nombre", "codigo_barras", "precio_detal", "grupo", "impuesto",
               "referencia_descripcion", "codigo_interno", "precio_detal_txt", "precio_mayor", "precio_mayor_sin_ceros", "precio_mayor_codigo"];
  const miles = n => (n === null || n === undefined || n === "" || !Number.isFinite(Number(n))) ? "" : Math.round(Number(n)).toLocaleString("es-CO");
  const sinPuntos = n => (n === null || n === undefined || n === "" || !Number.isFinite(Number(n))) ? "" : String(Math.round(Number(n)));
  // Precio al mayor para la etiqueta: se quitan las dos últimas cifras (centenas): 65000 -> 650 ; 107900 -> 1079
  const sinCeros = n => (n === null || n === undefined || n === "" || !Number.isFinite(Number(n)) || Number(n) <= 0) ? "" : String(Math.floor(Math.round(Number(n)) / 100));
  const filasEtiqueta = filas.map(f => ({
    ...f,
    // "Jean VFk11017". Productos traídos de CRS: la descripción ya trae la referencia y el código interno
    // es el mismo código de barras, así que solo va la descripción.
    referencia_descripcion: (!f.referencia || f.referencia === f.codigo_barras ||
                             String(f.nombre || "").toLowerCase().includes(String(f.referencia).toLowerCase()))
      ? (f.nombre || f.referencia || "") : `${f.nombre} ${f.referencia}`.trim(),
    codigo_interno: f.codigo_barras,               // en V.YA el código interno es el mismo código de barras
    precio_detal_txt: sinPuntos(f.precio_detal),   // 140900 (sin punto de miles, como la etiqueta impresa)
    precio_mayor: miles(f.precio_mayorista),
    precio_mayor_sin_ceros: sinCeros(f.precio_mayorista), // 65000 -> 650 ; 107900 -> 1079
    precio_mayor_codigo: sinCeros(f.precio_mayorista) ? `2*${sinCeros(f.precio_mayorista)}VY` : ""  // "2*1079VY"
  }));
  // Separador punto y coma: así Excel en español lo abre en columnas (en BarTender: delimitado por punto y coma)
  const csv = [cab.join(";"), ...filasEtiqueta.map(f => cab.map(k => csvEsc(f[k])).join(";"))].join("\r\n");
  const contenido = Buffer.from("\uFEFF" + csv, "utf8");
  // Cada impresión usa su propio archivo (se pasa a BarTender con /D): así no importa si etiquetas.csv
  // está abierto en Excel o lo tiene bloqueado otro BarTender.
  const TRABAJOS = path.join(ETIQUETAS, "trabajos");
  fs.mkdirSync(TRABAJOS, { recursive: true });
  const datos = path.join(TRABAJOS, `etiquetas_${sello()}_${Math.random().toString(36).slice(2, 6)}.csv`);
  fs.writeFileSync(datos, contenido);
  // Se guardan solo los últimos 30 trabajos (los que estén bloqueados se dejan para la próxima)
  try {
    fs.readdirSync(TRABAJOS).filter(n => /^etiquetas_.*\.csv$/i.test(n)).sort().slice(0, -30)
      .forEach(n => { try { fs.unlinkSync(path.join(TRABAJOS, n)); } catch (e) { /* bloqueado */ } });
  } catch (e) { /* nada */ }
  // etiquetas.csv (ruta fija con la que está enlazada la plantilla) se actualiza si se puede, sin frenar la impresión
  try { fs.writeFileSync(path.join(ETIQUETAS, "etiquetas.csv"), contenido); } catch (e) { registrar(`etiquetas.csv bloqueado (se imprime igual con ${path.basename(datos)}): ${e.code || e.message}`); }
  const args = (bt.argumentos || BT_ARGS)
    .replace("{plantilla}", bt.plantilla.replace(/"/g, ""))
    .replace("{datos}", datos)
    .replace("{impresora}", bt.impresora ? `/PRN="${bt.impresora.replace(/"/g, "")}"` : "")
    .replace("{copias}", String(Math.max(1, parseInt(copias ?? bt.copias, 10) || 1)))
    .replace(/\s+/g, " ").trim();
  const tarea = imprimiendo.then(() => new Promise(resolve => {
    registrar(`BarTender: "${bt.exe}" ${args}`);
    // Línea de comandos tal cual; la ruta del .exe va entre comillas porque suele tener espacios (Program Files)
    const p = spawn(bt.exe, [args], { windowsVerbatimArguments: true, windowsHide: true, argv0: `"${bt.exe}"` });
    const reloj = setTimeout(() => resolve({ ok: true, codigo: null, nota: "BarTender sigue abierto; revisa la impresora." }), 90000);
    p.on("error", e => { clearTimeout(reloj); resolve({ ok: false, error: e.message }); });
    const t0 = Date.now();
    p.on("exit", codigo => { clearTimeout(reloj); registrar(`BarTender terminó: código ${codigo}, ${Date.now() - t0} ms`); resolve({ ok: codigo === 0, codigo, error: codigo ? `BarTender terminó con código ${codigo}.` : null }); });
  }));
  imprimiendo = tarea.catch(() => {});
  return tarea.then(r => ({ ...r, comando: `"${bt.exe}" ${args}`, datos }));
}

// ---------- sesión de impresión ----------
// Vive en memoria: se reinicia solo cuando se cierra el programa. El log en disco no se borra nunca.
const LOG_IMPRESIONES = path.join(DATOS, "impresiones.log");
const MAX_ETIQUETAS = 1000; // por impresión
const sesion = { inicio: new Date().toISOString(), productos: [], historial: [] };
let siguienteId = 1;

function logImpresion(linea) {
  const d = new Date(), p = n => String(n).padStart(2, "0");
  const fecha = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  try { fs.appendFileSync(LOG_IMPRESIONES, `${fecha} | ${linea}\r\n`); } catch (e) { registrar(`No se pudo escribir el log de impresiones: ${e.message}`); }
}
function precioTxt(n) { return n === null || n === undefined || n === "" ? "sin precio" : "$" + Number(n).toLocaleString("es-CO"); }
function limpiarItem(x) {
  const t = v => (v === null || v === undefined ? "" : String(v).trim());
  const precio = x.precio_detal === "" || x.precio_detal === null || x.precio_detal === undefined ? null : Number(x.precio_detal);
  const mayor = x.precio_mayorista === "" || x.precio_mayorista === null || x.precio_mayorista === undefined ? null : Number(x.precio_mayorista);
  return { referencia: t(x.referencia), nombre: t(x.nombre), codigo_barras: t(x.codigo_barras),
           precio_detal: Number.isFinite(precio) ? precio : null, grupo: t(x.grupo), impuesto: t(x.impuesto),
           precio_mayorista: Number.isFinite(mayor) ? mayor : null,
           cantidad: parseInt(x.cantidad, 10) };
}

// items: [{ referencia, nombre, codigo_barras, precio_detal, grupo, impuesto, cantidad }]
async function imprimirItems(items, origen) {
  const lista = (Array.isArray(items) ? items : []).map(limpiarItem);
  if (!lista.length) throw new Error("No hay etiquetas para imprimir.");
  for (const it of lista) {
    if (!it.referencia || !it.codigo_barras) throw new Error(`La referencia ${it.referencia || "(sin referencia)"} no tiene código de barras.`);
    if (!Number.isInteger(it.cantidad) || it.cantidad < 1) throw new Error(`Cantidad no válida para ${it.referencia}.`);
    if (it.precio_detal !== null && it.precio_detal < 0) throw new Error(`Precio no válido para ${it.referencia}.`);
  }
  const total = lista.reduce((s, it) => s + it.cantidad, 0);
  if (total > MAX_ETIQUETAS) throw new Error(`Son ${total} etiquetas; el máximo por impresión es ${MAX_ETIQUETAS}.`);

  // Cada etiqueta es una fila del CSV: así cada producto sale con su propia cantidad en un solo trabajo
  const filas = lista.flatMap(it => Array.from({ length: it.cantidad }, () => it));
  let r;
  try { r = await imprimir(filas, 1); } catch (e) { r = { ok: false, error: e.message }; }

  const hora = new Date().toISOString();
  const impresora = config.bartender.impresora || "sin impresora elegida";
  for (const it of lista) {
    sesion.historial.unshift({ id: siguienteId++, hora, origen, ...it, ok: !!r.ok, error: r.ok ? null : (r.error || "sin detalle") });
    logImpresion(`${r.ok ? "IMPRESO" : "ERROR  "} | ${it.referencia} | ${it.nombre} | ${it.codigo_barras} | ${precioTxt(it.precio_detal)} | ` +
                 `${it.cantidad} etiqueta${it.cantidad > 1 ? "s" : ""} | ${impresora} | ${origen}${r.ok ? "" : " | " + (r.error || "sin detalle")}`);
  }
  return { ...r, total };
}

// Si no queda ninguna pestaña abierta durante este tiempo, el programa se cierra (y el historial se reinicia)
const SIN_VENTANAS_MS = Number(process.env.ALTA_SIN_VENTANAS_MS) || 10 * 60 * 1000;
let ultimoLatido = Date.now();
function cerrarPrograma(motivo) {
  logImpresion(`=== Programa cerrado (${motivo}). Historial de la sesión: ${sesion.historial.length} impresiones ===`);
  registrar(`Cerrado: ${motivo}`);
  setTimeout(() => process.exit(0), 200);
}
setInterval(() => {
  if (Date.now() - ultimoLatido > SIN_VENTANAS_MS) cerrarPrograma("sin ventanas abiertas");
}, Math.min(60000, SIN_VENTANAS_MS / 2)).unref();

// ---------- Siigo ----------
async function guardarPlantillaSiigo(buf, nombreOriginal) {
  const nombre = path.basename(nombreOriginal).replace(/[<>:"|?*]/g, "_");
  if (!/\.xlsx$/i.test(nombre)) throw new Error("La plantilla debe ser un archivo .xlsx.");
  try { await siigo.analizar(buf); } catch (e) { throw new Error("No pude leer la plantilla: " + e.message); }
  const destino = path.join(PLANTILLAS, nombre);
  if (fs.existsSync(destino) && !fs.readFileSync(destino).equals(buf)) fs.copyFileSync(destino, path.join(PLANTILLAS, `${nombre}.anterior_${sello()}`));
  escribirSeguro(destino, buf);
  config.siigo.plantilla = path.relative(RAIZ, destino);
  config.siigo.hoja = "";
  guardarConfig(config);
  return { ok: true, plantilla: nombre };
}
function rutaPlantillaSiigo() {
  if (config.siigo.plantilla && fs.existsSync(path.resolve(RAIZ, config.siigo.plantilla))) return path.resolve(RAIZ, config.siigo.plantilla);
  const f = fs.readdirSync(PLANTILLAS).filter(x => /\.xlsx$/i.test(x) && !x.startsWith("~$")).sort();
  return f.length ? path.join(PLANTILLAS, f[0]) : null;
}
async function infoSiigo() {
  const ruta = rutaPlantillaSiigo();
  if (!ruta) return { plantilla: null };
  const buf = fs.readFileSync(ruta);
  const a = await siigo.analizar(buf, config.siigo.hoja || undefined);
  return { plantilla: path.basename(ruta), buf, analisis: a, columnas: siigo.columnasEfectivas(a, config.siigo.columnas) };
}

// ---------- abrir carpetas y navegador ----------
function abrirCarpeta(ruta) { spawn("explorer.exe", [ruta], { detached: true, stdio: "ignore" }).unref(); }
function abrirNavegador(url) {
  if (process.env.ALTA_NO_ABRIR) return;
  spawn("rundll32.exe", ["url.dll,FileProtocolHandler", url], { detached: true, stdio: "ignore" }).unref();
}

// ---------- HTTP ----------
const TIPOS = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".wasm": "application/wasm",
                ".css": "text/css; charset=utf-8", ".ico": "image/x-icon" };
function responder(res, estado, cuerpo, tipo) {
  const buf = Buffer.isBuffer(cuerpo) ? cuerpo : Buffer.from(typeof cuerpo === "string" ? cuerpo : JSON.stringify(cuerpo));
  res.writeHead(estado, { "Content-Type": tipo || "application/json; charset=utf-8", "Cache-Control": "no-store", "Content-Length": buf.length });
  res.end(buf);
}
function leerCuerpo(req, limite = 200 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const partes = []; let n = 0;
    req.on("data", d => { n += d.length; if (n > limite) { reject(new Error("Archivo demasiado grande.")); req.destroy(); } else partes.push(d); });
    req.on("end", () => resolve(Buffer.concat(partes)));
    req.on("error", reject);
  });
}
async function leerJson(req) { const b = await leerCuerpo(req); return b.length ? JSON.parse(b.toString("utf8")) : {}; }

let puerto = null;
async function manejar(req, res) {
  const url = new URL(req.url, "http://127.0.0.1");
  // Solo se acepta el acceso por 127.0.0.1 o localhost (evita que otra página use el servidor)
  const host = String(req.headers.host || "");
  if (host !== `127.0.0.1:${puerto}` && host !== `localhost:${puerto}`) return responder(res, 403, { error: "Host no permitido" });

  if (url.pathname === "/api/ping") return responder(res, 200, { app: ID_APP, firma: FIRMA, raiz: RAIZ });

  if (!url.pathname.startsWith("/api/")) {
    if (req.method !== "GET") return responder(res, 405, { error: "Método no permitido" });
    const rel = url.pathname === "/" ? "index.html" : decodeURIComponent(url.pathname).replace(/^\/+/, "");
    const archivo = path.resolve(PUBLICO, rel);
    if (!archivo.startsWith(PUBLICO + path.sep) || !fs.existsSync(archivo) || fs.statSync(archivo).isDirectory()) {
      return responder(res, 404, "No encontrado", "text/plain; charset=utf-8");
    }
    let cuerpo = fs.readFileSync(archivo);
    if (archivo.endsWith(".html")) cuerpo = Buffer.from(cuerpo.toString("utf8").replace("__TOKEN__", TOKEN), "utf8");
    return responder(res, 200, cuerpo, TIPOS[path.extname(archivo)] || "application/octet-stream");
  }

  if (req.headers["x-token"] !== TOKEN) return responder(res, 403, { error: "Sesión no válida: recarga la página." });
  const ruta = `${req.method} ${url.pathname}`;
  ultimoLatido = Date.now();

  if (ruta === "POST /api/latido") return responder(res, 200, { ok: true });
  if (ruta === "GET /api/sesion") return responder(res, 200, sesion);
  if (ruta === "POST /api/sesion/producto") {
    const { producto } = await leerJson(req);
    const p = limpiarItem({ ...producto, cantidad: 1 });
    if (!p.referencia) return responder(res, 400, { error: "Falta la referencia." });
    delete p.cantidad;
    sesion.productos = sesion.productos.filter(x => x.referencia !== p.referencia);
    sesion.productos.push({ ...p, creado: new Date().toISOString() });
    return responder(res, 200, { ok: true });
  }
  if (ruta === "POST /api/sesion/reiniciar") {
    logImpresion(`=== Sesión reiniciada (Ctrl+F7). Se borraron ${sesion.productos.length} productos y ${sesion.historial.length} impresiones del historial ===`);
    sesion.inicio = new Date().toISOString();
    sesion.productos = [];
    sesion.historial = [];
    return responder(res, 200, { ok: true });
  }
  if (ruta === "POST /api/sesion/quitar") {
    const { referencia } = await leerJson(req);
    sesion.productos = sesion.productos.filter(x => x.referencia !== referencia);
    return responder(res, 200, { ok: true });
  }
  if (ruta === "POST /api/imprimir") {
    const { items, origen } = await leerJson(req);
    try { return responder(res, 200, await imprimirItems(items, origen === "lista" ? "botón Etiqueta" : origen === "reimpresion" ? "reimpresión" : "pestaña Imprimir")); }
    catch (e) { return responder(res, 400, { error: e.message }); }
  }

  if (ruta === "GET /api/estado") {
    const rb = rutaBase();
    const copias = fs.readdirSync(COPIAS).filter(f => f.endsWith(".sqlite")).sort();
    return responder(res, 200, { rutaBase: rb, existe: fs.existsSync(rb), copias: copias.length,
                                 ultimaCopia: copias[copias.length - 1] || null, raiz: RAIZ });
  }
  if (ruta === "GET /api/base") {
    const rb = rutaBase();
    res.setHeader("X-Version", versionArchivo(rb));
    if (!fs.existsSync(rb)) return responder(res, 200, Buffer.alloc(0), "application/octet-stream");
    return responder(res, 200, fs.readFileSync(rb), "application/octet-stream");
  }
  if (ruta === "PUT /api/base") {
    const r = guardarBase(await leerCuerpo(req), url.searchParams.get("version"), url.searchParams.get("reemplazo") === "1");
    if (r.estado !== 200) registrar(`Guardado rechazado: ${r.error}`);
    return responder(res, r.estado, r);
  }
  if (ruta === "POST /api/copia") {
    const d = hacerCopia("manual");
    return responder(res, d ? 200 : 400, d ? { archivo: path.basename(d) } : { error: "Aún no hay base guardada." });
  }
  if (ruta === "GET /api/config") return responder(res, 200, config);
  if (ruta === "PUT /api/config") {
    const nuevo = await leerJson(req);
    const anteriorBase = rutaBase();
    config = {
      ...config,
      bartender: { ...config.bartender, ...(nuevo.bartender || {}) },
      siigo: { ...config.siigo, ...(nuevo.siigo || {}) },
      barras: (() => {
        // El último código solo sube: nunca se guarda uno menor al que ya estaba
        const actual = config.barras || CONFIG_DEFECTO.barras;
        const n = Math.floor(Number(nuevo.barras && nuevo.barras.ultimoCorto));
        const forzar = !!(nuevo.barras && nuevo.barras.forzar);   // corrección manual desde «Seguir después del número»
        if (!Number.isFinite(n) || n < 0 || (!forzar && n <= (actual.ultimoCorto || 0))) return actual;
        return { ultimoCorto: n, ultimoCodigo: String(n).padStart(6, "0"), fecha: new Date().toISOString() };
      })(),
      ...(nuevo.rutaBase !== undefined ? { rutaBase: String(nuevo.rutaBase).trim() } : {})
    };
    if (nuevo.bartender && nuevo.bartender.impresora !== undefined) recordarImpresora(nuevo.bartender.impresora);
    // Cambio de ubicación: si en la nueva ruta no hay base, se copia la actual allá
    const nb = rutaBase();
    let nota = null;
    if (nb !== anteriorBase) {
      if (fs.existsSync(nb)) nota = "En la nueva ubicación ya había una base: se abrió esa.";
      else if (fs.existsSync(anteriorBase)) { escribirSeguro(nb, fs.readFileSync(anteriorBase)); nota = "La base actual se copió a la nueva ubicación."; }
    }
    guardarConfig(config);
    return responder(res, 200, { config, nota, cambioBase: nb !== anteriorBase });
  }
  if (ruta === "POST /api/carpeta") {
    const { cual } = await leerJson(req);
    const destinos = { copias: COPIAS, exportados: EXPORTADOS, plantillas: PLANTILLAS, datos: path.dirname(rutaBase()), etiquetas: ETIQUETAS, configuracion: CONFIG_DIR };
    if (cual === "log") {
      if (!fs.existsSync(LOG_IMPRESIONES)) fs.writeFileSync(LOG_IMPRESIONES, "");
      spawn("notepad.exe", [LOG_IMPRESIONES], { detached: true, stdio: "ignore" }).unref();
      return responder(res, 200, { ok: true });
    }
    if (!destinos[cual]) return responder(res, 400, { error: "Carpeta desconocida" });
    abrirCarpeta(destinos[cual]);
    return responder(res, 200, { ok: true });
  }
  if (ruta === "GET /api/siigo") {
    const i = await infoSiigo();
    if (!i.plantilla) return responder(res, 200, { plantilla: null, campos: siigo.CAMPOS });
    return responder(res, 200, { plantilla: i.plantilla, hoja: i.analisis.hoja, hojas: i.analisis.hojas, fila: i.analisis.filaEncabezados,
                                 deteccion: i.analisis.deteccion, columnas: i.columnas, campos: siigo.CAMPOS });
  }
  if (ruta === "POST /api/siigo/plantilla") {
    try { return responder(res, 200, await guardarPlantillaSiigo(await leerCuerpo(req), decodeURIComponent(String(req.headers["x-nombre"] || "plantilla_siigo.xlsx")))); }
    catch (e) { return responder(res, 400, { error: e.message }); }
  }
  if (ruta === "POST /api/elegir") {
    const { tipo } = await leerJson(req);
    try {
      const elegido = await elegirArchivo(tipo);
      if (!elegido) return responder(res, 200, { cancelado: true });
      const r = { ruta: elegido, nombre: path.basename(elegido) };
      if (tipo === "plantillaSiigo") Object.assign(r, await guardarPlantillaSiigo(fs.readFileSync(elegido), path.basename(elegido)));
      return responder(res, 200, r);
    } catch (e) { return responder(res, 400, { error: e.message }); }
  }
  if (ruta === "GET /api/elegido") {
    const elegido = elegidos[url.searchParams.get("tipo")];
    if (!elegido || !fs.existsSync(elegido)) return responder(res, 404, { error: "Archivo no encontrado." });
    return responder(res, 200, fs.readFileSync(elegido), "application/octet-stream");
  }
  if (ruta === "GET /api/impresoras") {
    return responder(res, 200, { instaladas: await impresorasInstaladas(), usadas: config.impresorasUsadas });
  }
  if (ruta === "POST /api/siigo/exportar") {
    const { productos, soloValidar, soloValidos } = await leerJson(req);
    const i = await infoSiigo();
    if (!i.plantilla) return responder(res, 400, { error: "Primero carga la plantilla de Siigo (.xlsx)." });
    if (!Array.isArray(productos) || !productos.length) return responder(res, 400, { error: "No hay productos para exportar." });
    const sinColumna = i.columnas.filter(k => k.obligatoria && !k.campo).map(k => k.titulo);
    const { validos, fallan } = siigo.validar(productos, i.columnas);
    const informe = { total: productos.length, validos: validos.length, fallan, sinColumna, lotes: Math.ceil(validos.length / siigo.MAX_POR_ARCHIVO) };
    if (soloValidar || !validos.length || (fallan.length && !soloValidos)) return responder(res, 200, { ...informe, archivos: [] });
    const archivos = await siigo.exportar(i.buf, i.analisis, i.columnas, validos);
    const base = `Siigo_productos_${sello()}`;
    const salida = archivos.map((a, n) => {
      const nombre = archivos.length > 1 ? `${base}_lote${n + 1}.xlsx` : `${base}.xlsx`;
      escribirSeguro(path.join(EXPORTADOS, nombre), a.buffer);
      return { nombre, filas: a.filas, sobrantes: a.sobrantes };
    });
    registrar(`Exportados ${validos.length} productos a Siigo en ${salida.length} archivo(s)`);
    return responder(res, 200, { ...informe, archivos: salida, carpeta: EXPORTADOS, exportados: validos.map(p => p.codigo_interno) });
  }
  if (ruta === "POST /api/salir") {
    const { motivo } = await leerJson(req);
    responder(res, 200, { ok: true });
    cerrarPrograma(motivo === "actualizacion" ? "se abrió una versión actualizada del programa" : "botón Cerrar el programa");
    return;
  }
  return responder(res, 404, { error: "Ruta desconocida" });
}

// ---------- arranque ----------
// Devuelve { estado: "libre" | "nuestra" | "otra", firma }. "nuestra" = esta misma carpeta de la app.
function ping(p) {
  return new Promise(resolve => {
    const r = http.get({ host: "127.0.0.1", port: p, path: "/api/ping", timeout: 800 }, res => {
      let t = ""; res.on("data", d => { t += d; });
      res.on("end", () => {
        try {
          const j = JSON.parse(t);
          resolve(j.app === ID_APP && path.resolve(j.raiz || "") === RAIZ ? { estado: "nuestra", firma: j.firma } : { estado: "otra" });
        } catch (e) { resolve({ estado: "otra" }); }
      });
    });
    r.on("timeout", () => { r.destroy(); resolve({ estado: "otra" }); });
    r.on("error", e => resolve({ estado: e.code === "ECONNREFUSED" ? "libre" : "otra" }));
  });
}
// Pide al proceso viejo que se cierre (con la clave que dejó en datos.token) y espera a que libere el puerto
async function cerrarViejo(p) {
  try {
    const token = fs.readFileSync(path.join(DATOS, ".token"), "utf8").trim();
    await fetch(`http://127.0.0.1:${p}/api/salir`, { method: "POST", headers: { "X-Token": token, "Content-Type": "application/json" },
                                                       body: JSON.stringify({ motivo: "actualizacion" }) });
  } catch (e) { registrar(`No se pudo cerrar la versión anterior: ${e.message}`); }
  for (let i = 0; i < 25; i++) {
    await new Promise(r => setTimeout(r, 200));
    if ((await ping(p)).estado === "libre") return true;
  }
  return false;
}
function escuchar(servidor, p) {
  return new Promise(resolve => {
    servidor.once("error", () => resolve(false));
    servidor.listen(p, "127.0.0.1", () => resolve(true));
  });
}

async function iniciar() {
  const base = Number(process.env.ALTA_PUERTO || config.puerto) || 3789;
  for (let p = base; p < base + 10; p++) {
    let { estado, firma } = await ping(p);
    if (estado === "nuestra" && firma !== FIRMA) {
      console.log(`Versión anterior abierta en el puerto ${p}: se cierra para abrir la actualizada`);
      if (await cerrarViejo(p)) estado = "libre";
    }
    if (estado === "nuestra") { abrirNavegador(`http://127.0.0.1:${p}/`); console.log(`Ya estaba abierto en el puerto ${p}`); return; }
    if (estado !== "libre") continue;
    const servidor = http.createServer((req, res) => manejar(req, res).catch(e => {
      registrar(`Error ${req.method} ${req.url}: ${e.stack || e}`);
      if (!res.headersSent) responder(res, 500, { error: e.message });
    }));
    if (!(await escuchar(servidor, p))) continue;
    puerto = p;
    try { fs.writeFileSync(path.join(DATOS, ".token"), TOKEN); } catch (e) { registrar(`No se pudo guardar la clave: ${e.message}`); }
    if (fs.existsSync(rutaBase())) hacerCopia("al-abrir");
    console.log(`Alta de productos en http://127.0.0.1:${p}/`);
    registrar(`Iniciado en el puerto ${p}. Base: ${rutaBase()}`);
    logImpresion("=== Programa abierto: historial de impresión nuevo ===");
    abrirNavegador(`http://127.0.0.1:${p}/`);
    return;
  }
  registrar("No hay puerto libre entre " + base + " y " + (base + 9));
  process.exit(1);
}
iniciar();
