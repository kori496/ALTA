// Excel para Siigo: lee la plantilla oficial (.xlsx) y escribe los productos en ella
// tocando solo las celdas de datos. Estilos, validaciones, listas, hojas ocultas y todo lo
// demás del archivo se copian tal cual: el .xlsx se abre como zip y solo se editan
// la hoja de datos, sharedStrings.xml y, si existe, la tabla de esa hoja.
const JSZip = require("jszip");

const MAX_POR_ARCHIVO = 500;

// Campos de la app que se pueden poner en una columna de la plantilla
const CAMPOS = {
  referencia: "Referencia",
  codigo_barras: "Código de barras",
  nombre: "Nombre",
  grupo: "Grupo",
  unidad: "Unidad de medida",
  impuesto: "Tipo de impuesto",
  tarifa_iva: "Tarifa IVA (19, 5 o 0)",
  costo: "Costo sin IVA",
  precio_detal: "Precio detal (con IVA)",
  base_detal: "Precio detal sin IVA",
  precio_mayorista: "Precio mayorista (con IVA)",
  inc: "INC %"
};
const NUMERICOS = new Set(["tarifa_iva", "costo", "precio_detal", "base_detal", "precio_mayorista", "inc"]);
const TARIFA_IVA = { "IVA 19%": 0.19, "IVA 5%": 0.05, "Exento": 0, "Excluido": 0, "No gravado": 0 };

// ---------- utilidades ----------
function norm(s) {
  return String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}
function xmlDec(s) {
  return s.replace(/&(lt|gt|quot|apos|amp|#x[0-9a-f]+|#\d+);/gi, (m, e) => {
    const k = e.toLowerCase();
    if (k === "lt") return "<"; if (k === "gt") return ">"; if (k === "quot") return '"';
    if (k === "apos") return "'"; if (k === "amp") return "&";
    return String.fromCodePoint(k[1] === "x" ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10));
  });
}
function xmlEsc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
}
function attr(tag, nombre) {
  const m = tag.match(new RegExp(`\\s${nombre}="([^"]*)"`));
  return m ? xmlDec(m[1]) : null;
}
function setAttr(tag, nombre, valor) {
  const re = new RegExp(`\\s${nombre}="[^"]*"`);
  if (valor === null) return tag.replace(re, "");
  if (re.test(tag)) return tag.replace(re, ` ${nombre}="${xmlEsc(valor)}"`);
  return tag.replace(/^<([\w:]+)/, `<$1 ${nombre}="${xmlEsc(valor)}"`);
}
function colIdx(letras) {
  let n = 0;
  for (const ch of letras.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function colLetras(i) {
  let s = "";
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}
function partirRef(ref) {
  const m = String(ref).replace(/\$/g, "").match(/^([A-Z]+)(\d+)$/i);
  return m ? { c: colIdx(m[1]), r: Number(m[2]) } : null;
}
function rangos(sqref) {
  return String(sqref).trim().split(/\s+/).map(p => {
    const [a, b] = p.split(":");
    const x = partirRef(a), y = partirRef(b || a);
    return x && y ? { c1: x.c, r1: x.r, c2: y.c, r2: y.r } : null;
  }).filter(Boolean);
}
function rutaZip(base, destino) {
  if (destino.startsWith("/")) return destino.slice(1);
  const partes = base.split("/"); partes.pop();
  for (const p of destino.split("/")) {
    if (p === "..") partes.pop(); else if (p !== ".") partes.push(p);
  }
  return partes.join("/");
}
function relsDe(ruta) {
  const i = ruta.lastIndexOf("/");
  return `${ruta.slice(0, i)}/_rels/${ruta.slice(i + 1)}.rels`;
}
async function leerRels(zip, ruta) {
  const f = zip.file(relsDe(ruta));
  if (!f) return [];
  const xml = await f.async("string");
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map(m => ({
    id: attr(m[0], "Id"), tipo: attr(m[0], "Type") || "", destino: rutaZip(ruta, attr(m[0], "Target") || "")
  }));
}

// ---------- colores (para saber qué encabezados son obligatorios) ----------
const INDEXADOS = ["000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF",
  "000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF","800000","008000","000080",
  "808000","800080","008080","C0C0C0","808080","9999FF","993366","FFFFCC","CCFFFF","660066","FF8080",
  "0066CC","CCCCFF","000080","FF00FF","FFFF00","00FFFF","800080","800000","008080","0000FF","00CCFF",
  "CCFFFF","CCFFCC","FFFF99","99CCFF","FF99CC","CC99FF","FFCC99","3366FF","33CCCC","99CC00","FFCC00",
  "FF9900","FF6600","666699","969696","003366","339966","003300","333300","993300","993366","333399","333333"];

function aplicarTinte(hex, tinte) {
  if (!tinte) return hex;
  return [0, 2, 4].map(i => {
    let v = parseInt(hex.slice(i, i + 2), 16);
    v = tinte < 0 ? v * (1 + tinte) : v + (255 - v) * tinte;
    return Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, "0");
  }).join("").toUpperCase();
}
function colorHex(tagColor, tema) {
  if (!tagColor) return null;
  const tinte = Number(attr(tagColor, "tint") || 0);
  const rgb = attr(tagColor, "rgb");
  if (rgb) return aplicarTinte(rgb.slice(-6).toUpperCase(), tinte);
  const th = attr(tagColor, "theme");
  if (th !== null && tema[Number(th)]) return aplicarTinte(tema[Number(th)], tinte);
  const ix = attr(tagColor, "indexed");
  if (ix !== null && INDEXADOS[Number(ix)]) return INDEXADOS[Number(ix)];
  return null;
}
function claridad(hex) {
  const [r, g, b] = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16) / 255);
  return (Math.max(r, g, b) + Math.min(r, g, b)) / 2;
}

async function leerEstilos(zip) {
  // Tema: orden de índices de Excel (0 lt1, 1 dk1, 2 lt2, 3 dk2, 4..9 accent, 10 hlink, 11 folHlink)
  const tema = [];
  const ft = zip.file("xl/theme/theme1.xml");
  if (ft) {
    const xml = await ft.async("string");
    const esquema = xml.match(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/);
    if (esquema) {
      const col = {};
      for (const m of esquema[0].matchAll(/<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>([\s\S]*?)<\/a:\1>/g)) {
        const v = m[2].match(/(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"/);
        if (v) col[m[1]] = v[1].toUpperCase();
      }
      ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
        .forEach((k, i) => { tema[i] = col[k] || null; });
    }
  }
  const fs = zip.file("xl/styles.xml");
  if (!fs) return { xfFill: [], fills: [], xfNumFmt: [] };
  const xml = await fs.async("string");
  const fillsXml = (xml.match(/<fills\b[\s\S]*?<\/fills>/) || [""])[0];
  const fills = [...fillsXml.matchAll(/<fill\b[^>]*?(?:\/>|>[\s\S]*?<\/fill>)/g)].map(m => {
    const pf = m[0].match(/<patternFill\b[^>]*?(?:\/>|>[\s\S]*?<\/patternFill>)/);
    if (!pf || (attr(pf[0], "patternType") || "none") === "none") return null;
    const fg = pf[0].match(/<fgColor\b[^>]*>/);
    return colorHex(fg && fg[0], tema);
  });
  const xfsXml = (xml.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/) || [""])[0];
  const xfs = [...xfsXml.matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)];
  return {
    fills,
    xfFill: xfs.map(m => Number(attr(m[0], "fillId") || 0)),
    xfNumFmt: xfs.map(m => Number(attr(m[0], "numFmtId") || 0))
  };
}

// ---------- lectura de la plantilla ----------
async function leerCompartidos(zip) {
  const f = zip.file("xl/sharedStrings.xml");
  if (!f) return null;
  const xml = await f.async("string");
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g)].map(m =>
    m[1] ? [...m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, "").matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t\/>/g)]
      .map(t => xmlDec(t[1] || "")).join("") : "");
}

function valorCelda(xmlCelda, compartidos) {
  const t = attr(xmlCelda.match(/^<c\b[^>]*>/)[0], "t");
  if (t === "inlineStr") {
    return [...xmlCelda.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(x => xmlDec(x[1])).join("");
  }
  const v = xmlCelda.match(/<v>([\s\S]*?)<\/v>/);
  if (!v) return "";
  if (t === "s") return compartidos ? (compartidos[Number(v[1])] ?? "") : "";
  return xmlDec(v[1]);
}

// Separa sheetData en filas y celdas, guardando el XML original de cada una
function partirHoja(xml) {
  const ini = xml.search(/<sheetData\b/);
  const finVacio = xml.slice(ini).match(/^<sheetData\b[^>]*\/>/);
  let abre, cuerpo, fin;
  if (finVacio) {
    abre = finVacio[0].replace(/\/>$/, ">"); cuerpo = ""; fin = ini + finVacio[0].length;
  } else {
    const a = xml.slice(ini).match(/^<sheetData\b[^>]*>/)[0];
    const cierre = xml.indexOf("</sheetData>", ini);
    abre = a; cuerpo = xml.slice(ini + a.length, cierre); fin = cierre + "</sheetData>".length;
  }
  const filas = new Map();
  for (const m of cuerpo.matchAll(/<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const tag = m[0].match(/^<row\b[^>]*?\/?>/)[0];
    const r = Number(attr(tag, "r"));
    const celdas = new Map();
    for (const c of (m[1] || "").matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) {
      const p = partirRef(attr(c[0], "r"));
      if (p) celdas.set(p.c, c[0]);
    }
    filas.set(r, { tag: tag.replace(/\/>$/, ">"), celdas, cambiada: false });
  }
  return { antes: xml.slice(0, ini), abre, despues: xml.slice(fin), filas };
}

function armarHoja(h) {
  const filas = [...h.filas.entries()].sort((a, b) => a[0] - b[0]).map(([, f]) => {
    let tag = f.tag;
    if (f.cambiada) tag = setAttr(tag, "spans", null);
    if (!f.celdas.size) return tag.replace(/>$/, "/>");
    const celdas = [...f.celdas.entries()].sort((a, b) => a[0] - b[0]).map(([, c]) => c).join("");
    return tag + celdas + "</row>";
  }).join("");
  return h.antes + h.abre + filas + "</sheetData>" + h.despues;
}

function estilosColumnas(xmlHoja) {
  const cols = [];
  for (const m of xmlHoja.matchAll(/<col\b[^>]*>/g)) {
    const s = attr(m[0], "style");
    if (s !== null) cols.push({ min: Number(attr(m[0], "min")) - 1, max: Number(attr(m[0], "max")) - 1, s });
  }
  return c => (cols.find(x => c >= x.min && c <= x.max) || {}).s ?? null;
}

// Validaciones de lista: normales y las de la extensión x14 (listas que apuntan a otra hoja)
function validacionesLista(xmlHoja) {
  const lista = [];
  for (const m of xmlHoja.matchAll(/<dataValidation\b[^>]*?(?:\/>|>([\s\S]*?)<\/dataValidation>)/g)) {
    const tag = m[0].match(/^<dataValidation\b[^>]*>/)[0];
    if (attr(tag, "type") !== "list") continue;
    const f = (m[1] || "").match(/<formula1>([\s\S]*?)<\/formula1>/);
    if (f) lista.push({ rangos: rangos(attr(tag, "sqref") || ""), formula: xmlDec(f[1]) });
  }
  for (const m of xmlHoja.matchAll(/<x14:dataValidation\b[^>]*>([\s\S]*?)<\/x14:dataValidation>/g)) {
    const tag = m[0].match(/^<x14:dataValidation\b[^>]*>/)[0];
    if (attr(tag, "type") !== "list") continue;
    const f = m[1].match(/<xm:f>([\s\S]*?)<\/xm:f>/);
    const sq = m[1].match(/<xm:sqref>([\s\S]*?)<\/xm:sqref>/);
    if (f && sq) lista.push({ rangos: rangos(sq[1]), formula: xmlDec(f[1]) });
  }
  return lista;
}

async function abrir(buf) {
  const zip = await JSZip.loadAsync(buf);
  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const rels = await leerRels(zip, "xl/workbook.xml");
  const hojas = [...wbXml.matchAll(/<sheet\b[^>]*>/g)].map(m => {
    const rid = attr(m[0], "r:id");
    const rel = rels.find(r => r.id === rid);
    return { nombre: attr(m[0], "name"), oculta: !!attr(m[0], "state") && attr(m[0], "state") !== "visible",
             ruta: rel ? rel.destino : null };
  }).filter(h => h.ruta && zip.file(h.ruta));
  const nombres = {};
  for (const m of wbXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    nombres[norm(attr("<x" + m[1] + ">", "name"))] = xmlDec(m[2]);
  }
  return { zip, hojas, nombres, compartidos: await leerCompartidos(zip), estilos: await leerEstilos(zip) };
}

// Valores permitidos de una lista de validación, o null si no se pueden leer
async function resolverLista(libro, formula, cacheHojas) {
  let f = formula.trim().replace(/^=/, "");
  if (/^".*"$/.test(f)) return f.slice(1, -1).split(",").map(s => s.trim()).filter(Boolean);
  if (libro.nombres[norm(f)] !== undefined) f = libro.nombres[norm(f)].replace(/^=/, "");
  const m = f.match(/^(?:'((?:[^']|'')+)'|([^!'"]+))!(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?$/i);
  if (!m) return null; // INDIRECTO, DESREF u otras fórmulas: no se validan
  const nombreHoja = (m[1] || m[2]).replace(/''/g, "'");
  const hoja = libro.hojas.find(h => h.nombre === nombreHoja);
  if (!hoja) return null;
  if (!cacheHojas[hoja.ruta]) cacheHojas[hoja.ruta] = partirHoja(await libro.zip.file(hoja.ruta).async("string"));
  const a = partirRef(m[3]), b = partirRef(m[4] || m[3]);
  const valores = [];
  for (const [r, fila] of cacheHojas[hoja.ruta].filas) {
    if (r < a.r || r > b.r) continue;
    for (const [c, xmlC] of fila.celdas) {
      if (c < a.c || c > b.c) continue;
      const v = valorCelda(xmlC, libro.compartidos).trim();
      if (v) valores.push(v);
    }
  }
  return valores;
}

// Busca la hoja y la fila de encabezados: la fila con más textos entre las primeras 30
async function analizar(buf, nombreHoja) {
  const libro = await abrir(buf);
  const candidatas = nombreHoja ? libro.hojas.filter(h => h.nombre === nombreHoja) : libro.hojas.filter(h => !h.oculta);
  if (!candidatas.length) throw new Error(nombreHoja ? `La plantilla no tiene la hoja «${nombreHoja}».` : "La plantilla no tiene hojas visibles.");

  let mejor = null;
  for (const hoja of candidatas) {
    const xml = await libro.zip.file(hoja.ruta).async("string");
    const h = partirHoja(xml);
    for (const [r, fila] of h.filas) {
      if (r > 30) continue;
      const textos = [...fila.celdas.values()].filter(c => /[a-z]/i.test(valorCelda(c, libro.compartidos))).length;
      if (textos >= 2 && (!mejor || textos > mejor.textos)) mejor = { hoja, xml, h, r, textos };
    }
  }
  if (!mejor) throw new Error("No encontré la fila de encabezados en la plantilla.");

  const { hoja, xml, h, r } = mejor;
  const cache = { [hoja.ruta]: h };
  const listas = validacionesLista(xml);
  const columnas = [];
  for (const [c, xmlC] of [...h.filas.get(r).celdas.entries()].sort((a, b) => a[0] - b[0])) {
    const titulo = valorCelda(xmlC, libro.compartidos).trim();
    if (!titulo) continue;
    const s = Number(attr(xmlC, "s") || 0);
    const color = libro.estilos.fills[libro.estilos.xfFill[s]] || null;
    const val = listas.find(v => v.rangos.some(g => c >= g.c1 && c <= g.c2 && r + 1 >= g.r1 && r + 1 <= g.r2));
    columnas.push({
      c, letra: colLetras(c), titulo, color,
      lista: val ? await resolverLista(libro, val.formula, cache) : null,
      formulaLista: val ? val.formula : null,
      marcaAsterisco: /\*/.test(titulo) || /obligatori/i.test(titulo)
    });
  }

  // Obligatorias: los encabezados más oscuros (azul oscuro) frente a los claros
  const claros = columnas.filter(k => k.color).map(k => claridad(k.color));
  const min = Math.min(...claros), max = Math.max(...claros);
  const porColor = claros.length >= 2 && max - min > 0.12;
  columnas.forEach(k => {
    k.obligatoriaPorColor = porColor && !!k.color && claridad(k.color) < (min + max) / 2;
    k.obligatoria = k.obligatoriaPorColor || k.marcaAsterisco;
    k.sugerido = sugerirCampo(k.titulo);
  });

  return {
    hoja: hoja.nombre, filaEncabezados: r, columnas,
    deteccion: porColor
      ? "Obligatorias detectadas por el color oscuro del encabezado."
      : "No se distinguen dos colores en los encabezados: revisa a mano cuáles son obligatorias.",
    hojas: libro.hojas.map(x => x.nombre)
  };
}

// Sugerencia de campo según el nombre del encabezado (se puede cambiar en la app)
function sugerirCampo(titulo) {
  const t = norm(titulo);
  const tiene = (...p) => p.some(x => new RegExp(`(^| )${x}( |$)`).test(t));
  if (tiene("incluye")) return ""; // columnas Sí/No: se llenan con un valor fijo
  if (tiene("barras", "ean", "gtin")) return "codigo_barras";
  if (tiene("costo")) return "costo";
  if (tiene("precio", "valor")) {
    if (tiene("mayor", "mayorista", "2")) return "precio_mayorista";
    if (tiene("sin iva", "base")) return "base_detal";
    return "precio_detal";
  }
  if (tiene("inc", "impoconsumo")) return "inc";
  if (tiene("tarifa") && tiene("iva")) return "tarifa_iva";
  if (tiene("impuesto", "iva") && !tiene("incluye")) return "impuesto";
  if (tiene("unidad")) return "unidad";
  if (tiene("grupo", "categoria", "linea", "clasificacion", "familia")) return "grupo";
  if (tiene("nombre") || (tiene("descripcion") && !tiene("larga"))) return "nombre";
  if (tiene("codigo", "referencia", "sku") && !tiene("fabrica")) return "referencia";
  return "";
}

// ---------- valores de cada producto ----------
function valorCampo(p, campo) {
  const tIva = TARIFA_IVA[p.impuesto] ?? 0;
  switch (campo) {
    case "referencia": return p.codigo_interno;
    case "tarifa_iva": return Math.round(tIva * 100);
    case "base_detal":
      return p.precio_detal == null ? null
        : Math.round(p.precio_detal / (1 + tIva + (p.inc || 0) / 100) * 100) / 100;
    case "unidad": return p.unidad || "Unidad";
    default: return p[campo];
  }
}

// cfgCol: { campo, fijo, obligatoria, equivalencias: { "valor app": "valor Siigo" } }
function valorColumna(p, cfgCol) {
  let v = cfgCol.campo === "_fijo" ? cfgCol.fijo : (cfgCol.campo ? valorCampo(p, cfgCol.campo) : null);
  if (v === undefined || v === null || v === "") return null;
  const eq = cfgCol.equivalencias || {};
  if (typeof v === "string" && eq[v] !== undefined) v = eq[v];
  if (cfgCol.campo !== "_fijo" && NUMERICOS.has(cfgCol.campo) && typeof v === "number") return v;
  return String(v).trim() === "" ? null : String(v).trim();
}

// Une el análisis de la plantilla con la configuración guardada
function columnasEfectivas(analisis, cfgColumnas) {
  return analisis.columnas.map(k => {
    const g = (cfgColumnas || {})[k.titulo] || {};
    return {
      ...k,
      campo: g.campo !== undefined ? g.campo : k.sugerido,
      fijo: g.fijo ?? "",
      equivalencias: g.equivalencias || {},
      obligatoria: g.obligatoria !== undefined ? !!g.obligatoria : k.obligatoria
    };
  });
}

// Revisa cada producto antes de exportar. Devuelve { validos, fallan: [{ ref, nombre, problemas }] }
function validar(productos, columnas) {
  const validos = [], fallan = [];
  for (const p of productos) {
    const problemas = [];
    for (const k of columnas) {
      const v = valorColumna(p, k);
      if (v === null) {
        if (k.obligatoria) problemas.push(`falta «${k.titulo}»`);
        continue;
      }
      if (k.lista && k.lista.length && !k.lista.some(x => norm(x) === norm(v))) {
        problemas.push(`«${k.titulo}»: «${v}» no está en la lista de Siigo`);
      }
    }
    if (problemas.length) fallan.push({ ref: p.codigo_interno, nombre: p.nombre, problemas });
    else validos.push(p);
  }
  return { validos, fallan };
}

// ---------- escritura ----------
async function escribirLote(buf, analisis, columnas, productos) {
  const libro = await abrir(buf);
  const hoja = libro.hojas.find(x => x.nombre === analisis.hoja);
  const xml = await libro.zip.file(hoja.ruta).async("string");
  const h = partirHoja(xml);
  const estiloCol = estilosColumnas(xml);

  // Textos nuevos al final de sharedStrings (sin tocar los que ya existen)
  const compartidos = libro.compartidos;
  const indice = new Map((compartidos || []).map((s, i) => [s, i]));
  const nuevos = [];
  let usos = 0;
  function celdaTexto(ref, s, texto) {
    if (!compartidos) {
      return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(texto)}</t></is></c>`;
    }
    let i = indice.get(texto);
    if (i === undefined) {
      i = compartidos.length + nuevos.length;
      indice.set(texto, i);
      nuevos.push(texto);
    }
    usos++;
    return `<c r="${ref}"${s} t="s"><v>${i}</v></c>`;
  }

  const r0 = analisis.filaEncabezados;
  productos.forEach((p, n) => {
    const r = r0 + 1 + n;
    if (!h.filas.has(r)) h.filas.set(r, { tag: `<row r="${r}">`, celdas: new Map(), cambiada: true });
    const fila = h.filas.get(r);
    fila.cambiada = true;
    const estiloFila = attr(fila.tag, "customFormat") === "1" ? attr(fila.tag, "s") : null;
    for (const k of columnas) {
      const ref = colLetras(k.c) + r;
      const previa = fila.celdas.get(k.c);
      const sv = (previa && attr(previa.match(/^<c\b[^>]*>/)[0], "s")) ?? estiloFila ?? estiloCol(k.c);
      const s = sv !== null && sv !== undefined ? ` s="${sv}"` : "";
      const v = valorColumna(p, k);
      if (v === null) {
        if (previa) fila.celdas.set(k.c, `<c r="${ref}"${s}/>`);
      } else if (typeof v === "number") {
        fila.celdas.set(k.c, `<c r="${ref}"${s}><v>${v}</v></c>`);
      } else {
        fila.celdas.set(k.c, celdaTexto(ref, s, v));
      }
    }
  });

  // Filas con datos que ya traía la plantilla después de las nuestras (aviso)
  const ultima = r0 + productos.length;
  const sobrantes = [...h.filas.entries()].filter(([r, f]) => r > ultima &&
    [...f.celdas.values()].some(c => valorCelda(c, compartidos) !== "")).length;

  let nuevoXml = armarHoja(h);
  // Dimensión de la hoja
  let maxR = 1, maxC = 0;
  for (const [r, f] of h.filas) {
    if (f.celdas.size) { maxR = Math.max(maxR, r); maxC = Math.max(maxC, ...f.celdas.keys()); }
  }
  nuevoXml = nuevoXml.replace(/<dimension\b[^>]*\/>/, d => {
    const desde = (attr(d, "ref") || "A1").split(":")[0];
    return `<dimension ref="${desde}:${colLetras(maxC)}${maxR}"/>`;
  });
  libro.zip.file(hoja.ruta, nuevoXml);

  if (nuevos.length) {
    const ruta = "xl/sharedStrings.xml";
    let sst = await libro.zip.file(ruta).async("string");
    const extra = nuevos.map(t => `<si><t xml:space="preserve">${xmlEsc(t)}</t></si>`).join("");
    sst = sst.replace(/<sst\b[^>]*?\/>/, m => m.replace(/\/>$/, "></sst>"));
    sst = sst.replace(/<\/sst>\s*$/, extra + "</sst>");
    sst = sst.replace(/<sst\b[^>]*>/, tag => {
      let t = tag;
      const uc = attr(t, "uniqueCount");
      if (uc !== null) t = setAttr(t, "uniqueCount", String(Number(uc) + nuevos.length));
      const ct = attr(t, "count");
      if (ct !== null) t = setAttr(t, "count", String(Number(ct) + usos));
      return t;
    });
    libro.zip.file(ruta, sst);
  } else if (usos && compartidos) {
    const ruta = "xl/sharedStrings.xml";
    let sst = await libro.zip.file(ruta).async("string");
    sst = sst.replace(/<sst\b[^>]*>/, tag => {
      const ct = attr(tag, "count");
      return ct !== null ? setAttr(tag, "count", String(Number(ct) + usos)) : tag;
    });
    libro.zip.file(ruta, sst);
  }

  // Si la hoja tiene una tabla de Excel que empieza en los encabezados, se alarga hasta la última fila
  for (const rel of await leerRels(libro.zip, hoja.ruta)) {
    if (!/\/table$/.test(rel.tipo) || !libro.zip.file(rel.destino)) continue;
    let t = await libro.zip.file(rel.destino).async("string");
    const tag = t.match(/<table\b[^>]*>/)[0];
    const [a, b] = (attr(tag, "ref") || "").split(":");
    const pa = partirRef(a), pb = partirRef(b || a);
    if (!pa || pa.r !== r0 || pb.r >= ultima) continue;
    const ref = `${a}:${b.replace(/\d+$/, String(ultima))}`;
    t = t.replace(tag, setAttr(tag, "ref", ref));
    t = t.replace(/<autoFilter\b[^>]*>/, af => setAttr(af, "ref", ref));
    libro.zip.file(rel.destino, t);
  }

  const buffer = await libro.zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return { buffer, sobrantes };
}

// Exporta en archivos de máximo 500 productos
async function exportar(buf, analisis, columnas, productos) {
  const lotes = [];
  for (let i = 0; i < productos.length; i += MAX_POR_ARCHIVO) lotes.push(productos.slice(i, i + MAX_POR_ARCHIVO));
  const archivos = [];
  for (const lote of lotes) archivos.push({ ...(await escribirLote(buf, analisis, columnas, lote)), filas: lote.length });
  return archivos;
}

module.exports = { CAMPOS, MAX_POR_ARCHIVO, analizar, columnasEfectivas, validar, exportar, sugerirCampo, valorColumna, norm };
