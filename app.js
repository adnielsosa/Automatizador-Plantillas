/* ============================================================================
   LOC_MINICENEFA_CENEFAFASTFOODKOXKALOC1_FASTFOODKOXKA(1PROMO)
   Motor compartido entre index.html (inyector) y render.html (resultado limpio).
   Requiere que default-assets.js se cargue ANTES que este archivo (define
   window.DEFAULT_ASSETS con los 8 assets por defecto en base64).

   Estado -> localStorage (misma clave, mismo origen) -> ambas páginas leen de
   ahí para poder mostrar exactamente lo mismo. Si tu navegador restringe
   localStorage al abrir los .html por doble clic (file://), sirve la carpeta
   con "python3 -m http.server" y ábrelos por http://localhost:8000/.
   ========================================================================== */

// v5: TEMPLATE ahora se reescala según EXPORT_DPI (300dpi, antes 200) — todas
// las coordenadas base cambiaron, así que un manualOffset guardado con la
// versión anterior ya no calza; se bumpea la clave para partir limpio.
const STORAGE_KEY = "minicenefa_1promo_state_v5";

// DESIGN_DPI: la resolución a la que psd-tools extrajo las coordenadas de
// abajo (fuente única de verdad del layout, en px @200dpi). EXPORT_DPI: la
// resolución física final de exportación (PNG/PDF) — se puede subir sin
// tocar ningún número del layout, porque TEMPLATE se reescala por completo
// (canvas + cada capa) según SCALE = EXPORT_DPI / DESIGN_DPI antes de usarse
// en cualquier otro lado del archivo. Un solo número (EXPORT_DPI) controla
// todo el pipeline: tamaño de #stage, posiciones, el chunk pHYs del PNG y el
// tamaño físico en puntos del PDF — el tamaño de impresión real (12.4"x6.2")
// no cambia, solo la densidad de pixeles.
const DESIGN_DPI = 200;
const EXPORT_DPI = 300;
const SCALE = EXPORT_DPI / DESIGN_DPI;
function scalePx(n) { return Math.round(n * SCALE); }

// Coordenadas extraídas 1:1 del PSD de estructura a 200dpi (LEGAL ajustado a
// 936x75 según lo pedido). Esta es exactamente la forma que debería tener el
// feed cuando lo conectemos a Excel — reemplazar TEMPLATE_LAYERS_200DPI por
// datos por fila, no la lógica.
const TEMPLATE_LAYERS_200DPI = [
  { id: "fondo",       name: "FONDO",       type: "fondo",       x: 0,    y: 0,   w: 2480, h: 1239 },
  { id: "vector1",     name: "VECTOR1",     type: "vector",      x: 155,  y: 393, w: 906,  h: 792  },
  { id: "vector2",     name: "VECTOR2",     type: "vector",      x: 2104, y: 753, w: 464,  h: 464  },
  { id: "producto",    name: "PRODUCTO / BODEGÓN", type: "bodegon", x: 120,  y: 235, w: 1208, h: 834  },
  { id: "promo",       name: "PROMO (vertical)",   type: "promo",   x: 1363, y: 179, w: 761,  h: 946  },
  { id: "corporativo", name: "CORPORATIVO", type: "corporativo", x: 435,  y: 79,  w: 626,  h: 82   },
  { id: "legal",       name: "LEGAL",       type: "legal",       x: 1113, y: 80,  w: 936,  h: 75   },
  { id: "qr",          name: "QR",          type: "qr",          x: 2183, y: 179, w: 177,  h: 243  },
];

const TEMPLATE = {
  name: "LOC_MINICENEFA_CENEFAFASTFOODKOXKALOC1_FASTFOODKOXKA(1PROMO)",
  canvas: { width: scalePx(2480), height: scalePx(1240) },
  layers: TEMPLATE_LAYERS_200DPI.map(L => ({
    ...L, x: scalePx(L.x), y: scalePx(L.y), w: scalePx(L.w), h: scalePx(L.h),
  })),
};

// Separación mínima que se conserva entre el contenido real del bodegón y el
// de la promo cuando se "juntan" (ver computeComboLayout más abajo).
const COMBO_GAP = 35;

// SUAJE.svg inlinado tal cual (viewBox 2:1, idéntico al canvas 2480x1240,
// así que se estira sin distorsión — "centrado en el formato").
const SUAJE_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595.28 297.64">
  <path class="st0" d="M0,0v297.64h595.28V0H0ZM494.9,265.56c-50.27,8.67-91.37,17.64-201.42,18-103.54.33-159.47-13-220.65-25.75-61.18-12.75-58.68-64.57-58.68-64.57V14.07l566.89.02.03,182.52s5.54,53.14-86.17,68.95Z"/>
  <path class="st1" d="M494.9,265.56c-50.27,8.67-91.37,17.64-201.42,18-103.54.33-159.47-13-220.65-25.75-61.18-12.75-58.68-64.57-58.68-64.57V14.07l566.89.02.03,182.52s5.54,53.14-86.17,68.95Z"/>
</svg>`;

// La silueta de corte en sí (el mismo "d" del path st1 de arriba, la línea
// visible), aparte como constante reutilizable: se usa para recortar el PNG
// exportado con precisión vectorial (Path2D en canvas) en vez de depender de
// un fill blanco por encima.
const SUAJE_PATH_D = "M494.9,265.56c-50.27,8.67-91.37,17.64-201.42,18-103.54.33-159.47-13-220.65-25.75-61.18-12.75-58.68-64.57-58.68-64.57V14.07l566.89.02.03,182.52s5.54,53.14-86.17,68.95Z";
const SUAJE_VIEWBOX = { w: 595.28, h: 297.64 };

// ---------------------------------------------------------------------------
// Fuente Inter empaquetada en base64 (fonts.js, offline, ver comentario ahí)
// como @font-face de pantalla/PNG — y, más importante, como la MISMA fuente
// que opentype.js vectoriza glifo por glifo para el texto legal del PDF. Se
// usa un nombre propio ("InterCenefa") para no chocar con una fuente "Inter"
// que ya pueda estar instalada en el sistema con métricas distintas — así el
// PDF (trazado con opentype.js) y la pantalla/PNG (CSS) parten siempre del
// mismo archivo exacto.
// ---------------------------------------------------------------------------
const FONT_FACE_CSS = (() => {
  const F = (typeof FONT_ASSETS !== "undefined") ? FONT_ASSETS : {};
  return `
    @font-face{ font-family:"InterCenefa"; src:url(${F.Regular}) format("woff"); font-weight:400; font-style:normal; }
    @font-face{ font-family:"InterCenefa"; src:url(${F.Bold}) format("woff"); font-weight:700; font-style:normal; }
    @font-face{ font-family:"InterCenefa"; src:url(${F.Italic}) format("woff"); font-weight:400; font-style:italic; }
    @font-face{ font-family:"InterCenefa"; src:url(${F.BoldItalic}) format("woff"); font-weight:700; font-style:italic; }
  `;
})();

if (typeof document !== "undefined") {
  const fontStyleTag = document.createElement("style");
  fontStyleTag.id = "cenefa-fontface";
  fontStyleTag.textContent = FONT_FACE_CSS;
  document.head.appendChild(fontStyleTag);
}

// Parsers opentype.js de los 4 cortes (normal/bold/italic/bold-italic),
// creados una sola vez y cacheados — cada uno viene de la MISMA data: URI que
// ya carga el @font-face de arriba, decodificada a bytes vía atob() en vez de
// fetch() (que no funciona sobre file://, ver notas en exportStagePDF).
const INTER_FONTS = (() => {
  if (typeof FONT_ASSETS === "undefined" || typeof opentype === "undefined") return null;
  function parse(dataUrl) {
    const b64 = dataUrl.split(",")[1];
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return opentype.parse(bytes.buffer);
  }
  return {
    regular: parse(FONT_ASSETS.Regular),
    bold: parse(FONT_ASSETS.Bold),
    italic: parse(FONT_ASSETS.Italic),
    boldItalic: parse(FONT_ASSETS.BoldItalic),
  };
})();
function pickInterFont(bold, italic) {
  if (!INTER_FONTS) return null;
  if (bold && italic) return INTER_FONTS.boldItalic;
  if (bold) return INTER_FONTS.bold;
  if (italic) return INTER_FONTS.italic;
  return INTER_FONTS.regular;
}

// ---------------------------------------------------------------------------
// Estado
// ---------------------------------------------------------------------------

function defaultState() {
  const A = (typeof DEFAULT_ASSETS !== "undefined") ? DEFAULT_ASSETS : {};
  return {
    fondo: { hex: "#ffffff" },
    vector1: { dataUrl: A.vector1 || null },
    vector2: { dataUrl: A.vector2 || null },
    producto: { dataUrl: null, svgMarkup: null },
    promo: { dataUrl: null, svgMarkup: null },
    corporativo: {
      // OXXO no se sube manualmente: la plantilla elige sola entre sus 2
      // variantes según el fondo (ver pickCorpVariant). Si algún día hace
      // falta forzar un logo distinto, se setea aquí.
      oxxoOverride: null,
      selloOverride: null,
      // SELLO y EDAD18 arrancan invisibles (no se renderiza nada, "alpha 0")
      // hasta que se inyecte un archivo — así sea el mismo asset de siempre.
      // El dataUrl subido solo actúa como activador: la imagen que
      // efectivamente se muestra la sigue eligiendo pickCorpVariant() entre
      // sus 2 variantes automáticas según el fondo, no lo que se subió acá.
      sello: { dataUrl: null },
      edad18: { dataUrl: null },
    },
    // legal.html guarda el contenido del editor de texto enriquecido tal cual
    // innerHTML (negrita/cursiva/color aplicados por selección, no global).
    legal: { html: "", outline: true },
    qr: { dataUrl: A.qr || null }, // el código alfanumérico ya viene dentro del asset
    // Ajuste manual (drag) por capa: { [layerId]: {dx, dy, rot, scale} }, en
    // px del canvas real (dx/dy), grados (rot) y factor multiplicativo
    // (scale, 1 = tamaño original), SUMADO/COMPUESTO encima de la posición
    // que ya calculó el sistema automático (combos, agrupamiento OXXO+legal,
    // etc.) — no la reemplaza. Ver enableLayerDragging().
    manualOffset: {},
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    const merged = defaultState();
    for (const k of Object.keys(merged)) {
      if (parsed[k] && typeof parsed[k] === "object") Object.assign(merged[k], parsed[k]);
    }
    return merged;
  } catch (e) {
    console.warn("No se pudo leer el estado guardado, se usa el estado por defecto.", e);
    return defaultState();
  }
}

function saveState(state) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    console.warn("No se pudo guardar el estado (localStorage lleno o bloqueado).", e);
  }
}

function fileToDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function fileToText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsText(file);
  });
}

function isSvgFile(file) {
  return file.type === "image/svg+xml" || /\.svg$/i.test(file.name || "");
}

// ---------------------------------------------------------------------------
// OXXO / SELLO: selección automática de variante según el fondo
// ---------------------------------------------------------------------------

function hexToRgb(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec((hex || "").trim());
  if (!m) return { r: 255, g: 255, b: 255 };
  const n = parseInt(m[1], 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

// Heurística simple: "rojizo" = el canal rojo domina claramente sobre verde
// y azul, y no es un rojo tan oscuro que ya sea casi negro. Ajustable si en
// la práctica algún fondo se clasifica mal.
function isReddishBackground(hex) {
  const { r, g, b } = hexToRgb(hex);
  return r > 110 && r - g > 45 && r - b > 45;
}

function pickCorpVariant(state) {
  const A = (typeof DEFAULT_ASSETS !== "undefined") ? DEFAULT_ASSETS : {};
  const reddish = isReddishBackground(state.fondo && state.fondo.hex);
  const oxxo = (state.corporativo && state.corporativo.oxxoOverride) ||
    (reddish ? A.oxxoLetraBlanca : A.oxxoLetraRoja);
  // El sello arranca apagado: solo se muestra si se inyectó algo en su
  // slot (ver defaultState). Una vez activado, la imagen que se pinta la
  // sigue eligiendo esta misma función entre sus 2 variantes según el
  // fondo — lo inyectado es solo el interruptor, no reemplaza la imagen.
  const selloActive = !!(state.corporativo && state.corporativo.sello && state.corporativo.sello.dataUrl);
  const sello = selloActive
    ? ((state.corporativo && state.corporativo.selloOverride) || (reddish ? A.selloLetraBlanca : A.selloLetraRoja))
    : null;
  return { reddish, oxxo, sello };
}

// ---------------------------------------------------------------------------
// Regla general para BODEGÓN y PROMO: la imagen se adapta al ancho o al alto
// de su caja según cuál de los dos sea el límite ("object-fit: contain"),
// quedando siempre completa, centrada y usando el máximo espacio posible.
// ---------------------------------------------------------------------------

function imgContain(src, altLabel) {
  const img = document.createElement("img");
  img.src = src;
  if (altLabel) img.alt = altLabel;
  img.style.cssText = "width:100%;height:100%;object-fit:contain;display:block;";
  return img;
}

// ---------------------------------------------------------------------------
// SVG inyectado EN LÍNEA (no como <img>): un <img src="data:image/svg+xml,...">
// trata al SVG como una imagen opaca — el navegador no expone su DOM interno
// a la página, así que no hay forma de tocar sus colores desde acá. Insertando
// el propio markup del SVG directamente en el layer (innerHTML) en cambio, sí
// queda accesible: podemos buscar sus elementos y cambiarles el fill.
//
// Convención esperada dentro del SVG para que un color reaccione al fondo
// (misma heurística "rojizo" que ya usamos para elegir el logo/sello):
//   <path data-fill-normal="#e4032e" data-fill-reddish="#ffffff" d="..."/>
// Si el fondo elegido es rojizo, ese elemento pasa a data-fill-reddish; si
// no, se queda en data-fill-normal. Sin esos atributos, el SVG se muestra
// tal cual viene, sin tocar sus colores.
// ---------------------------------------------------------------------------

function isSvgAsset(asset) {
  return !!(asset && asset.svgMarkup);
}

// Espejo EXACTO de las condiciones que usa renderStage para decidir si una
// capa pinta algo propio (mismos "if" que gatillan cada appendChild ahí) —
// se usa SOLO para el pase raster compartido del PDF (buildRasterCropAssets):
// una caja vacía (sin asset propio) es un <div> transparente que no bloquea
// lo que haya debajo en el canvas aplanado, así que si dos cajas de la
// plantilla se superponen (p.ej. un vector decorativo de fondo asomándose
// bajo el placeholder de bodegón/promo antes de subir esa imagen), recortar
// la caja vacía "roba" por accidente los píxeles de la capa vecina que sí
// tiene contenido — normalmente invisible (queda pintado dos veces, exacto
// encima de sí mismo), pero se vuelve un artefacto visible en cuanto esa
// capa vecina tiene un giro/escala manual (la copia "fantasma" de acá no se
// mueve con ella). Filtrar ANTES del recorte compartido evita generar esa
// copia fantasma directamente, para cualquier capa realmente vacía.
function layerHasStateContent(state, id) {
  switch (id) {
    case "vector1": case "vector2":
      return !!(state[id] && state[id].dataUrl);
    case "producto": case "promo":
      return isSvgAsset(state[id]) || !!(state[id] && state[id].dataUrl);
    case "corporativo":
      return true; // el logo OXXO siempre está presente (ver pickCorpVariant)
    case "legal":
      return !!(state.legal && state.legal.html && state.legal.html.replace(/<[^>]+>/g, "").trim());
    case "qr":
      return !!(state.qr && state.qr.dataUrl);
    default:
      return true;
  }
}

function svgContain(markup, reddish) {
  const wrap = document.createElement("div");
  wrap.style.cssText = "width:100%;height:100%;";
  wrap.innerHTML = markup;
  const svg = wrap.querySelector("svg");
  if (!svg) return wrap;

  // Si el SVG no trae viewBox pero sí width/height numéricos, se lo
  // derivamos para que pueda escalar de forma responsiva dentro de su caja
  // (si no, un <svg width="600" height="800"> con medida fija no se encoge).
  if (!svg.getAttribute("viewBox")) {
    const wAttr = parseFloat(svg.getAttribute("width"));
    const hAttr = parseFloat(svg.getAttribute("height"));
    if (wAttr && hAttr) svg.setAttribute("viewBox", `0 0 ${wAttr} ${hAttr}`);
  }
  svg.style.width = "100%";
  svg.style.height = "100%";
  svg.style.display = "block";
  svg.removeAttribute("width");
  svg.removeAttribute("height");

  svg.querySelectorAll("[data-fill-normal]").forEach(el => {
    const normal = el.getAttribute("data-fill-normal");
    const alt = el.getAttribute("data-fill-reddish");
    el.style.fill = (reddish && alt) ? alt : normal;
  });

  return wrap;
}

// ---------------------------------------------------------------------------
// Auto-contraste GENÉRICO para el SVG de bodegón/promo: a diferencia de la
// convención data-fill-normal/-reddish (que requiere marcar el SVG a mano),
// esto funciona con cualquier SVG tal cual lo entrega diseño, sin tocarlo.
//
// Se comparan los colores que YA trae el SVG contra el hex del fondo elegido;
// si alguno se parece lo suficiente (misma familia de color, no necesita ser
// idéntico), se asume que ese color "se pierde" contra el fondo y se pasan
// TODAS las formas del SVG a blanco de una sola vez (igual que pediste: no es
// un ajuste forma por forma, es todo el gráfico a blanco cuando hay conflicto).
//
// Necesita que el <svg> ya esté conectado al documento (dentro de #stage) al
// momento de llamarse, porque los SVG de Illustrator casi siempre definen sus
// colores vía clases + <style> interno (class="st0", etc.) en vez de atributos
// fill="..." sueltos — y getComputedStyle solo resuelve esas clases una vez
// que el <style> del propio SVG ya forma parte del árbol vivo del documento.
// ---------------------------------------------------------------------------

const SVG_AUTOCONTRAST_TAGS = ["path", "rect", "circle", "ellipse", "polygon", "polyline", "text", "tspan", "line"];
const SVG_COLOR_MATCH_DISTANCE = 90; // distancia euclidiana en RGB (0-441); más alto = más permisivo para considerar "el mismo color"

function parseCssColor(str) {
  const m = /rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)/.exec(str || "");
  if (!m) return null;
  return { r: +m[1], g: +m[2], b: +m[3] };
}

function colorDistance(a, b) {
  const dr = a.r - b.r, dg = a.g - b.g, db = a.b - b.b;
  return Math.sqrt(dr * dr + dg * dg + db * db);
}

function applySvgAutoContrast(svgEl, fondoHex) {
  const fondoRgb = hexToRgb(fondoHex);
  const shapes = [];
  svgEl.querySelectorAll(SVG_AUTOCONTRAST_TAGS.join(",")).forEach(node => {
    const fill = getComputedStyle(node).fill;
    const rgb = parseCssColor(fill);
    if (rgb) shapes.push({ el: node, rgb });
  });

  const hasConflict = shapes.some(s => colorDistance(s.rgb, fondoRgb) < SVG_COLOR_MATCH_DISTANCE);
  if (!hasConflict) return;

  shapes.forEach(s => { s.el.style.fill = "#ffffff"; });
}

// Tamaño "natural" de un SVG leído de su viewBox/width/height — sin esto,
// el SVG no participa en la regla de "acercar" bodegón+promo (ver
// computeComboLayout), porque esa regla necesita saber su proporción real.
function getSvgNaturalSize(markup) {
  if (!markup) return null;
  const wrap = document.createElement("div");
  wrap.innerHTML = markup;
  const svg = wrap.querySelector("svg");
  if (!svg) return null;
  let w = parseFloat(svg.getAttribute("width"));
  let h = parseFloat(svg.getAttribute("height"));
  if ((!w || !h) && svg.getAttribute("viewBox")) {
    const parts = svg.getAttribute("viewBox").trim().split(/[\s,]+/).map(Number);
    if (parts.length === 4) { w = parts[2]; h = parts[3]; }
  }
  if (!w || !h) return null;
  return { w, h };
}

function getAssetNaturalSize(asset) {
  if (isSvgAsset(asset)) return Promise.resolve(getSvgNaturalSize(asset.svgMarkup));
  return getImageNaturalSize(asset && asset.dataUrl);
}

function getImageNaturalSize(dataUrl) {
  return new Promise((resolve) => {
    if (!dataUrl) { resolve(null); return; }
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve(null);
    img.src = dataUrl;
  });
}

// "Aprovechar el espacio": si el bodegón y/o la promo no llenan el ancho de
// su propia caja (porque su proporción es angosta dentro de una caja ancha),
// se acercan entre sí en vez de dejar un vacío grande — sin agrandar ninguna
// de las dos más allá de su caja máxima — y el par resultante queda centrado
// dentro del espacio total que ambas cajas ocupaban originalmente juntas.
async function computeComboLayout(state) {
  const bodegonBox = TEMPLATE.layers.find(l => l.id === "producto");
  const promoBox = TEMPLATE.layers.find(l => l.id === "promo");

  const [bNat, pNat] = await Promise.all([
    getAssetNaturalSize(state.producto),
    getAssetNaturalSize(state.promo),
  ]);

  function containedW(box, nat) {
    if (!nat || !nat.w || !nat.h) return box.w; // sin contenido: se queda con la caja completa
    const scale = Math.min(box.w / nat.w, box.h / nat.h);
    return Math.min(box.w, Math.round(nat.w * scale));
  }

  const cw1 = containedW(bodegonBox, bNat);
  const cw2 = containedW(promoBox, pNat);

  const spanLeft = bodegonBox.x;
  const spanRight = promoBox.x + promoBox.w;
  const combined = cw1 + COMBO_GAP + cw2;

  // Centrar el par respecto al CANVAS completo (no solo respecto al hueco
  // que dejan ambas cajas juntas), porque esa zona no es simétrica: el
  // margen izquierdo (hasta bodegonBox.x) es angosto mientras que a la
  // derecha de promoBox queda reservada la columna del QR + vector2, así
  // que centrar solo dentro de [spanLeft, spanRight] sesgaba el combo hacia
  // la izquierda del formato completo. Se acota para que nunca invada el
  // límite izquierdo original ni empuje el borde derecho más allá de donde
  // terminaba la caja de promo (para no encimarse con el QR/vector2).
  const idealLeft = TEMPLATE.canvas.width / 2 - combined / 2;
  const newBodegonX = Math.round(Math.max(spanLeft, Math.min(idealLeft, spanRight - combined)));
  const newPromoX = Math.round(newBodegonX + cw1 + COMBO_GAP);

  return {
    producto: { x: newBodegonX, y: bodegonBox.y, w: cw1, h: bodegonBox.h },
    promo: { x: newPromoX, y: promoBox.y, w: cw2, h: promoBox.h },
  };
}

// ---------------------------------------------------------------------------
// LEGAL: texto enriquecido (negrita/cursiva/color POR SELECCIÓN, no global) +
// autoajuste de tamaño + contorno "hacia afuera" (estilo Photoshop)
// ---------------------------------------------------------------------------

// El estilo del legal no debe depender SOLO de seleccionar texto a mano en
// el editor: también tiene que poder venir de otro lado (texto pegado desde
// afuera, y más adelante una celda de excel) — para eso reconoce una
// convención tipo markdown dentro del texto plano:
//   **negrita**   *cursiva*   {{#e4032e}}color{{/}}
// legalMarkdownToHtml() es el mismo punto de entrada que usará el
// importador de excel el día que se conecte: se le pasa el texto de la
// celda tal cual y devuelve el html listo para state.legal.html.
function escapeHtmlText(str) {
  const div = document.createElement("div");
  div.textContent = str || "";
  return div.innerHTML;
}

function legalMarkdownToHtml(text) {
  let out = escapeHtmlText(text);
  // color primero: {{#rrggbb}}texto{{/}} (3 o 6 dígitos hex)
  out = out.replace(/\{\{(#[0-9a-fA-F]{3,6})\}\}([\s\S]*?)\{\{\/\}\}/g,
    (_, color, inner) => `<span style="color:${color}">${inner}</span>`);
  // negrita antes que cursiva, para no comerse los ** al buscar *
  out = out.replace(/\*\*([\s\S]+?)\*\*/g, (_, inner) => `<span style="font-weight:bold">${inner}</span>`);
  out = out.replace(/\*([\s\S]+?)\*/g, (_, inner) => `<span style="font-style:italic">${inner}</span>`);
  out = out.replace(/\r\n|\r|\n/g, "<br>");
  return out;
}

// Punto de entrada sugerido para cuando el legal se alimente desde el excel
// más adelante: toma el texto tal cual viene en la celda (con la misma
// convención de arriba) y deja state.legal.html listo para renderStage().
function setLegalFromPlainText(state, text) {
  state.legal.html = legalMarkdownToHtml(text);
}

// El editor de legal es un contenteditable manejado con document.execCommand
// (bold/italic/foreColor sobre la selección activa) — con styleWithCSS
// activado (ver index.html) el navegador arma TODO como
// <span style="font-weight:bold/font-style:italic/color:...">, nunca <b>/<i>
// como etiqueta. Antes de pintarlo se limpia: de cada nodo solo sobrevive su
// negrita/cursiva/color EFECTIVOS (ver más abajo), reconstruidos como un
// <span style="..."> propio — cualquier otra cosa (tamaños de letra,
// fuentes, fondos, clases, scripts, atributos on*...) se descarta.
//
// La MISMA función también sanitiza el HTML que llega pegado desde afuera
// (Word, Google Docs, Excel, una página web) — ver el listener de "paste" en
// index.html, que ahora usa el clipboard con formato real (text/html) en vez
// de solo texto plano, así que el negrita/cursiva/color que ya traía la
// fuente original se reconoce solo, sin que el usuario tenga que
// reseleccionar y volver a aplicar el estilo a mano. Por eso, en vez de un
// allowlist de ETIQUETAS (un <td> de Excel con font-weight en su propio
// style, o el wrapper <b style="font-weight:normal"> que arma Google Docs
// como simple contenedor sin intención visual, no siguen la convención
// "la etiqueta ES el estilo" de acá), se calcula el estilo efectivo de CADA
// nodo combinando lo que implica su etiqueta (b/strong/i/em) con su propio
// style inline si lo trae — el style, cuando está presente, manda por
// encima de lo que diga la etiqueta (mismo criterio que ya usa
// parseLegalRuns: cualquier font-weight numérico ≥600/"bold"/"bolder" cuenta
// como negrita, "italic"/"oblique" como cursiva). La limpieza es recursiva
// (lo de adentro de cada nodo sale filtrado ANTES de decidir qué hacer con
// el nodo mismo), así un wrapper típico de Word/Docs varios niveles adentro
// no se cuela sin pasar por el filtro.
//
// Etiquetas de bloque comunes en HTML pegado desde afuera (párrafos de Word/
// Docs, filas de una tabla de Excel, títulos, listas...): al desenvolverse
// agregan un salto de línea en su lugar, para no pegar todo en un renglón
// corrido.
const LEGAL_BLOCK_BREAK_TAGS = new Set(["DIV", "P", "LI", "TR", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE"]);
// Estas se tiran ENTERAS, con todo y contenido (a diferencia de las demás,
// que se desenvuelven conservando su texto) — el HTML pegado desde un
// navegador/Word puede traer un <style>/<script> suelto dentro del propio
// fragmento copiado, y desenvolverlo normal insertaría ese CSS/JS como si
// fuera texto visible del legal.
const LEGAL_DISCARD_ENTIRELY_TAGS = new Set(["SCRIPT", "STYLE", "HEAD", "TITLE", "META", "LINK", "OBJECT", "IFRAME", "NOSCRIPT"]);

function sanitizeLegalHtml(html) {
  const wrap = document.createElement("div");
  wrap.innerHTML = html || "";

  (function clean(node) {
    Array.from(node.childNodes).forEach(child => {
      if (child.nodeType === Node.TEXT_NODE) {
        // Colapsa espacios/tabs/saltos de línea de puro FORMATO del HTML
        // fuente a un solo espacio (mismo criterio de colapso de whitespace
        // que ya aplica cualquier navegador al RENDERIZAR HTML) — sin esto,
        // un origen que pega markup "indentado" (saltos de línea entre
        // etiquetas que en pantalla no se notan porque el navegador los
        // colapsa) los cuela tal cual dentro del string de texto real que
        // arma buildPsdLegalTextLayer, y ahí SÍ se interpretan como saltos
        // de línea de verdad (Photoshop los guarda como "\r"), partiendo el
        // texto en muchas líneas de más.
        child.textContent = child.textContent.replace(/[ \t\r\n]+/g, " ");
        return;
      }
      if (child.nodeType !== Node.ELEMENT_NODE) { child.remove(); return; } // comentarios, etc.

      const tag = child.tagName;
      if (LEGAL_DISCARD_ENTIRELY_TAGS.has(tag)) { child.remove(); return; }
      if (tag === "BR") return; // ya viene bien tal cual

      // Se limpia POR DENTRO primero (recursivo) — así lo que termine
      // adentro del <span> de reemplazo (o desenvuelto suelto) ya salió
      // filtrado, en vez de colarse tal cual.
      clean(child);

      const cs = child.style || {};
      let bold = tag === "B" || tag === "STRONG";
      if (cs.fontWeight) bold = cs.fontWeight === "bold" || cs.fontWeight === "bolder" || parseInt(cs.fontWeight, 10) >= 600;
      let italic = tag === "I" || tag === "EM";
      if (cs.fontStyle) italic = cs.fontStyle === "italic" || cs.fontStyle === "oblique";
      const color = cs.color || null;

      const replacement = [];
      if (bold || italic || color) {
        const span = document.createElement("span");
        if (bold) span.style.fontWeight = "bold";
        if (italic) span.style.fontStyle = "italic";
        if (color) span.style.color = color;
        span.append(...child.childNodes);
        replacement.push(span);
      } else {
        replacement.push(...child.childNodes); // sin estilo propio que valga la pena: se desenvuelve
      }
      if (LEGAL_BLOCK_BREAK_TAGS.has(tag)) replacement.push(document.createElement("br"));
      child.replaceWith(...replacement);
    });
  })(wrap);

  // Recorta <br> colgantes al final (por ejemplo el último párrafo pegado
  // agrega uno de más al desenvolverse) para no sumar una línea en blanco —
  // saltando también espacios en blanco sueltos alrededor (típico de HTML
  // pegado desde Word/una página, con saltos de línea de puro formato entre
  // etiquetas que en HTML no pintan nada, pero que si no se saltan acá
  // esconden el <br> real y lo dejan sin recortar).
  while (wrap.lastChild) {
    const last = wrap.lastChild;
    if (last.nodeName === "BR") { wrap.removeChild(last); continue; }
    if (last.nodeType === Node.TEXT_NODE && !last.textContent.trim()) { wrap.removeChild(last); continue; }
    break;
  }

  return wrap.innerHTML;
}

// El ancho es fijo, así que solo hay que vigilar el alto tras el wrap.
function fitLegalText(el, boxW, boxH) {
  el.style.width = boxW + "px";
  let fontSize = Math.max(8, Math.floor(boxH * 0.5));
  el.style.fontSize = fontSize + "px";
  let guard = 0;
  while (fontSize > 8 && el.scrollHeight > boxH && guard < 300) {
    fontSize -= 1;
    el.style.fontSize = fontSize + "px";
    guard++;
  }
  return fontSize;
}

// -webkit-text-stroke dibuja el trazo CENTRADO sobre el borde de cada letra
// (mitad hacia adentro), y en texto bold eso cierra los contra-huecos de
// letras como "a"/"o" y se ve como manchas. En vez de eso, apilamos copias
// del texto desplazadas en 8 direcciones (el truco clásico de "outline hacia
// afuera" en CSS) — el glifo real queda intacto y solo se le suma un halo
// alrededor, que es justo el efecto "Position: Outside" de Photoshop.
function applyLegalOutline(el, fontSize, color) {
  const d = Math.max(0.5, fontSize * 0.045);
  const dirs = [
    [1, 0], [-1, 0], [0, 1], [0, -1],
    [0.72, 0.72], [-0.72, 0.72], [0.72, -0.72], [-0.72, -0.72],
  ];
  el.style.textShadow = dirs.map(([x, y]) => `${(x * d).toFixed(2)}px ${(y * d).toFixed(2)}px 0 ${color}`).join(", ");
}

function clearLegalOutline(el) {
  el.style.textShadow = "none";
}

// ---------------------------------------------------------------------------
// Grupo OXXO + LEGAL: cuando el combo corporativo se queda solo con el logo
// (sello y -18 apagados), ese logo y el bloque legal se acercan entre sí (sin
// agrandarse) y el par queda centrado en el espacio total que las dos cajas
// ocupaban juntas — misma idea que computeComboLayout() para bodegón+promo,
// solo que acá corre DESPUÉS de que el legal ya se autoajustó (necesita su
// ancho real ya medido, no se puede saber de antemano como con una imagen).
// ---------------------------------------------------------------------------
function getManualOffset(state, layerId) {
  const off = state.manualOffset && state.manualOffset[layerId];
  return { dx: (off && off.dx) || 0, dy: (off && off.dy) || 0, rot: (off && off.rot) || 0, scale: (off && off.scale) || 1 };
}

// Arma el string CSS transform de una capa a partir de su ajuste manual —
// mismo criterio en pantalla/PNG (acá) y en la rueda del mouse
// (enableLayerDragging): rotate() y scale() alrededor del centro de la caja
// (transform-origin por defecto, 50%/50%), y cada término se omite si está
// en su valor neutro para no ensuciar el atributo cuando no hace falta.
function manualTransformCss(off) {
  const parts = [];
  if (off.rot) parts.push(`rotate(${off.rot}deg)`);
  if (off.scale && off.scale !== 1) parts.push(`scale(${off.scale})`);
  return parts.join(" ");
}

function groupCorpWithLegal(state, corpLayerEl, legalLayerEl, legalTextEl, oxxoNat) {
  const corpBox = TEMPLATE.layers.find(l => l.id === "corporativo");
  const legalBox = TEMPLATE.layers.find(l => l.id === "legal");

  const cw1 = (oxxoNat && oxxoNat.w && oxxoNat.h)
    ? Math.min(corpBox.w, Math.round(oxxoNat.w * Math.min(corpBox.w / oxxoNat.w, corpBox.h / oxxoNat.h)))
    : corpBox.w;

  // Ancho real del texto legal ya ajustado: se mide sin wrap a su tamaño de
  // fuente actual — si le alcanza en una sola línea, ese es su ancho real;
  // si necesita más de una línea para caber en el alto, ya está usando todo
  // el ancho de su caja, así que no hay nada que acercar de ese lado.
  const prevWhiteSpace = legalTextEl.style.whiteSpace;
  legalTextEl.style.whiteSpace = "nowrap";
  const cw2 = Math.min(legalBox.w, legalTextEl.scrollWidth || legalBox.w);
  legalTextEl.style.whiteSpace = prevWhiteSpace || "";

  const gap = legalBox.x - (corpBox.x + corpBox.w); // mismo espacio que ya traía la plantilla entre ambas cajas
  const spanLeft = corpBox.x;
  const spanRight = legalBox.x + legalBox.w;
  const combined = cw1 + gap + cw2;

  const idealLeft = TEMPLATE.canvas.width / 2 - combined / 2;
  const newCorpX = Math.round(Math.max(spanLeft, Math.min(idealLeft, spanRight - combined)));
  const newLegalX = Math.round(newCorpX + cw1 + gap);

  const corpOff = getManualOffset(state, "corporativo");
  const legalOff = getManualOffset(state, "legal");
  corpLayerEl.dataset.baseX = newCorpX;
  corpLayerEl.dataset.baseY = corpBox.y;
  legalLayerEl.dataset.baseX = newLegalX;
  legalLayerEl.dataset.baseY = legalBox.y;
  corpLayerEl.style.left = (newCorpX + corpOff.dx) + "px";
  corpLayerEl.style.top = (corpBox.y + corpOff.dy) + "px";
  corpLayerEl.style.width = cw1 + "px";
  legalLayerEl.style.left = (newLegalX + legalOff.dx) + "px";
  legalLayerEl.style.top = (legalBox.y + legalOff.dy) + "px";
  legalLayerEl.style.width = cw2 + "px";
  // el texto interno se achica al mismo ancho que su contenedor: si cw2 es
  // su ancho natural en una sola línea, a ese ancho exacto sigue cabiendo
  // sin wrap; si cw2 quedó igual a legalBox.w (necesitaba varias líneas
  // para caber en el alto), el wrap sigue exactamente igual a como ya se
  // había ajustado.
  legalTextEl.style.width = cw2 + "px";
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

async function renderStage(stageEl, state, opts) {
  opts = opts || {};
  stageEl.innerHTML = "";
  stageEl.style.width = TEMPLATE.canvas.width + "px";
  stageEl.style.height = TEMPLATE.canvas.height + "px";

  const corp = pickCorpVariant(state);
  // Cuando el combo corporativo se queda solo con el logo OXXO (sello y -18
  // apagados), ese logo + el bloque legal se tratan como UN grupo y se
  // centran juntos en el formato — misma idea que bodegón+promo. Si hay
  // sello y/o -18 activos, cada caja se queda tal cual viene de la
  // plantilla (sin este acercamiento).
  const corpOxxoOnly = !corp.sello && !(state.corporativo.edad18 && state.corporativo.edad18.dataUrl);
  const [combo, oxxoNat] = await Promise.all([
    computeComboLayout(state),
    corpOxxoOnly ? getImageNaturalSize(corp.oxxo) : Promise.resolve(null),
  ]);

  let corpLayerEl = null, legalLayerEl = null, legalTextEl = null;

  for (const layer of TEMPLATE.layers) {
    const box = combo[layer.id] || layer; // producto/promo usan la caja ya "acercada"
    const el = document.createElement("div");
    el.className = "layer layer-" + layer.type;
    // El ajuste manual (drag) se SUMA encima de la posición ya calculada —
    // no la reemplaza. dataset.baseX/baseY guardan la posición "de antes del
    // arrastre" para que enableLayerDragging() pueda moverla en vivo sin
    // tener que rehacer todo el cálculo automático en cada frame.
    const off = getManualOffset(state, layer.id);
    el.style.left = (box.x + off.dx) + "px";
    el.style.top = (box.y + off.dy) + "px";
    el.style.width = box.w + "px";
    el.style.height = box.h + "px";
    // Giro/escala manual (drag-and-drop): transform-origin por defecto es
    // 50%/50%, o sea el centro de esta misma caja (box.w x box.h) — coincide
    // con el pivote que usa el PDF/PSD (ver matRotateScaleAround sobre
    // boxes[id] en exportStagePDF y rotatedLayerCropRect en el PSD), así
    // pantalla/PNG, PDF y PSD giran/escalan igual.
    el.style.transform = manualTransformCss(off);
    el.dataset.layerId = layer.id;
    el.dataset.baseX = box.x;
    el.dataset.baseY = box.y;

    if (layer.id === "fondo") {
      el.style.background = (state.fondo && state.fondo.hex) || "#ffffff";
    } else if (layer.id === "vector1" || layer.id === "vector2") {
      const s = state[layer.id];
      if (s && s.dataUrl) el.appendChild(imgContain(s.dataUrl, layer.name));
    } else if (layer.id === "producto") {
      if (isSvgAsset(state.producto)) el.appendChild(svgContain(state.producto.svgMarkup, corp.reddish));
      else if (state.producto && state.producto.dataUrl) el.appendChild(imgContain(state.producto.dataUrl, "Bodegón"));
    } else if (layer.id === "promo") {
      if (isSvgAsset(state.promo)) el.appendChild(svgContain(state.promo.svgMarkup, corp.reddish));
      else if (state.promo && state.promo.dataUrl) el.appendChild(imgContain(state.promo.dataUrl, "Promo"));
    } else if (layer.id === "corporativo") {
      corpLayerEl = el;
      const row = document.createElement("div");
      row.className = "corp-row";
      const items = [
        [corp.oxxo, "OXXO"],
        [corp.sello, "Responsabilidad"],
        [state.corporativo.edad18 && state.corporativo.edad18.dataUrl, "-18"],
      ].filter(([src]) => !!src);

      // Si sello y -18 están apagados, OXXO queda solo — en vez de dejarlo
      // pegado a la izquierda con todo ese espacio vacío a su derecha, se
      // manda al extremo derecho del contenedor. En cuanto se activa
      // cualquiera de los dos, vuelve al acomodo normal de izquierda a
      // derecha (OXXO, sello, -18 en ese orden).
      row.style.justifyContent = items.length === 1 ? "flex-end" : "flex-start";

      items.forEach(([src, label]) => {
        const cell = document.createElement("div");
        cell.className = "corp-item";
        const img = document.createElement("img");
        img.src = src;
        img.alt = label;
        img.style.cssText = "height:100%;width:auto;max-width:100%;object-fit:contain;display:block;";
        cell.appendChild(img);
        row.appendChild(cell);
      });
      el.appendChild(row);
    } else if (layer.id === "legal") {
      legalLayerEl = el;
      const p = document.createElement("div");
      p.className = "legal-text";
      p.innerHTML = sanitizeLegalHtml((state.legal && state.legal.html) || "");
      clearLegalOutline(p);
      el.appendChild(p);
      legalTextEl = p;
      requestAnimationFrame(() => {
        const finalSize = fitLegalText(p, box.w, box.h);
        if (state.legal && state.legal.outline) {
          applyLegalOutline(p, finalSize, "#ffffff");
        }
        if (corpOxxoOnly && corpLayerEl) {
          groupCorpWithLegal(state, corpLayerEl, legalLayerEl, p, oxxoNat);
        }
      });
    } else if (layer.id === "qr") {
      // El código alfanumérico ya viene horneado dentro del asset (no se
      // inyecta aparte), así que el QR es una sola imagen "contain" en toda
      // la caja — su ancho sigue definiendo el tamaño del módulo cuadrado.
      if (state.qr && state.qr.dataUrl) el.appendChild(imgContain(state.qr.dataUrl, "QR"));
    }

    if (opts.showLabels) {
      const tag = document.createElement("p");
      tag.className = "tag";
      tag.textContent = layer.name;
      el.appendChild(tag);
    }
    stageEl.appendChild(el);
  }

  // Auto-contraste del SVG de bodegón/promo: se hace DESPUÉS de que todo ya
  // quedó insertado en stageEl (que sí está en el documento), porque recién
  // ahí getComputedStyle puede resolver los colores que vienen de clases +
  // <style> interno del propio SVG (el caso típico de exports de Illustrator).
  for (const layerId of ["producto", "promo"]) {
    const asset = state[layerId];
    if (!isSvgAsset(asset)) continue;
    const container = stageEl.querySelector(`[data-layer-id="${layerId}"]`);
    const svgEl = container && container.querySelector("svg");
    if (svgEl) applySvgAutoContrast(svgEl, (state.fondo && state.fondo.hex) || "#ffffff");
  }

  if (opts.showSuaje !== false) {
    const wrap = document.createElement("div");
    wrap.id = "suaje-overlay";
    // El tamaño real vive en styles.css como fallback, pero se fija acá en
    // línea (igual que #stage arriba) para que quede correcto sin importar
    // el EXPORT_DPI/SCALE vigente, en vez de depender de un valor fijo en CSS.
    wrap.style.width = TEMPLATE.canvas.width + "px";
    wrap.style.height = TEMPLATE.canvas.height + "px";
    wrap.innerHTML = SUAJE_SVG;
    stageEl.appendChild(wrap);
  }

  return corp; // por si el llamador quiere mostrar qué variante quedó activa
}

// CSS transform:scale() no encoge la caja de layout de un elemento, así que
// un padre flex que intente centrar el elemento *visualmente* escalado no
// va a cuadrar a menos que algo más cargue el tamaño real (ya escalado).
// #stageScaler es ese algo: lo dimensionamos en px a exactamente
// scale*canvas, y el padre flex centra *eso* — #stage solo lo llena
// visualmente vía el transform.
function fitStageToContainer(stage, container, padding) {
  padding = padding || 0;
  const availW = container.clientWidth - padding * 2;
  const availH = container.clientHeight - padding * 2;
  const scale = Math.min(availW / TEMPLATE.canvas.width, availH / TEMPLATE.canvas.height, 1);
  stage.style.transform = "scale(" + scale + ")";
  const scaler = stage.parentElement;
  if (scaler && scaler.id === "stageScaler") {
    scaler.style.width = TEMPLATE.canvas.width * scale + "px";
    scaler.style.height = TEMPLATE.canvas.height * scale + "px";
  }
  return scale;
}

// ---------------------------------------------------------------------------
// Ajuste manual por arrastre ("último ajuste antes del PNG"): mueve una capa
// con el mouse y ese desplazamiento se guarda en state.manualOffset[id] — se
// SUMA encima de lo que ya calculó el sistema automático (combos, el
// agrupamiento OXXO+legal, etc.), nunca lo reemplaza. Si más adelante cambia
// el contenido de esa capa y el cálculo automático la reubica, el mismo
// desplazamiento se le vuelve a sumar desde la nueva posición — por eso hay
// un botón para reiniciar el ajuste manual capa por capa o de todas.
//
// Solo se activa cuando #stage tiene la clase "drag-mode" (la prende/apaga
// el checkbox del toolbar) — así no se interpone con el uso normal del
// inyector. La capa "fondo" nunca es arrastrable (no tiene sentido moverla).
// ---------------------------------------------------------------------------
function enableLayerDragging(stageEl, state, opts) {
  opts = opts || {};
  const exclude = new Set(opts.excludeIds || ["fondo"]);
  const onChange = opts.onChange || function () {};
  // margen que se permite arrastrar "de más" sin perder la capa por completo
  // fuera del canvas — se calcula como fracción del propio tamaño de la capa.
  const OVERDRAG_FRACTION = 0.5;

  let dragging = null;

  function currentScale() {
    const rect = stageEl.getBoundingClientRect();
    return rect.width ? rect.width / TEMPLATE.canvas.width : 1;
  }

  function onPointerDown(e) {
    if (!stageEl.classList.contains("drag-mode")) return;
    const layerEl = e.target.closest(".layer");
    if (!layerEl || !stageEl.contains(layerEl) || !layerEl.dataset.layerId) return;
    if (exclude.has(layerEl.dataset.layerId)) return;
    e.preventDefault();
    const id = layerEl.dataset.layerId;
    const off = getManualOffset(state, id);
    dragging = {
      id, el: layerEl,
      startClientX: e.clientX, startClientY: e.clientY,
      startDx: off.dx, startDy: off.dy,
      baseX: parseFloat(layerEl.dataset.baseX) || 0,
      baseY: parseFloat(layerEl.dataset.baseY) || 0,
      // offsetWidth/Height ya están en unidades del canvas real (2480x1240): el
      // transform:scale() de #stage para el preview afecta el render visual, no
      // el layout box, así que NO hay que dividir por currentScale() aquí.
      w: layerEl.offsetWidth || parseFloat(layerEl.style.width) || 0,
      h: layerEl.offsetHeight || parseFloat(layerEl.style.height) || 0,
      currentDx: off.dx, currentDy: off.dy,
    };
    layerEl.classList.add("dragging");
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  }

  function onPointerMove(e) {
    if (!dragging) return;
    const scale = currentScale() || 1;
    let dx = dragging.startDx + (e.clientX - dragging.startClientX) / scale;
    let dy = dragging.startDy + (e.clientY - dragging.startClientY) / scale;

    const marginX = dragging.w * OVERDRAG_FRACTION;
    const marginY = dragging.h * OVERDRAG_FRACTION;
    const minX = -dragging.baseX - marginX;
    const maxX = TEMPLATE.canvas.width - dragging.baseX - dragging.w + marginX;
    const minY = -dragging.baseY - marginY;
    const maxY = TEMPLATE.canvas.height - dragging.baseY - dragging.h + marginY;
    dx = Math.max(minX, Math.min(maxX, dx));
    dy = Math.max(minY, Math.min(maxY, dy));

    dragging.currentDx = dx;
    dragging.currentDy = dy;
    dragging.el.style.left = (dragging.baseX + dx) + "px";
    dragging.el.style.top = (dragging.baseY + dy) + "px";
  }

  function onPointerUp() {
    if (!dragging) return;
    dragging.el.classList.remove("dragging");
    if (!state.manualOffset) state.manualOffset = {};
    state.manualOffset[dragging.id] = { dx: Math.round(dragging.currentDx), dy: Math.round(dragging.currentDy) };
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
    const id = dragging.id;
    dragging = null;
    onChange(id);
  }

  stageEl.addEventListener("pointerdown", onPointerDown);

  // Giro y escala: rueda del mouse sobre una capa, en drag-mode.
  //   - Rueda sola = girar: 2°/notch normal, 15°/notch con Shift (grueso),
  //     0.5°/notch con Alt (fino).
  //   - Ctrl/Cmd + rueda = escalar (mismo gesto que el "zoom" de
  //     mapas/editores con trackpad o Ctrl+rueda — por eso el
  //     preventDefault: si no, el navegador lo toma como zoom de página):
  //     5%/notch normal, 20%/notch con Shift (grueso), 1%/notch con Alt
  //     (fino), acotado a [20%, 400%] para no perder la capa de vista ni
  //     reventar el recorte del PSD/PDF. Rueda hacia arriba/adelante
  //     (deltaY<0) agranda, igual que un zoom-in normal.
  // Se aplica el transform de inmediato (feedback en vivo) pero el
  // guardado+redibujado completo (onChange) se posterga un poco tras el
  // último evento de rueda, para no re-renderizar todo el stage en cada
  // notch mientras se gira/escala.
  let wheelSaveTimer = null;
  function onWheel(e) {
    if (!stageEl.classList.contains("drag-mode")) return;
    const layerEl = e.target.closest(".layer");
    if (!layerEl || !stageEl.contains(layerEl) || !layerEl.dataset.layerId) return;
    if (exclude.has(layerEl.dataset.layerId)) return;
    e.preventDefault();
    const id = layerEl.dataset.layerId;
    const off = getManualOffset(state, id);
    if (!state.manualOffset) state.manualOffset = {};
    if (!state.manualOffset[id]) state.manualOffset[id] = { dx: off.dx, dy: off.dy };

    if (e.ctrlKey || e.metaKey) {
      const step = e.shiftKey ? 0.20 : (e.altKey ? 0.01 : 0.05);
      const dir = e.deltaY > 0 ? -1 : 1;
      let scale = Math.round(off.scale * (1 + dir * step) * 1000) / 1000;
      scale = Math.max(0.2, Math.min(4, scale));
      state.manualOffset[id].scale = scale;
      layerEl.style.transform = manualTransformCss({ rot: off.rot, scale });
    } else {
      const step = e.shiftKey ? 15 : (e.altKey ? 0.5 : 2);
      const dir = e.deltaY > 0 ? 1 : -1;
      let rot = off.rot + dir * step;
      rot = ((rot % 360) + 360) % 360;
      if (rot > 180) rot -= 360; // rango (-180,180], más fácil de leer y de volver a 0
      state.manualOffset[id].rot = rot;
      layerEl.style.transform = manualTransformCss({ rot, scale: off.scale });
    }
    clearTimeout(wheelSaveTimer);
    wheelSaveTimer = setTimeout(() => onChange(id), 250);
  }
  stageEl.addEventListener("wheel", onWheel, { passive: false });
}

// ---------------------------------------------------------------------------
// Exportar PNG del resultado (sin tags ni cajas), a resolución real 2480x1240
// ---------------------------------------------------------------------------

// CSS mínimo necesario para que el clon se vea correcto DENTRO del
// foreignObject. No se reutiliza styles.css tal cual porque, sobre file://,
// una hoja enlazada con <link> dispara "Cannot access rules" al leer
// sheet.cssRules (Chrome la trata como si fuera de otro origen) y fetch()
// tampoco funciona sobre file:// — así que estas reglas viven acá, duplicadas
// a propósito, como la única fuente confiable dentro del SVG exportado.
const EXPORT_CSS = `
  ${FONT_FACE_CSS}
  *{ box-sizing:border-box; margin:0; padding:0; }
  .layer{ position:absolute; overflow:hidden; border:none; }
  .corp-row{ display:flex; align-items:center; gap:20px; width:100%; height:100%; padding:2px; }
  .corp-item{ height:100%; display:flex; align-items:center; }
  .legal-text{
    font-family: "InterCenefa", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    line-height:1.15; overflow:hidden; word-break:break-word; color:#000000;
  }
  .legal-text b, .legal-text strong{ font-weight:700; }
  .legal-text i, .legal-text em{ font-style:italic; }
  #suaje-overlay{ position:absolute; left:0; top:0; width:${TEMPLATE.canvas.width}px; height:${TEMPLATE.canvas.height}px; pointer-events:none; z-index:50; }
  #suaje-overlay svg{ width:100%; height:100%; display:block; }
  #suaje-overlay .st0{ fill:#ff2fb0; fill-opacity:.28; }
  #suaje-overlay .st1{ fill:none; stroke:#000; stroke-width:0.3px; }
`;

// ---------------------------------------------------------------------------
// Metadatos de resolución del PNG: canvas.toBlob() genera el PNG sin chunk
// pHYs, así que programas como Photoshop lo interpretan a 72dpi aunque el
// pixel-grid ya sea el correcto (12.4"x6.2" físicas al EXPORT_DPI vigente).
// Se le inyecta el chunk pHYs manualmente (sin librerías) para que la
// resolución quede marcada en EXPORT_DPI y el tamaño físico se reporte bien.
// ---------------------------------------------------------------------------

const PNG_CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    table[n] = c >>> 0;
  }
  return table;
})();

function png_crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    c = PNG_CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

async function setPngDpi(blob, dpi) {
  const buf = await blob.arrayBuffer();
  const bytes = new Uint8Array(buf);
  const dv = new DataView(buf);

  // El primer chunk tras la firma de 8 bytes SIEMPRE es IHDR: length(4) +
  // "IHDR"(4) + datos(13) + crc(4) = 25 bytes de datos -> 4+4+13+4 = 25... en
  // realidad length viene del propio campo, así que lo leemos en vez de
  // asumir 13 fijo (siempre es 13 en PNG válido, pero por robustez se lee).
  const ihdrDataLen = dv.getUint32(8, false);
  const insertAt = 8 + 4 + 4 + ihdrDataLen + 4; // fin del chunk IHDR completo

  const pixelsPerMeter = Math.round(dpi / 0.0254);
  const type = new Uint8Array([0x70, 0x48, 0x59, 0x73]); // "pHYs"
  const data = new Uint8Array(9);
  const dataView = new DataView(data.buffer);
  dataView.setUint32(0, pixelsPerMeter, false); // pixeles por metro, eje X
  dataView.setUint32(4, pixelsPerMeter, false); // pixeles por metro, eje Y
  data[8] = 1; // unidad: 1 = metro

  const crcInput = new Uint8Array(type.length + data.length);
  crcInput.set(type, 0);
  crcInput.set(data, type.length);
  const crc = png_crc32(crcInput);

  const chunk = new Uint8Array(4 + 4 + 9 + 4);
  const chunkView = new DataView(chunk.buffer);
  chunkView.setUint32(0, 9, false); // longitud de los datos del chunk
  chunk.set(type, 4);
  chunk.set(data, 8);
  chunkView.setUint32(17, crc, false);

  const out = new Uint8Array(bytes.length + chunk.length);
  out.set(bytes.subarray(0, insertAt), 0);
  out.set(chunk, insertAt);
  out.set(bytes.subarray(insertAt), insertAt + chunk.length);

  return new Blob([out], { type: "image/png" });
}

// (EXPORT_DPI vive arriba, junto a TEMPLATE/SCALE — un solo lugar controla
// tanto el tamaño del layout como el dpi marcado en PNG/PDF.)

// Arma el <canvas> final (según TEMPLATE.canvas, con el mismo recorte al
// troquel y la misma línea de corte que ya se usaba solo para el PNG) — PNG y
// PDF parten del mismo pixel-source exacto, así que lo compartido vive en una
// sola función en vez de duplicar el pipeline SVG→foreignObject→Image→canvas.
function buildExportCanvas(sourceStage, exportOpts) {
  exportOpts = exportOpts || {};
  const w = TEMPLATE.canvas.width;
  const h = TEMPLATE.canvas.height;

  const clone = sourceStage.cloneNode(true);
  clone.style.transform = "none";
  clone.style.boxShadow = "none";
  clone.style.position = "relative";
  clone.querySelectorAll(".tag").forEach(t => t.remove());
  clone.querySelectorAll(".layer").forEach(l => {
    l.style.border = "none";
    // Usado por el pase raster compartido del PDF: cada capa se rasteriza
    // SIN su giro/escala manual (si tiene) — ambos se re-aplican después,
    // vectorialmente, en la matriz de colocación del XObject (ver
    // exportStagePDF), para no perder recortes fuera de la caja original
    // cuando una capa girada o agrandada "sobresale" de su bounding box sin
    // girar/escalar. "transform: none" limpia rotate() Y scale() de un saque
    // (son el mismo atributo CSS).
    if (exportOpts.stripLayerTransforms) l.style.transform = "none";
  });
  // El troquel de referencia (línea + tinte magenta) se ve en pantalla solo
  // para QA. En el arte final no se rasteriza junto con el resto: se recorta
  // con un Path2D (misma silueta, precisión vectorial) y la línea de corte se
  // redibuja aparte directo en el canvas — así queda transparente TODO lo que
  // caiga por fuera del troquel (canal alpha real, no un fill blanco encima)
  // y la línea negra se ve nítida sobre eso.
  const suajeInClone = clone.querySelector("#suaje-overlay");
  if (suajeInClone && !exportOpts.keepSuajeTint) suajeInClone.remove();

  // Usado por el PDF vectorial: las capas que ya se van a dibujar como
  // vector (fondo, texto legal, arte SVG cuando aplica) se ocultan acá para
  // no rasterizarlas también en el pase raster compartido (ver
  // buildRasterCropCanvas más abajo) — position:absolute en el resto de
  // capas hace que ocultar una no mueva a las demás.
  (exportOpts.hideLayerIds || []).forEach(id => {
    const el = clone.querySelector(`[data-layer-id="${id}"]`);
    if (el) el.style.display = "none";
  });

  const cssText = EXPORT_CSS;
  const xhtml = new XMLSerializer().serializeToString(clone);
  const svgString = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">` +
    `<foreignObject width="100%" height="100%">` +
    `<div xmlns="http://www.w3.org/1999/xhtml" style="width:${w}px;height:${h}px;">` +
    `<style>${cssText}</style>${xhtml}</div></foreignObject></svg>`;

  // Un blob: URL para el SVG hace que Chrome marque el canvas como "tainted"
  // en cuanto el foreignObject trae <img> adentro, aunque esas imágenes sean
  // data: URIs — usando en cambio un data: URL para el SVG mismo, el canvas
  // queda limpio y toBlob()/toDataURL()/getImageData() funcionan.
  const svgDataUrl = "data:image/svg+xml;charset=utf-8;base64," + btoa(unescape(encodeURIComponent(svgString)));

  // scale: factor de sobremuestreo opcional (solo lo usa el PSD, para pedir
  // más píxeles por pulgada que el PNG/PDF en TODO el documento exportado,
  // sin tocar el DOM en vivo ni su layout — ver PSD_EXPORT_SCALE). 1 = sin
  // cambios respecto de antes (PNG/PDF siguen pidiendo scale por defecto).
  const scale = exportOpts.scale || 1;

  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = w * scale;
      canvas.height = h * scale;
      const ctx = canvas.getContext("2d");
      // Todo lo que sigue está escrito en coordenadas LÓGICAS (0..w, 0..h,
      // igual que antes) — este scale() de entrada es lo único que hace
      // falta para que se dibuje más grande sobre el canvas físico más
      // grande, sin tocar ninguna cuenta de más abajo.
      ctx.scale(scale, scale);

      const doClip = !exportOpts.keepSuajeTint && exportOpts.clipToSuaje !== false;
      const scaleX = w / SUAJE_VIEWBOX.w;
      const scaleY = h / SUAJE_VIEWBOX.h;

      if (doClip) {
        // El lienzo nace transparente (rgba 0,0,0,0). Se recorta a la
        // silueta del troquel ANTES de pintar el arte, así que ningún pixel
        // por fuera de esa forma llega a pintarse — queda alpha=0 real, no
        // un blanco encima. El clip queda "grabado" en espacio de canvas al
        // llamar a ctx.clip(), así que después se puede resetear la
        // transformación (preservando el `scale` de entrada — por eso NO es
        // un reset a la identidad pura) para dibujar la imagen ya compuesta
        // a su tamaño lógico 1:1.
        ctx.save();
        ctx.scale(scaleX, scaleY);
        ctx.clip(new Path2D(SUAJE_PATH_D));
        ctx.setTransform(scale, 0, 0, scale, 0, 0);
        ctx.drawImage(img, 0, 0, w, h);
        ctx.restore(); // quita el clip para poder dibujar la línea de corte sin que también quede recortada

        if (exportOpts.showSuajeLine !== false) {
          ctx.save();
          ctx.scale(scaleX, scaleY);
          ctx.lineWidth = 0.3; // mismo grosor que el troquel en pantalla (stroke-width:0.3 sobre el viewBox 595.28x297.64) — 10% del grosor original (3)
          ctx.strokeStyle = "#000000";
          ctx.stroke(new Path2D(SUAJE_PATH_D));
          ctx.restore();
        }
      } else {
        ctx.drawImage(img, 0, 0, w, h);
      }

      resolve(canvas);
    };
    img.onerror = (e) => reject(e);
    img.src = svgDataUrl;
  });
}

function triggerDownload(blob, filename) {
  const dlUrl = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = dlUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(dlUrl), 2000);
}

async function exportStagePNG(sourceStage, filename, exportOpts) {
  const canvas = await buildExportCanvas(sourceStage, exportOpts);
  return new Promise((resolve, reject) => {
    canvas.toBlob(async (blob) => {
      try {
        if (!blob) { reject(new Error("No se pudo generar el blob PNG.")); return; }
        const finalBlob = await setPngDpi(blob, EXPORT_DPI);
        triggerDownload(finalBlob, filename || "render.png");
        resolve();
      } catch (err) {
        reject(err);
      }
    }, "image/png");
  });
}

// ---------------------------------------------------------------------------
// Geometría vectorial compartida: parser de paths SVG ("d") + matrices 2D +
// conversión arco→bézier + emisor de operadores de trazo PDF. Es la pieza
// común que reutilizan tanto el arte SVG (promo/producto), la línea del
// troquel, como los contornos de cada glifo del texto legal (ver más abajo)
// — todos terminan siendo la misma lista normalizada de comandos
// {op:'M'|'L'|'C'|'Z', ...puntos}, sin importar de dónde vinieron.
// ---------------------------------------------------------------------------

const SVG_NUM_RE = /[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/;
const SVG_TOKEN_RE = new RegExp(`[MmLlHhVvCcSsQqTtAaZz]|${SVG_NUM_RE.source}`, "g");

function parseSvgPathD(d) {
  const tokens = (d || "").match(SVG_TOKEN_RE) || [];
  let i = 0;
  function num() { return parseFloat(tokens[i++]); }
  let cx = 0, cy = 0, sx = 0, sy = 0;
  let prevCmd = null;   // letra del último comando (para el "repetir implícito")
  let prevCtrl = null;  // último punto de control (para reflejar en S/T)
  const out = [];

  while (i < tokens.length) {
    let letter = tokens[i];
    if (/^[A-Za-z]$/.test(letter)) { i++; } else {
      letter = prevCmd; // número suelto: repite el comando anterior
      if (!letter) break;
    }
    const abs = letter === letter.toUpperCase();
    const cmd = letter.toUpperCase();

    if (cmd === "M") {
      let x = num(), y = num();
      if (!abs) { x += cx; y += cy; }
      out.push({ op: "M", x, y });
      cx = x; cy = y; sx = x; sy = y; prevCtrl = null;
      letter = abs ? "L" : "l"; // pares extra tras "M" son "L" implícitos
    } else if (cmd === "L") {
      let x = num(), y = num();
      if (!abs) { x += cx; y += cy; }
      out.push({ op: "L", x, y });
      cx = x; cy = y; prevCtrl = null;
    } else if (cmd === "H") {
      let x = num();
      if (!abs) x += cx;
      out.push({ op: "L", x, y: cy });
      cx = x; prevCtrl = null;
    } else if (cmd === "V") {
      let y = num();
      if (!abs) y += cy;
      out.push({ op: "L", x: cx, y });
      cy = y; prevCtrl = null;
    } else if (cmd === "C") {
      let x1 = num(), y1 = num(), x2 = num(), y2 = num(), x = num(), y = num();
      if (!abs) { x1 += cx; y1 += cy; x2 += cx; y2 += cy; x += cx; y += cy; }
      out.push({ op: "C", x1, y1, x2, y2, x, y });
      cx = x; cy = y; prevCtrl = { x: x2, y: y2 };
    } else if (cmd === "S") {
      let x2 = num(), y2 = num(), x = num(), y = num();
      if (!abs) { x2 += cx; y2 += cy; x += cx; y += cy; }
      const x1 = prevCtrl ? 2 * cx - prevCtrl.x : cx;
      const y1 = prevCtrl ? 2 * cy - prevCtrl.y : cy;
      out.push({ op: "C", x1, y1, x2, y2, x, y });
      cx = x; cy = y; prevCtrl = { x: x2, y: y2 };
    } else if (cmd === "Q") {
      let x1 = num(), y1 = num(), x = num(), y = num();
      if (!abs) { x1 += cx; y1 += cy; x += cx; y += cy; }
      // elevar cuadrática a cúbica (fórmula estándar)
      const c1x = cx + (2 / 3) * (x1 - cx), c1y = cy + (2 / 3) * (y1 - cy);
      const c2x = x + (2 / 3) * (x1 - x), c2y = y + (2 / 3) * (y1 - y);
      out.push({ op: "C", x1: c1x, y1: c1y, x2: c2x, y2: c2y, x, y });
      cx = x; cy = y; prevCtrl = { x: x1, y: y1 };
    } else if (cmd === "T") {
      let x = num(), y = num();
      if (!abs) { x += cx; y += cy; }
      const x1 = prevCtrl ? 2 * cx - prevCtrl.x : cx;
      const y1 = prevCtrl ? 2 * cy - prevCtrl.y : cy;
      const c1x = cx + (2 / 3) * (x1 - cx), c1y = cy + (2 / 3) * (y1 - cy);
      const c2x = x + (2 / 3) * (x1 - x), c2y = y + (2 / 3) * (y1 - y);
      out.push({ op: "C", x1: c1x, y1: c1y, x2: c2x, y2: c2y, x, y });
      cx = x; cy = y; prevCtrl = { x: x1, y: y1 };
    } else if (cmd === "A") {
      let rx = num(), ry = num(), rot = num(), laf = num(), sf = num(), x = num(), y = num();
      if (!abs) { x += cx; y += cy; }
      arcToBezierSegments(cx, cy, rx, ry, rot, laf !== 0, sf !== 0, x, y)
        .forEach(seg => out.push({ op: "C", ...seg }));
      cx = x; cy = y; prevCtrl = null;
    } else if (cmd === "Z") {
      out.push({ op: "Z" });
      cx = sx; cy = sy; prevCtrl = null;
    }
    prevCmd = letter;
  }
  return out;
}

// Conversión estándar de arco elíptico SVG a una serie de curvas cúbicas
// (parametrización de endpoint → centro, spec SVG 1.1 apéndice F.6, dividido
// en segmentos ≤90° con el factor k=4/3*tan(Δ/4) para cada uno).
function arcToBezierSegments(x1, y1, rx, ry, angleDeg, largeArc, sweep, x2, y2) {
  if (rx === 0 || ry === 0 || (x1 === x2 && y1 === y2)) {
    return [{ x1, y1, x2, y2, x: x2, y: y2 }];
  }
  rx = Math.abs(rx); ry = Math.abs(ry);
  const angle = (angleDeg * Math.PI) / 180;
  const cosA = Math.cos(angle), sinA = Math.sin(angle);

  const dx2 = (x1 - x2) / 2, dy2 = (y1 - y2) / 2;
  const x1p = cosA * dx2 + sinA * dy2;
  const y1p = -sinA * dx2 + cosA * dy2;

  let rxSq = rx * rx, rySq = ry * ry;
  const x1pSq = x1p * x1p, y1pSq = y1p * y1p;
  const radiiCheck = x1pSq / rxSq + y1pSq / rySq;
  if (radiiCheck > 1) {
    const s = Math.sqrt(radiiCheck);
    rx *= s; ry *= s; rxSq = rx * rx; rySq = ry * ry;
  }

  const sign = largeArc !== sweep ? 1 : -1;
  let sq = (rxSq * rySq - rxSq * y1pSq - rySq * x1pSq) / (rxSq * y1pSq + rySq * x1pSq);
  sq = sq < 0 ? 0 : sq;
  const coef = sign * Math.sqrt(sq);
  const cxp = coef * (rx * y1p / ry);
  const cyp = coef * -(ry * x1p / rx);

  const cx = cosA * cxp - sinA * cyp + (x1 + x2) / 2;
  const cy = sinA * cxp + cosA * cyp + (y1 + y2) / 2;

  function angleBetween(ux, uy, vx, vy) {
    const sgn = ux * vy - uy * vx < 0 ? -1 : 1;
    const len = Math.sqrt(ux * ux + uy * uy) * Math.sqrt(vx * vx + vy * vy);
    const dot = Math.min(1, Math.max(-1, (ux * vx + uy * vy) / len));
    return sgn * Math.acos(dot);
  }
  const ux = (x1p - cxp) / rx, uy = (y1p - cyp) / ry;
  const vx = (-x1p - cxp) / rx, vy = (-y1p - cyp) / ry;
  const angleStart = angleBetween(1, 0, ux, uy);
  let angleExtent = angleBetween(ux, uy, vx, vy);
  if (!sweep && angleExtent > 0) angleExtent -= 2 * Math.PI;
  if (sweep && angleExtent < 0) angleExtent += 2 * Math.PI;

  const numSegs = Math.max(1, Math.ceil(Math.abs(angleExtent) / (Math.PI / 2)));
  const delta = angleExtent / numSegs;
  const t = (4 / 3) * Math.tan(delta / 4);
  const segments = [];
  let a1 = angleStart;
  for (let i = 0; i < numSegs; i++) {
    const a2 = a1 + delta;
    const cos1 = Math.cos(a1), sin1 = Math.sin(a1), cos2 = Math.cos(a2), sin2 = Math.sin(a2);
    const q1x = cos1 - t * sin1, q1y = sin1 + t * cos1;
    const q2x = cos2 + t * sin2, q2y = sin2 - t * cos2;
    function mapPt(px, py) {
      const ex = rx * px, ey = ry * py;
      return { x: cx + (cosA * ex - sinA * ey), y: cy + (sinA * ex + cosA * ey) };
    }
    const Q1 = mapPt(q1x, q1y), Q2 = mapPt(q2x, q2y), P2 = mapPt(cos2, sin2);
    segments.push({ x1: Q1.x, y1: Q1.y, x2: Q2.x, y2: Q2.y, x: P2.x, y: P2.y });
    a1 = a2;
  }
  return segments;
}

function matIdentity() { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; }
function matTranslate(tx, ty) { return { a: 1, b: 0, c: 0, d: 1, e: tx, f: ty }; }
// m1 ∘ m2: aplica m2 primero y luego m1 (misma convención que DOMMatrix.multiply)
function matMultiply(m1, m2) {
  return {
    a: m1.a * m2.a + m1.c * m2.b, b: m1.b * m2.a + m1.d * m2.b,
    c: m1.a * m2.c + m1.c * m2.d, d: m1.b * m2.c + m1.d * m2.d,
    e: m1.a * m2.e + m1.c * m2.f + m1.e, f: m1.b * m2.e + m1.d * m2.f + m1.f,
  };
}
function matApplyPt(m, x, y) { return { x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f }; }
// Misma convención de signo que CSS transform:rotate(deg) en un espacio
// y-hacia-abajo (canvas-px): positivo = sentido horario visualmente.
function matRotate(deg) {
  const r = (deg * Math.PI) / 180, c = Math.cos(r), s = Math.sin(r);
  return { a: c, b: s, c: -s, d: c, e: 0, f: 0 };
}
function matRotateAround(cx, cy, deg) {
  if (!deg) return matIdentity();
  return matMultiply(matTranslate(cx, cy), matMultiply(matRotate(deg), matTranslate(-cx, -cy)));
}
function matScaleAround(cx, cy, sx, sy) {
  if (sx === 1 && sy === 1) return matIdentity();
  return matMultiply(matTranslate(cx, cy), matMultiply({ a: sx, b: 0, c: 0, d: sy, e: 0, f: 0 }, matTranslate(-cx, -cy)));
}
// Escala manual (drag-and-drop) y giro manual combinados, alrededor del
// mismo centro (el pivote de ambos ajustes) — se escala primero y el giro
// queda "afuera", aunque para el único caso que usa esto (escala uniforme,
// sx===sy) el orden entre ambos no cambia el resultado visual.
function matRotateScaleAround(cx, cy, deg, sx, sy) {
  return matMultiply(matRotateAround(cx, cy, deg || 0), matScaleAround(cx, cy, sx == null ? 1 : sx, sy == null ? 1 : sy));
}

function transformCommands(cmds, m) {
  return cmds.map(c => {
    if (c.op === "Z") return c;
    if (c.op === "C") {
      const p1 = matApplyPt(m, c.x1, c.y1), p2 = matApplyPt(m, c.x2, c.y2), p = matApplyPt(m, c.x, c.y);
      return { op: "C", x1: p1.x, y1: p1.y, x2: p2.x, y2: p2.y, x: p.x, y: p.y };
    }
    const p = matApplyPt(m, c.x, c.y);
    return { op: c.op, x: p.x, y: p.y };
  });
}

// Convierte el path de un glifo de opentype.js (comandos M/L/C/Q/Z, en el
// mismo espacio que se le pasó a getPath — ya en px de canvas, ver
// buildLegalGlyphRuns más abajo) a la misma lista normalizada de arriba. Los
// contornos TrueType (como Inter) son cuadráticos ("Q"), así que se elevan a
// cúbicos con la misma fórmula que ya usa el parser de SVG.
function glyphPathToCommands(path) {
  const out = [];
  let cx = 0, cy = 0;
  path.commands.forEach(c => {
    if (c.type === "M") { out.push({ op: "M", x: c.x, y: c.y }); cx = c.x; cy = c.y; }
    else if (c.type === "L") { out.push({ op: "L", x: c.x, y: c.y }); cx = c.x; cy = c.y; }
    else if (c.type === "C") { out.push({ op: "C", x1: c.x1, y1: c.y1, x2: c.x2, y2: c.y2, x: c.x, y: c.y }); cx = c.x; cy = c.y; }
    else if (c.type === "Q") {
      const x1 = cx + (2 / 3) * (c.x1 - cx), y1 = cy + (2 / 3) * (c.y1 - cy);
      const x2 = c.x + (2 / 3) * (c.x1 - c.x), y2 = c.y + (2 / 3) * (c.y1 - c.y);
      out.push({ op: "C", x1, y1, x2, y2, x: c.x, y: c.y });
      cx = c.x; cy = c.y;
    } else if (c.type === "Z") { out.push({ op: "Z" }); }
  });
  return out;
}

function pdfNum(n) {
  // 3 decimales alcanza de sobra para precisión de impresión y mantiene el
  // content stream razonablemente compacto.
  const r = Math.round(n * 1000) / 1000;
  return (Object.is(r, -0) ? 0 : r).toString();
}
function commandsToPdfOps(cmds) {
  let s = "";
  for (const c of cmds) {
    if (c.op === "M") s += `${pdfNum(c.x)} ${pdfNum(c.y)} m\n`;
    else if (c.op === "L") s += `${pdfNum(c.x)} ${pdfNum(c.y)} l\n`;
    else if (c.op === "C") s += `${pdfNum(c.x1)} ${pdfNum(c.y1)} ${pdfNum(c.x2)} ${pdfNum(c.y2)} ${pdfNum(c.x)} ${pdfNum(c.y)} c\n`;
    else if (c.op === "Z") s += "h\n";
  }
  return s;
}

// ---------------------------------------------------------------------------
// Texto legal vectorizado para el PDF: un layout propio (word-wrap +
// autoshrink), independiente del que hace el navegador para pantalla/PNG,
// medido con las métricas REALES de los glifos (opentype.js) en vez de
// reconstruir el layout de CSS. Nunca queda pixel-idéntico al HTML — usa su
// propio wrap — pero es el mismo texto, con el mismo bold/italic/color por
// selección, dentro de la misma caja, autoajustado igual de agresivo.
//
// Importante: SOLO se usan las APIs de opentype.js por-glifo (charToGlyph,
// glyph.getPath, glyph.advanceWidth) — las de más alto nivel que hacen
// "shaping" de strings completos (font.getPath(texto), font.getAdvanceWidth
// con kerning) tiran una excepción con esta fuente en particular (una
// característica GSUB de Inter que opentype.js todavía no soporta), así que
// ni siquiera se intenta usar kerning entre pares.
// ---------------------------------------------------------------------------

function anyTextColorToRgb(str) {
  const rgb = parseCssColor(str);
  if (rgb) return rgb;
  const m = /^#([0-9a-f]{6})$/i.exec((str || "").trim());
  if (m) {
    const n = parseInt(m[1], 16);
    return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
  }
  return { r: 0, g: 0, b: 0 };
}

function parseLegalRuns(html) {
  const div = document.createElement("div");
  div.innerHTML = html || "";
  const runs = [];
  function walk(node, bold, italic, color) {
    if (node.nodeType === 3) { // TEXT_NODE
      if (node.textContent) runs.push({ text: node.textContent, bold, italic, color });
      return;
    }
    if (node.nodeType !== 1) return; // no ELEMENT_NODE
    const tag = node.tagName;
    if (tag === "BR") { runs.push({ brk: true }); return; }
    let b = bold, it = italic, col = color;
    if (tag === "B" || tag === "STRONG") b = true;
    if (tag === "I" || tag === "EM") it = true;
    if (node.style && node.style.color) col = node.style.color;
    if (node.style && node.style.fontWeight) {
      const fw = node.style.fontWeight;
      if (fw === "bold" || parseInt(fw, 10) >= 600) b = true;
    }
    if (node.style && node.style.fontStyle === "italic") it = true;
    Array.prototype.forEach.call(node.childNodes, child => walk(child, b, it, col));
  }
  Array.prototype.forEach.call(div.childNodes, n => walk(n, false, false, "#000000"));
  return runs;
}

function tokenizeLegalRuns(runs) {
  const tokens = [];
  runs.forEach(r => {
    if (r.brk) { tokens.push({ type: "break" }); return; }
    r.text.split(/(\s+)/).forEach(part => {
      if (part === "") return;
      const base = { text: part, bold: r.bold, italic: r.italic, color: r.color };
      if (/^\s+$/.test(part)) tokens.push({ type: "space", ...base });
      else tokens.push({ type: "word", ...base });
    });
  });
  return tokens;
}

function measureGlyphRun(text, bold, italic, fontSizePx) {
  const font = pickInterFont(bold, italic);
  const scale = fontSizePx / font.unitsPerEm;
  let w = 0;
  for (const ch of text) w += font.charToGlyph(ch).advanceWidth * scale;
  return w;
}

// Wrap tipo "greedy": si una palabra sola no cabe ni en una línea vacía, se
// parte carácter por carácter (equivalente a word-break:break-word de CSS).
function layoutLegalTokens(tokens, boxW, fontSizePx) {
  const lines = [[]];
  let lineWidth = 0;
  function pushNewLine() { lines.push([]); lineWidth = 0; }

  tokens.forEach(tok => {
    if (tok.type === "break") { pushNewLine(); return; }
    if (tok.type === "space") {
      if (lines[lines.length - 1].length > 0) {
        lines[lines.length - 1].push(tok);
        lineWidth += measureGlyphRun(tok.text, tok.bold, tok.italic, fontSizePx);
      }
      return;
    }
    const w = measureGlyphRun(tok.text, tok.bold, tok.italic, fontSizePx);
    if (w > boxW) {
      if (lineWidth > 0) pushNewLine();
      let chunk = "", chunkW = 0;
      for (const ch of tok.text) {
        const cw = measureGlyphRun(ch, tok.bold, tok.italic, fontSizePx);
        if (chunkW + cw > boxW && chunk) {
          lines[lines.length - 1].push({ type: "word", text: chunk, bold: tok.bold, italic: tok.italic, color: tok.color });
          pushNewLine();
          chunk = ch; chunkW = cw;
        } else { chunk += ch; chunkW += cw; }
      }
      if (chunk) { lines[lines.length - 1].push({ type: "word", text: chunk, bold: tok.bold, italic: tok.italic, color: tok.color }); lineWidth = chunkW; }
      return;
    }
    if (lineWidth + w > boxW && lines[lines.length - 1].length > 0) pushNewLine();
    lines[lines.length - 1].push(tok);
    lineWidth += w;
  });

  while (lines.length > 1 && lines[lines.length - 1].length === 0) lines.pop();
  return lines;
}

// Mismo criterio de arranque/piso que fitLegalText (CSS, pantalla/PNG): parte
// de boxH*0.5 y baja de a 1px hasta que quepa, con piso en 8px.
function fitLegalTextVector(box, legalHtml) {
  const tokens = tokenizeLegalRuns(parseLegalRuns(legalHtml));
  const lineHeightFactor = 1.15;
  let fontSize = Math.max(8, Math.floor(box.h * 0.5));
  let lines = layoutLegalTokens(tokens, box.w, fontSize);
  while (fontSize > 8 && lines.length * fontSize * lineHeightFactor > box.h) {
    fontSize -= 1;
    lines = layoutLegalTokens(tokens, box.w, fontSize);
  }
  return { fontSize, lines, lineHeightFactor };
}

function legalLinesToGlyphItems(box, fontSize, lineHeightFactor, lines) {
  const items = [];
  const lineHeight = fontSize * lineHeightFactor;
  const ascent = (INTER_FONTS.regular.ascender / INTER_FONTS.regular.unitsPerEm) * fontSize;
  lines.forEach((line, li) => {
    let penX = box.x;
    const baselineY = box.y + li * lineHeight + ascent;
    line.forEach(tok => {
      const font = pickInterFont(tok.bold, tok.italic);
      const scale = fontSize / font.unitsPerEm;
      for (const ch of tok.text) {
        const glyph = font.charToGlyph(ch);
        if (tok.type === "word" && ch !== " ") {
          items.push({ glyph, x: penX, y: baselineY, fontSize, color: anyTextColorToRgb(tok.color) });
        }
        penX += glyph.advanceWidth * scale;
      }
    });
  });
  return items;
}

// Acumula todos los glifos del MISMO color en un solo path y un solo "f" —
// solo corta/emite cuando el color cambia (o al final), en vez de un
// fill-por-glifo, para no inflar el content stream de más.
function emitGlyphItemsColored(items, mat) {
  let ops = "", pending = "", lastKey = null, lastColor = null;
  function flush() {
    if (!pending) return;
    ops += `${(lastColor.r / 255).toFixed(3)} ${(lastColor.g / 255).toFixed(3)} ${(lastColor.b / 255).toFixed(3)} rg\n${pending}f\n`;
    pending = "";
  }
  items.forEach(it => {
    const key = it.color.r + "," + it.color.g + "," + it.color.b;
    if (lastKey !== null && key !== lastKey) flush();
    lastKey = key; lastColor = it.color;
    let cmds = glyphPathToCommands(it.glyph.getPath(it.x, it.y, it.fontSize));
    if (mat) cmds = transformCommands(cmds, mat);
    pending += commandsToPdfOps(cmds);
  });
  flush();
  return ops;
}

// Halo blanco detrás del texto (equivalente vectorial del text-shadow de
// applyLegalOutline en pantalla): TODOS los glifos en blanco, con relleno Y
// trazo grueso (operador "B"), dibujado ANTES del pase de color normal.
function emitGlyphItemsHalo(items, mat) {
  if (!items.length) return "";
  const avgFontSize = items[0].fontSize;
  const strokeW = Math.max(1, avgFontSize * 0.09); // ~2x el offset de applyLegalOutline, para cubrir el halo completo
  let pending = "";
  items.forEach(it => {
    let cmds = glyphPathToCommands(it.glyph.getPath(it.x, it.y, it.fontSize));
    if (mat) cmds = transformCommands(cmds, mat);
    pending += commandsToPdfOps(cmds);
  });
  return `1 1 1 RG\n1 1 1 rg\n${pdfNum(strokeW)} w\n1 J\n1 j\n${pending}B\n`;
}

// rotDeg/scale (opcionales): ajuste manual (drag-and-drop) alrededor del
// centro de box — el layout/salto de línea se calcula SIEMPRE con box.w/h
// sin escalar (fitLegalTextVector), y recién el bloque ya armado se gira y
// escala como un todo, igual que el desborde a propósito del PSD
// (LEGAL_TEXT_PSD_SCALE) y que el transform:scale() de la vista en pantalla.
function buildLegalContentOps(box, legalHtml, outlineOn, rotDeg, scale) {
  if (!legalHtml || !legalHtml.replace(/<[^>]+>/g, "").trim()) return "";
  const { fontSize, lines, lineHeightFactor } = fitLegalTextVector(box, legalHtml);
  const items = legalLinesToGlyphItems(box, fontSize, lineHeightFactor, lines);
  const hasTransform = !!rotDeg || (scale != null && scale !== 1);
  const mat = hasTransform ? matRotateScaleAround(box.x + box.w / 2, box.y + box.h / 2, rotDeg || 0, scale == null ? 1 : scale, scale == null ? 1 : scale) : null;
  let ops = "";
  if (outlineOn) ops += emitGlyphItemsHalo(items, mat);
  ops += emitGlyphItemsColored(items, mat);
  return ops;
}

// ---------------------------------------------------------------------------
// Exportar PDF de impresión (1 página, mismo pixel-source que el PNG) — se
// arma el binario a mano, sin librerías, igual que el chunk pHYs del PNG.
// La imagen se comprime con CompressionStream("deflate"), que produce un
// stream zlib estándar (cabecera 0x78 0x9C) — exactamente lo que el filtro
// /FlateDecode de PDF espera, así que no hace falta reimplementar deflate.
// El canal alfa (para que el recorte del troquel siga siendo transparente
// también en el PDF) va como imagen /SMask aparte, en escala de grises.
// ---------------------------------------------------------------------------

async function deflateBytes(bytes) {
  const cs = new CompressionStream("deflate");
  const writer = cs.writable.getWriter();
  writer.write(bytes);
  writer.close();
  const buf = await new Response(cs.readable).arrayBuffer();
  return new Uint8Array(buf);
}

// ---------------------------------------------------------------------------
// Arte SVG (promo/producto) vectorizado para el PDF: solo se intenta cuando
// el SVG completo está compuesto ÚNICAMENTE de <path>/<rect> con relleno de
// color sólido (sin gradientes/patrones, sin <clipPath>, sin <use>/<text>/
// formas que no sean path o rect) — cualquier otra cosa dispara el fallback
// a raster para ESA capa completa, en vez de arriesgarse a perder contenido
// silenciosamente.
// ---------------------------------------------------------------------------

// lineargradient/radialgradient/stop viven dentro de <defs> como DEFINICIONES,
// no como formas a dibujar — se ignoran en el recorrido de svgCanFullyVectorize
// (su contenido se lee aparte, vía resolvePaint/buildGradientDescriptor, solo
// para los path/rect que realmente los referencian con fill/stroke:url(#...)).
const PDF_STRUCTURAL_SVG_TAGS = new Set(["svg", "g", "defs", "title", "desc", "metadata", "style", "lineargradient", "radialgradient", "stop"]);
const PDF_VECTORIZABLE_SVG_TAGS = new Set(["path", "rect"]);

// ---------------------------------------------------------------------------
// Gradientes SVG (linearGradient/radialGradient) -> PDF Shading Pattern.
// Un fill/stroke resuelve a 'url("#id")' vía getComputedStyle cuando referencia
// un gradiente. Se soporta spreadMethod="pad" (por defecto) con stops de
// opacidad 1 (una opacidad variable por stop necesitaría un soft-mask de
// luminosidad -> fuera de alcance por ahora, cae a raster igual que un patrón).
// ---------------------------------------------------------------------------
function parseGradientRef(cssColorStr) {
  if (!cssColorStr) return null;
  const m = /^url\(\s*["']?#([^"')\s]+)["']?\s*\)/.exec(cssColorStr.trim());
  return m ? m[1] : null;
}

function findGradientEl(svgEl, id) {
  try { return svgEl.querySelector(`[id="${CSS.escape(id)}"]`); } catch (e) { return null; }
}

// Los stops pueden vivir en un gradiente "base" compartido vía href/xlink:href
// (patrón común de exports de Illustrator: varios gradientes con su propia
// geometría pero apuntando a los mismos <stop>) — se sigue esa cadena SOLO
// para encontrar los stops, la geometría propia (x1/y1/.../gradientTransform)
// siempre se lee del gradiente referenciado directamente.
function getGradientStops(svgEl, gradEl) {
  let el = gradEl, seen = new Set(), stopEls = [];
  while (el && !seen.has(el)) {
    seen.add(el);
    stopEls = Array.from(el.querySelectorAll(":scope > stop"));
    if (stopEls.length) break;
    const href = el.getAttribute("href") || el.getAttributeNS("http://www.w3.org/1999/xlink", "href");
    if (!href || !href.startsWith("#")) break;
    el = findGradientEl(svgEl, href.slice(1));
  }
  if (!stopEls.length) return null;
  let prevOffset = 0;
  return stopEls.map(s => {
    const cs = getComputedStyle(s);
    const offsetAttr = (s.getAttribute("offset") || "0").trim();
    let offset = offsetAttr.endsWith("%") ? parseFloat(offsetAttr) / 100 : parseFloat(offsetAttr);
    if (!isFinite(offset)) offset = 0;
    offset = Math.max(0, Math.min(1, offset));
    if (offset < prevOffset) offset = prevOffset; // fuerza orden no decreciente (igual que exige el spec SVG)
    prevOffset = offset;
    const colorStr = (cs.stopColor && cs.stopColor !== "") ? cs.stopColor : (s.getAttribute("stop-color") || "#000000");
    const rgb = parseCssColor(colorStr) || hexToRgb(colorStr) || { r: 0, g: 0, b: 0 };
    let opacity = parseFloat(cs.stopOpacity);
    if (!isFinite(opacity)) opacity = parseFloat(s.getAttribute("stop-opacity"));
    if (!isFinite(opacity)) opacity = 1;
    return { offset, rgb, opacity };
  });
}

function buildGradientDescriptor(svgEl, id) {
  const gradEl = findGradientEl(svgEl, id);
  if (!gradEl) return null;
  const tag = gradEl.tagName.toLowerCase();
  if (tag !== "lineargradient" && tag !== "radialgradient") return null;
  const stops = getGradientStops(svgEl, gradEl);
  if (!stops || !stops.length) return null;
  if (stops.some(s => s.opacity < 0.999)) return null; // opacidad variable: no soportado aún
  const spread = gradEl.getAttribute("spreadMethod") || "pad";
  if (spread !== "pad") return null; // reflect/repeat no representables con un shading simple

  let gt = matIdentity();
  try {
    if (gradEl.gradientTransform && gradEl.gradientTransform.baseVal.numberOfItems) {
      gt = domMatrixToPlain(gradEl.gradientTransform.baseVal.consolidate().matrix);
    }
  } catch (e) { /* sin gradientTransform válido: identidad */ }

  const units = gradEl.getAttribute("gradientUnits") || "objectBoundingBox";
  function num(attr, def) {
    const v = gradEl.getAttribute(attr);
    if (v == null || v.trim() === "") return def;
    return v.trim().endsWith("%") ? parseFloat(v) / 100 : parseFloat(v);
  }

  if (tag === "lineargradient") {
    return {
      kind: "linear", units, gradientTransform: gt, stops,
      coords: { x1: num("x1", 0), y1: num("y1", 0), x2: num("x2", 1), y2: num("y2", 0) },
    };
  }
  const cx = num("cx", 0.5), cy = num("cy", 0.5), r = num("r", 0.5);
  if (!(r > 0)) return null;
  return {
    kind: "radial", units, gradientTransform: gt, stops,
    coords: { cx, cy, r, fx: num("fx", cx), fy: num("fy", cy) },
  };
}

// Resuelve un fill/stroke computado a { type:"solid", rgb } o
// { type:"gradient", gradient } — null si es "none", un patrón, o un
// gradiente con alguna característica no soportada (ver arriba).
function resolvePaint(cssColorStr, svgEl) {
  if (!cssColorStr || cssColorStr === "none") return null;
  const solid = parseCssColor(cssColorStr);
  if (solid) return { type: "solid", rgb: solid };
  const gradId = parseGradientRef(cssColorStr);
  if (gradId) {
    const gradient = buildGradientDescriptor(svgEl, gradId);
    if (gradient) return { type: "gradient", gradient };
  }
  return null;
}

function svgCanFullyVectorize(svgEl) {
  const all = svgEl.querySelectorAll("*");
  for (const el of all) {
    const tag = el.tagName.toLowerCase();
    if (PDF_STRUCTURAL_SVG_TAGS.has(tag)) continue;
    if (!PDF_VECTORIZABLE_SVG_TAGS.has(tag)) return false; // use/text/circle/clipPath/etc.
    const cs = getComputedStyle(el);
    if (cs.clipPath && cs.clipPath !== "none") return false;
    const hasFill = cs.fill && cs.fill !== "none";
    const hasStroke = cs.stroke && cs.stroke !== "none";
    if (hasFill && !resolvePaint(cs.fill, svgEl)) return false;     // relleno sin resolver (patrón, gradiente no soportado)
    if (hasStroke && !resolvePaint(cs.stroke, svgEl)) return false;
    if (!hasFill && !hasStroke) continue; // forma invisible, no aporta ni descalifica
  }
  return true;
}

function rectElementToPathD(el) {
  const x = parseFloat(el.getAttribute("x")) || 0;
  const y = parseFloat(el.getAttribute("y")) || 0;
  const rw = parseFloat(el.getAttribute("width")) || 0;
  const rh = parseFloat(el.getAttribute("height")) || 0;
  return `M${x},${y} L${x + rw},${y} L${x + rw},${y + rh} L${x},${y + rh} Z`;
}

function domMatrixToPlain(m) { return { a: m.a, b: m.b, c: m.c, d: m.d, e: m.e, f: m.f }; }

function rgbTriplet(rgb) { return `${(rgb.r / 255).toFixed(3)} ${(rgb.g / 255).toFixed(3)} ${(rgb.b / 255).toFixed(3)}`; }

// boxX/boxY: posición (canvas-px) de la capa que contiene el SVG — getCTM()
// ya resuelve viewBox + cualquier <g transform> anidado hasta el propio
// <svg>, así que solo falta sumarle dónde está ese <svg> dentro del canvas.
// extraMat (opcional): matriz adicional aplicada DESPUÉS de lo anterior — se
// usa para el giro manual (drag-and-drop) alrededor del centro de la caja.
function buildSvgVectorShapes(svgEl, boxX, boxY, extraMat) {
  extraMat = extraMat || matIdentity();
  const shapes = [];
  svgEl.querySelectorAll("path,rect").forEach(el => {
    const tag = el.tagName.toLowerCase();
    const dStr = tag === "rect" ? rectElementToPathD(el) : (el.getAttribute("d") || "");
    if (!dStr) return;
    const cs = getComputedStyle(el);
    if (parseFloat(cs.opacity) === 0) return;
    let mode = "fill", paint = resolvePaint(cs.fill, svgEl);
    if (!paint) {
      paint = resolvePaint(cs.stroke, svgEl);
      if (!paint) return; // sin relleno ni trazo resoluble: nada que dibujar
      mode = "stroke";
    }
    let ctm = null;
    try { ctm = el.getCTM(); } catch (e) { /* elemento no renderizado */ }
    if (!ctm) return;
    const m = matMultiply(extraMat, matMultiply(matTranslate(boxX, boxY), domMatrixToPlain(ctm)));
    const cmds = transformCommands(parseSvgPathD(dStr), m);
    let bbox = null;
    if (paint.type === "gradient" && paint.gradient.units !== "userSpaceOnUse") {
      // objectBoundingBox: el gradiente necesita el bbox LOCAL del propio
      // elemento (mismo espacio que su "d", antes de su propio CTM).
      try { bbox = el.getBBox(); } catch (e) { bbox = null; }
      if (!bbox || !bbox.width || !bbox.height) return;
    }
    shapes.push({
      cmds, paint, mode, shapeMatrix: m, bbox,
      strokeWidth: mode === "stroke" ? (parseFloat(cs.strokeWidth) || 1) : 0,
    });
  });
  return shapes;
}

// Matriz /Matrix de un PDF Shading Pattern: mapea "espacio del gradiente"
// (unidades del propio linearGradient/radialGradient) al espacio POR DEFECTO
// de la página (puntos PDF) — un Pattern ignora el "cm" activo del content
// stream, así que hay que hornear TODA la cadena acá, no solo hasta canvas-px.
function buildGradientPatternMatrix(shapeMatrix, bbox, gradient, cmPxToPtMat) {
  let unitMat = matIdentity();
  if (gradient.units !== "userSpaceOnUse" && bbox) {
    // objectBoundingBox: [0,1]x[0,1] -> bbox local del elemento.
    unitMat = matMultiply(matTranslate(bbox.x, bbox.y), { a: bbox.width, b: 0, c: 0, d: bbox.height, e: 0, f: 0 });
  }
  const chain = matMultiply(shapeMatrix, matMultiply(unitMat, gradient.gradientTransform));
  return matMultiply(cmPxToPtMat, chain);
}

// registerGradientPattern(gradient, patternMatrix) -> nombre de recurso (p.ej. "Pg0"),
// provisto por exportStagePDF (acumula los patterns a serializar más abajo).
function emitVectorShapesOps(shapes, cmPxToPtMat, registerGradientPattern) {
  let ops = "";
  shapes.forEach(s => {
    const pathOps = commandsToPdfOps(s.cmds);
    if (s.paint.type === "gradient") {
      const matrix = buildGradientPatternMatrix(s.shapeMatrix, s.bbox, s.paint.gradient, cmPxToPtMat);
      const name = registerGradientPattern(s.paint.gradient, matrix);
      if (s.mode === "fill") {
        ops += `/Pattern cs\n/${name} scn\n${pathOps}f\n`;
      } else {
        ops += `/Pattern CS\n/${name} SCN\n${pdfNum(s.strokeWidth)} w\n${pathOps}S\n`;
      }
    } else {
      const rgbStr = rgbTriplet(s.paint.rgb);
      if (s.mode === "fill") ops += `${rgbStr} rg\n${pathOps}f\n`;
      else ops += `${rgbStr} RG\n${pdfNum(s.strokeWidth)} w\n${pathOps}S\n`;
    }
  });
  return ops;
}

function emitFondoRectOps(box, hex) {
  const rgb = hexToRgb(hex);
  const rgbStr = `${(rgb.r / 255).toFixed(3)} ${(rgb.g / 255).toFixed(3)} ${(rgb.b / 255).toFixed(3)}`;
  return `${rgbStr} rg\n${pdfNum(box.x)} ${pdfNum(box.y)} ${pdfNum(box.w)} ${pdfNum(box.h)} re\nf\n`;
}

// Pase raster COMPARTIDO para todo lo que no se vectorizó (fotos, logos, QR,
// vectores decorativos no-SVG) — se rasteriza el stage completo UNA vez
// (con las capas ya vectorizadas ocultas para no duplicarlas, y SIN recorte
// al troquel: eso lo hace el clip vectorial global del PDF por igual para
// vector y raster) y de ahí se recorta cada capa a su propia cajita.
async function buildRasterCropAssets(sourceStage, rasterLayerIds, hideLayerIds, boxes) {
  if (!rasterLayerIds.length) return [];
  const canvas = await buildExportCanvas(sourceStage, { clipToSuaje: false, keepSuajeTint: false, hideLayerIds, stripLayerTransforms: true });
  const ctx = canvas.getContext("2d");
  const assets = [];
  for (const id of rasterLayerIds) {
    const box = boxes[id];
    if (!box) continue;
    const bx = Math.max(0, Math.round(box.x)), by = Math.max(0, Math.round(box.y));
    const bw = Math.min(canvas.width - bx, Math.round(box.w));
    const bh = Math.min(canvas.height - by, Math.round(box.h));
    if (bw <= 0 || bh <= 0) continue;
    const pixels = ctx.getImageData(bx, by, bw, bh).data;
    let hasContent = false;
    for (let k = 3; k < pixels.length; k += 4) { if (pixels[k] !== 0) { hasContent = true; break; } }
    if (!hasContent) continue; // nada visible en esa caja: no gastar un XObject en ella

    const rgb = new Uint8Array(bw * bh * 3);
    const alpha = new Uint8Array(bw * bh);
    for (let i = 0, j = 0, k = 0; i < pixels.length; i += 4, j += 3, k++) {
      rgb[j] = pixels[i]; rgb[j + 1] = pixels[i + 1]; rgb[j + 2] = pixels[i + 2]; alpha[k] = pixels[i + 3];
    }
    const [rgbZ, alphaZ] = await Promise.all([deflateBytes(rgb), deflateBytes(alpha)]);
    assets.push({ id, x: bx, y: by, w: bw, h: bh, rgbZ, alphaZ });
  }
  return assets;
}

// Función PDF de un gradiente: Type 2 (exponencial) para 2 stops, o Type 3
// (stitching de N-1 funciones Type 2, una por tramo entre stops consecutivos)
// para 3+. allocObj() entrega el próximo número de objeto libre (comparte el
// mismo contador que las imágenes raster, ver exportStagePDF).
function buildPdfGradientFunctionObjs(stops, allocObj) {
  const bodies = [];
  if (stops.length <= 2) {
    const n = allocObj();
    const c0 = stops[0].rgb, c1 = stops[stops.length - 1].rgb;
    bodies.push([n, `<< /FunctionType 2 /Domain [0 1] /C0 [${rgbTriplet(c0)}] /C1 [${rgbTriplet(c1)}] /N 1 >>`]);
    return { mainObjNum: n, bodies };
  }
  const subNums = [];
  for (let i = 0; i < stops.length - 1; i++) {
    const n = allocObj();
    subNums.push(n);
    bodies.push([n, `<< /FunctionType 2 /Domain [0 1] /C0 [${rgbTriplet(stops[i].rgb)}] /C1 [${rgbTriplet(stops[i + 1].rgb)}] /N 1 >>`]);
  }
  const bounds = stops.slice(1, -1).map(s => pdfNum(s.offset)).join(" ");
  const encode = subNums.map(() => "0 1").join(" ");
  const fnRefs = subNums.map(n => `${n} 0 R`).join(" ");
  const stitchNum = allocObj();
  bodies.push([stitchNum, `<< /FunctionType 3 /Domain [0 1] /Functions [${fnRefs}] /Bounds [${bounds}] /Encode [${encode}] >>`]);
  return { mainObjNum: stitchNum, bodies };
}

function buildShadingBody(gradient, fnObjNum) {
  if (gradient.kind === "linear") {
    const c = gradient.coords;
    return `<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [${pdfNum(c.x1)} ${pdfNum(c.y1)} ${pdfNum(c.x2)} ${pdfNum(c.y2)}] /Function ${fnObjNum} 0 R /Extend [true true] >>`;
  }
  const c = gradient.coords;
  return `<< /ShadingType 3 /ColorSpace /DeviceRGB /Coords [${pdfNum(c.fx)} ${pdfNum(c.fy)} 0 ${pdfNum(c.cx)} ${pdfNum(c.cy)} ${pdfNum(c.r)}] /Function ${fnObjNum} 0 R /Extend [true true] >>`;
}

async function exportStagePDF(sourceStage, state, filename, exportOpts) {
  exportOpts = exportOpts || {};
  const w = TEMPLATE.canvas.width, h = TEMPLATE.canvas.height;

  // 1. Posiciones resueltas de cada capa, leídas directo del DOM en vivo —
  //    ya incluyen todo el layout automático (combos, agrupamiento
  //    OXXO+legal) y el ajuste manual de arrastre, sin recalcular nada. El
  //    giro/escala manual (drag-and-drop) NO se reflejan en left/top/
  //    offsetWidth/offsetHeight (transform no afecta la caja de layout) así
  //    que se leen aparte, directo del estado.
  const boxes = {};
  TEMPLATE.layers.forEach(L => {
    const el = sourceStage.querySelector(`[data-layer-id="${L.id}"]`);
    if (!el) return;
    const off = state.manualOffset && state.manualOffset[L.id];
    boxes[L.id] = {
      x: parseFloat(el.style.left) || 0, y: parseFloat(el.style.top) || 0,
      w: el.offsetWidth, h: el.offsetHeight, el,
      rot: (off && off.rot) || 0, scale: (off && off.scale) || 1,
    };
  });

  // Tamaño físico de página en puntos (1in = 72pt) — el mismo EXPORT_DPI que
  // ya marca el PNG, así el PDF imprime al tamaño real del formato (12.4in x
  // 6.2in) sin que el usuario tenga que reescalar nada en impresión. Se
  // calcula ACÁ (antes de lo que sigue) porque las matrices /Matrix de los
  // patterns de gradiente (paso 2) también la necesitan — un Pattern ignora
  // el "cm" activo del content stream, así que su matriz debe llegar hasta
  // puntos PDF por su cuenta (ver buildGradientPatternMatrix).
  const ptsPerPx = 72 / EXPORT_DPI;
  const pw = +(w * ptsPerPx).toFixed(2);
  const ph = +(h * ptsPerPx).toFixed(2);
  const cmPxToPt = `${pdfNum(ptsPerPx)} 0 0 ${pdfNum(-ptsPerPx)} 0 ${pdfNum(ph)} cm\n`;
  const cmPxToPtMat = { a: ptsPerPx, b: 0, c: 0, d: -ptsPerPx, e: 0, f: ph };

  // 2. producto/promo: vector cuando vienen de un SVG 100% vectorizable
  //    (ver svgCanFullyVectorize, ahora acepta relleno/trazo sólido O un
  //    gradiente lineal/radial soportado), raster en cualquier otro caso.
  //    Los gradientes se acumulan en pendingPatterns y se serializan más
  //    abajo, junto con los recortes raster.
  const pendingPatterns = [];
  function registerGradientPattern(gradient, patternMatrix) {
    const name = "Pg" + pendingPatterns.length;
    pendingPatterns.push({ name, gradient, matrix: patternMatrix });
    return name;
  }
  const vectorSvgOps = {};
  const vectorizedIds = [];
  ["producto", "promo"].forEach(id => {
    const asset = state[id];
    const box = boxes[id];
    if (!box || !isSvgAsset(asset)) return;
    const svgEl = box.el.querySelector("svg");
    if (!svgEl || !svgCanFullyVectorize(svgEl)) return;
    const placeMat = matRotateScaleAround(box.x + box.w / 2, box.y + box.h / 2, box.rot, box.scale, box.scale);
    const shapes = buildSvgVectorShapes(svgEl, box.x, box.y, placeMat);
    if (!shapes.length) return;
    vectorSvgOps[id] = emitVectorShapesOps(shapes, cmPxToPtMat, registerGradientPattern);
    vectorizedIds.push(id);
  });

  // 3. Texto legal: contornos reales (glifos trazados con opentype.js).
  const legalOps = boxes.legal ? buildLegalContentOps(boxes.legal, state.legal && state.legal.html, !!(state.legal && state.legal.outline), boxes.legal.rot, boxes.legal.scale) : "";
  if (boxes.legal && legalOps) vectorizedIds.push("legal");

  // 4. Fondo: rectángulo vectorial de color sólido (nunca gira: excluido del
  //    drag-and-drop, ver enableLayerDragging).
  const fondoOps = boxes.fondo ? emitFondoRectOps(boxes.fondo, (state.fondo && state.fondo.hex) || "#ffffff") : "";
  if (boxes.fondo) vectorizedIds.push("fondo");

  // 5. Todo lo demás (fotos, sello/edad18/oxxo, QR, vectores decorativos, y
  //    producto/promo cuando no se pudieron vectorizar) como recortes raster
  //    — SIN su giro/escala manual (si tiene): el pase raster compartido
  //    captura cada capa "derecha y a tamaño original", y el giro/escala se
  //    re-aplican vectorialmente en la matriz de colocación de cada XObject
  //    más abajo (así una capa girada o agrandada que "sobresale" de su caja
  //    original no queda recortada de más). Se excluyen acá mismo las cajas
  //    sin contenido propio (ver layerHasStateContent) — si no, una caja
  //    vacía cuyo rectángulo se superpone con el de una capa vecina "roba"
  //    los píxeles de esa vecina en el pase compartido (ver el comentario
  //    grande de layerHasStateContent).
  const rasterIds = TEMPLATE.layers.map(L => L.id).filter(id => boxes[id] && !vectorizedIds.includes(id) && layerHasStateContent(state, id));
  const rasterAssets = await buildRasterCropAssets(sourceStage, rasterIds, vectorizedIds, boxes);
  rasterAssets.forEach((a, idx) => { a.pdfName = "Im" + idx; });

  // 6. Línea de troquel (vector): mismo "d" y mismo criterio de escala que ya
  //    usa el recorte del PNG (viewBox 595.28x297.64 -> tamaño real del
  //    canvas), reutilizada dos veces más abajo — como CLIP global y, sin el
  //    clip, como el trazo negro visible.
  const troquelScaleX = w / SUAJE_VIEWBOX.w, troquelScaleY = h / SUAJE_VIEWBOX.h;
  const troquelCmds = transformCommands(parseSvgPathD(SUAJE_PATH_D), { a: troquelScaleX, b: 0, c: 0, d: troquelScaleY, e: 0, f: 0 });
  const troquelPathOps = commandsToPdfOps(troquelCmds);
  const troquelLineW = 0.3 * troquelScaleX; // mismo 10% del grosor original que en pantalla/PNG

  // ---- content stream: todo en espacio canvas-px, un único "cm" al inicio
  // de cada bloque q/Q convierte a puntos PDF (ver cmPxToPt arriba) ----
  let content = "q\n" + cmPxToPt;
  content += troquelPathOps + "W n\n"; // clip global a la silueta del troquel — igual que el PNG, transparencia real por fuera
  content += fondoOps;
  TEMPLATE.layers.forEach(L => {
    if (L.id === "fondo") return; // ya pintado arriba
    if (vectorSvgOps[L.id]) content += vectorSvgOps[L.id];
    if (L.id === "legal" && legalOps) content += legalOps;
    const asset = rasterAssets.find(a => a.id === L.id);
    if (asset) {
      const box = boxes[L.id];
      // Un Image XObject de PDF ubica el renglón 0 de sus datos en v=1 del
      // cuadrado unitario (arriba) — como acá ya se trabaja en un espacio
      // "canvas-px, y crece hacia abajo", ese cm local necesita su PROPIO
      // flip vertical (h negativo + trasladar y+h) o la imagen queda al
      // revés dentro de su propia caja, aunque el flip global de arriba ya
      // esté bien. El giro/escala manual (si los hay) se componen ENCIMA de
      // esa colocación, alrededor del centro de la caja original sin girar
      // ni escalar.
      let m = { a: asset.w, b: 0, c: 0, d: -asset.h, e: asset.x, f: asset.y + asset.h };
      if (box && (box.rot || (box.scale && box.scale !== 1))) {
        m = matMultiply(matRotateScaleAround(box.x + box.w / 2, box.y + box.h / 2, box.rot, box.scale, box.scale), m);
      }
      content += `q\n${pdfNum(m.a)} ${pdfNum(m.b)} ${pdfNum(m.c)} ${pdfNum(m.d)} ${pdfNum(m.e)} ${pdfNum(m.f)} cm\n/${asset.pdfName} Do\nQ\n`;
    }
  });
  content += "Q\n"; // suelta el clip ANTES de la línea de corte (si no, queda cortada a la mitad justo en el borde)
  content += "q\n" + cmPxToPt + `0 0 0 RG\n${pdfNum(troquelLineW)} w\n${troquelPathOps}S\nQ\n`;

  // ---- serializar el PDF ----
  let nextObj = 5; // 1..4 = catálogo/pages/page/content stream
  rasterAssets.forEach(a => { a.imgObjNum = nextObj++; a.smaskObjNum = nextObj++; });
  // Objetos de cada pattern de gradiente: Function(es) + Shading + Pattern.
  pendingPatterns.forEach(p => {
    const fnInfo = buildPdfGradientFunctionObjs(p.gradient.stops, () => nextObj++);
    p.functionObjNum = fnInfo.mainObjNum;
    p.functionBodies = fnInfo.bodies;
    p.shadingObjNum = nextObj++;
    p.shadingBody = buildShadingBody(p.gradient, p.functionObjNum);
    p.patternObjNum = nextObj++;
  });

  const enc = new TextEncoder();
  const parts = [];
  let offset = 0;
  const objOffset = {};
  function push(bytesOrStr) {
    const bytes = (typeof bytesOrStr === "string") ? enc.encode(bytesOrStr) : bytesOrStr;
    parts.push(bytes);
    offset += bytes.length;
  }
  function beginObj(n) { objOffset[n] = offset; push(n + " 0 obj\n"); }

  push("%PDF-1.4\n");

  beginObj(1);
  push("<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

  beginObj(2);
  push("<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n");

  const xobjEntries = rasterAssets.map(a => `/${a.pdfName} ${a.imgObjNum} 0 R`).join(" ");
  const patEntries = pendingPatterns.map(p => `/${p.name} ${p.patternObjNum} 0 R`).join(" ");
  beginObj(3);
  push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pw} ${ph}] /Resources << /XObject << ${xobjEntries} >>${patEntries ? ` /Pattern << ${patEntries} >>` : ""} >> /Contents 4 0 R >>\nendobj\n`);

  const contentBytes = enc.encode(content);
  beginObj(4);
  push(`<< /Length ${contentBytes.length} >>\nstream\n`);
  push(contentBytes);
  push("\nendstream\nendobj\n");

  for (const a of rasterAssets) {
    beginObj(a.imgObjNum);
    push(`<< /Type /XObject /Subtype /Image /Width ${a.w} /Height ${a.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /SMask ${a.smaskObjNum} 0 R /Length ${a.rgbZ.length} >>\nstream\n`);
    push(a.rgbZ);
    push("\nendstream\nendobj\n");

    beginObj(a.smaskObjNum);
    push(`<< /Type /XObject /Subtype /Image /Width ${a.w} /Height ${a.h} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode /Length ${a.alphaZ.length} >>\nstream\n`);
    push(a.alphaZ);
    push("\nendstream\nendobj\n");
  }

  for (const p of pendingPatterns) {
    p.functionBodies.forEach(([n, body]) => { beginObj(n); push(body + "\nendobj\n"); });
    beginObj(p.shadingObjNum);
    push(p.shadingBody + "\nendobj\n");
    beginObj(p.patternObjNum);
    const m = p.matrix;
    push(`<< /Type /Pattern /PatternType 2 /Shading ${p.shadingObjNum} 0 R /Matrix [${pdfNum(m.a)} ${pdfNum(m.b)} ${pdfNum(m.c)} ${pdfNum(m.d)} ${pdfNum(m.e)} ${pdfNum(m.f)}] >>\nendobj\n`);
  }

  const totalObjs = nextObj - 1;
  const xrefStart = offset;
  let xref = `xref\n0 ${totalObjs + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i <= totalObjs; i++) {
    xref += String(objOffset[i] || 0).padStart(10, "0") + " 00000 n \n";
  }
  push(xref);
  push(`trailer\n<< /Size ${totalObjs + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`);

  const blob = new Blob(parts, { type: "application/pdf" });
  triggerDownload(blob, filename || "render.pdf");
}

// ---------------------------------------------------------------------------
// Exportar PSD por capas (Photoshop) vía ag-psd (vendor/ag-psd.min.js) — una
// librería externa madura, ya probada contra Photoshop real, en vez del
// escritor binario armado a mano que usábamos antes (nos costó bastante
// dejarlo bien: orden de canales, "layer blending ranges", etc. — con
// ag-psd ese terreno ya está resuelto por otros). Cada capa de la plantilla
// sigue siendo una capa de PÍXELES rasterizada (mismo criterio que el PDF),
// salvo "LEGAL": esa sale como una capa de TEXTO real y editable de
// Photoshop (ver buildPsdLegalTextLayer), con su fuente/color/negrita-
// cursiva por rango de texto. OJO: la PRIMERA vez que se abre el archivo,
// Photoshop va a pedir "Actualizar" esa capa de texto (un clic) para
// redibujarla con su propio motor — recién ahí queda 100% editable; hasta
// entonces ya se ve bien, porque además le mandamos una vista previa
// rasterizada idéntica a la que ve el PDF.
// ---------------------------------------------------------------------------

// PostScript name real de cada variante de Inter (para que el "Actualizar"
// de Photoshop use la fuente correcta en vez de una negrita/cursiva
// sintética) — leído directo del archivo de fuente ya parseado por
// opentype.js (mismo INTER_FONTS que usa el PDF), con reserva por si algún
// día cambia el paquete de fuentes y falta el nombre.
function interPostScriptName(bold, italic) {
  const font = pickInterFont(bold, italic);
  const plat = font && font.names && (font.names.windows || font.names.macintosh || font.names.mac);
  return (plat && plat.postScriptName && plat.postScriptName.en) || "Inter-Regular";
}

// Caja alineada a ejes que alcanza a cubrir POR COMPLETO una caja de w x h
// escalada `scale` y luego girada `rot` grados, ambos alrededor de su propio
// centro — igual pivote que usa la pantalla/PDF (ver matRotateScaleAround) —
// recortada a los límites del canvas. El centro (cx,cy) se calcula con el
// w/h SIN escalar: transform-origin 50%/50% no se mueve al escalar.
function rotatedLayerCropRect(box, canvasW, canvasH) {
  const scale = box.scale || 1;
  const sw = box.w * scale, sh = box.h * scale;
  const rad = ((box.rot || 0) * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rad)), sin = Math.abs(Math.sin(rad));
  let newW = Math.ceil(sw * cos + sh * sin);
  let newH = Math.ceil(sw * sin + sh * cos);
  const cx = box.x + box.w / 2, cy = box.y + box.h / 2;
  let cropX = Math.round(cx - newW / 2), cropY = Math.round(cy - newH / 2);
  if (cropX < 0) { newW += cropX; cropX = 0; }
  if (cropY < 0) { newH += cropY; cropY = 0; }
  if (cropX + newW > canvasW) newW = canvasW - cropX;
  if (cropY + newH > canvasH) newH = canvasH - cropY;
  return { x: cropX, y: cropY, w: Math.max(0, Math.round(newW)), h: Math.max(0, Math.round(newH)) };
}

function getLivePsdLayerBox(sourceStage, state, id) {
  const el = sourceStage.querySelector(`[data-layer-id="${id}"]`);
  if (!el) return null;
  const off = state.manualOffset && state.manualOffset[id];
  return {
    x: parseFloat(el.style.left) || 0, y: parseFloat(el.style.top) || 0, w: el.offsetWidth, h: el.offsetHeight,
    rot: (off && off.rot) || 0, scale: (off && off.scale) || 1,
  };
}

// Recorta una región ya renderizada de un canvas grande a un canvas nuevo
// del tamaño exacto de esa región (para armar cada capa del PSD). `rect` va
// en las mismas unidades físicas que `sourceCanvas` (ya escaladas si
// corresponde — ver scaleRect).
function cropCanvasRegion(sourceCanvas, rect) {
  const out = document.createElement("canvas");
  out.width = Math.max(1, rect.w); out.height = Math.max(1, rect.h);
  out.getContext("2d").drawImage(sourceCanvas, rect.x, rect.y, rect.w, rect.h, 0, 0, rect.w, rect.h);
  return out;
}

// Lleva un rect en coordenadas LÓGICAS (0..TEMPLATE.canvas.width/height,
// las mismas que usan PNG/PDF) a coordenadas FÍSICAS del PSD, multiplicando
// por PSD_EXPORT_SCALE — ver la nota grande más abajo, antes de
// exportStagePSD.
function scaleRect(rect, scale) {
  return { x: Math.round(rect.x * scale), y: Math.round(rect.y * scale), w: Math.round(rect.w * scale), h: Math.round(rect.h * scale) };
}

// Rasteriza UNA sola capa de la plantilla (todas las demás ocultas) con su
// giro manual YA horneado en los píxeles (sin recorte al troquel — arte
// completo), recortada a rotatedLayerCropRect. null si queda 100%
// transparente (nada que aportar como capa de Photoshop). `scale` (ver
// PSD_EXPORT_SCALE) pide ese tanto más de píxeles por pulgada a
// buildExportCanvas — el recorte final se hace ya en esa resolución física
// más alta, no una ampliación posterior de un recorte chico.
async function buildPsdLayerAsset(sourceStage, state, layerId, layerName, scale) {
  const box = getLivePsdLayerBox(sourceStage, state, layerId);
  if (!box) return null;
  const rect = rotatedLayerCropRect(box, TEMPLATE.canvas.width, TEMPLATE.canvas.height);
  if (rect.w <= 0 || rect.h <= 0) return null;
  const hideIds = TEMPLATE.layers.map(L => L.id).filter(id => id !== layerId);
  const canvas = await buildExportCanvas(sourceStage, { clipToSuaje: false, keepSuajeTint: false, hideLayerIds: hideIds, scale });
  const physRect = scaleRect(rect, scale);
  const pixels = canvas.getContext("2d").getImageData(physRect.x, physRect.y, physRect.w, physRect.h).data;
  let hasContent = false;
  for (let k = 3; k < pixels.length; k += 4) { if (pixels[k] !== 0) { hasContent = true; break; } }
  if (!hasContent) return null;
  return { name: layerName, top: physRect.y, left: physRect.x, bottom: physRect.y + physRect.h, right: physRect.x + physRect.w, canvas: cropCanvasRegion(canvas, physRect) };
}

// A pedido: en el PSD exportado (no en pantalla ni en el PDF) el tamaño de
// letra de "legal" sale 7x (700%) más grande que el que hace entrar el
// texto en su caja, a propósito para que se desborde — ver
// buildPsdLegalTextLayer.
const LEGAL_TEXT_PSD_SCALE = 7;

// Capa "LEGAL": a diferencia de las demás (rasterizadas desde una foto del
// DOM), esta se arma directo con fitLegalTextVector/legalLinesToGlyphItems —
// el mismo motor de layout que ya usa el PDF para vectorizar este texto en
// curvas — así la posición/tamaño de letra de la vista previa acá y del PDF
// coinciden exactamente, y de paso se reutiliza parseLegalRuns para separar
// el HTML enriquecido (negrita/cursiva/color por span) en los "runs" que
// necesita la capa de texto real de Photoshop. Devuelve null si no hay texto
// o si la capa tiene un giro o escala manual (caso raro; un transform con
// rotación/escala para texto editable de Photoshop suma complejidad para un
// caso de uso que no se da en la práctica) — en cualquiera de esos casos,
// "legal" cae al camino raster normal como las demás capas (que sí soporta
// giro/escala, ver rotatedLayerCropRect/buildPsdLayerAsset), aunque ahí
// pierde el texto editable y el desborde a propósito de LEGAL_TEXT_PSD_SCALE.
function buildPsdLegalTextLayer(box, legalHtml, outlineOn, scale) {
  if (box.rot || (box.scale && box.scale !== 1)) return null;
  if (!legalHtml || !legalHtml.replace(/<[^>]+>/g, "").trim()) return null;

  // fitLegalTextVector calcula el tamaño que hace ENTRAR el texto en la
  // caja (mismo criterio que el PDF) — eso define los saltos de línea
  // (lines/lineHeightFactor), pero a pedido el tamaño de letra que
  // realmente se exporta al PSD es LEGAL_TEXT_PSD_SCALE veces más grande
  // a propósito, para que se desborde de la caja. Como los saltos de línea
  // ya están decididos con el tamaño chico, el desborde es sobre todo
  // horizontal (cada línea queda más ancha) y vertical (más alto por
  // renglón) — no se vuelven a recalcular los cortes de línea.
  //
  // `box`, `fontSize`, `items` (x/y/fontSize) están todos en coordenadas
  // LÓGICAS (0..TEMPLATE.canvas.width/height, las mismas que PNG/PDF) —
  // recién al dibujar/declarar el rect de la capa se multiplica por `scale`
  // (PSD_EXPORT_SCALE), IGUAL que en buildPsdLayerAsset, salvo el tamaño de
  // letra en PUNTOS (fontSizePt más abajo), que es independiente de la
  // resolución del documento y por eso NO se multiplica por scale.
  const w = TEMPLATE.canvas.width * scale, h = TEMPLATE.canvas.height * scale;
  const { fontSize: fitFontSize, lines, lineHeightFactor } = fitLegalTextVector(box, legalHtml);
  const fontSize = fitFontSize * LEGAL_TEXT_PSD_SCALE;
  const items = legalLinesToGlyphItems(box, fontSize, lineHeightFactor, lines);

  // Vista previa rasterizada (para que se vea bien desde que se abre, antes
  // del clic en "Actualizar" que hace falta para que quede editable de
  // verdad) — a TODO el canvas del documento (no solo la caja de "legal"),
  // porque a este tamaño el texto se sale de esa caja. Mismo truco de 2
  // pasadas que el halo del PDF (emitGlyphItemsHalo + emitGlyphItemsColored):
  // todos los glifos en blanco relleno+trazo grueso primero, y recién
  // después el relleno de color real por encima.
  const canvas = document.createElement("canvas");
  canvas.width = w; canvas.height = h;
  const ctx = canvas.getContext("2d");
  const strokeW = Math.max(1, fontSize * 0.09) * scale;
  if (outlineOn) {
    items.forEach(it => {
      const path = it.glyph.getPath(it.x * scale, it.y * scale, it.fontSize * scale);
      path.fill = "#ffffff";
      path.stroke = "#ffffff";
      path.strokeWidth = strokeW * 2;
      path.draw(ctx);
    });
  }
  items.forEach(it => {
    const path = it.glyph.getPath(it.x * scale, it.y * scale, it.fontSize * scale);
    path.fill = `rgb(${it.color.r},${it.color.g},${it.color.b})`;
    path.stroke = null;
    path.draw(ctx);
  });

  // Runs de texto/estilo reales (lo que hace falta para que Photoshop pueda
  // editarlo de verdad tras "Actualizar") — reusa parseLegalRuns, el mismo
  // tokenizador que arma los runs para el PDF.
  const runs = parseLegalRuns(legalHtml);
  let text = "";
  const styleRuns = [];
  runs.forEach(r => {
    if (r.brk) { text += "\n"; styleRuns.push({ length: 1, style: {} }); return; }
    if (!r.text) return;
    text += r.text;
    styleRuns.push({
      length: r.text.length,
      style: { font: { name: interPostScriptName(r.bold, r.italic) }, fillColor: anyTextColorToRgb(r.color) },
    });
  });

  // Photoshop guarda el tamaño de letra de una capa de texto SIEMPRE en
  // PUNTOS (72pt = 1 pulgada), sin importar el DPI del documento — por eso
  // se convierte acá fontSize (en píxeles, a EXPORT_DPI, YA multiplicado
  // por LEGAL_TEXT_PSD_SCALE) a puntos.
  const fontSizePt = (fontSize * 72) / EXPORT_DPI;

  const layer = {
    name: "LEGAL",
    top: 0, left: 0, bottom: h, right: w, // toda la vista previa ocupa el canvas completo (ver arriba)
    canvas,
    text: {
      text,
      transform: [1, 0, 0, 1, box.x * scale, box.y * scale], // posición absoluta = origen de la caja
      shapeType: "box",
      boxBounds: [0, 0, box.w * scale, box.h * scale], // ancho/alto RELATIVOS al origen del transform
      style: { font: { name: interPostScriptName(false, false) }, fontSize: fontSizePt, fillColor: { r: 0, g: 0, b: 0 } },
      styleRuns,
      paragraphStyle: { justification: "left" },
    },
  };

  if (outlineOn) {
    // Halo blanco "Position: Outside" como efecto de capa NATIVO de
    // Photoshop (no un truco de píxeles) — mismo look que applyLegalOutline
    // en pantalla y el trazo grueso del PDF, pero este sobrevive al
    // "Actualizar" (el trazo de la vista previa no).
    layer.effects = {
      stroke: [{
        enabled: true, position: "outside", fillType: "color",
        color: { r: 255, g: 255, b: 255 }, opacity: 100, blendMode: "normal",
        size: { units: "Pixels", value: strokeW },
      }],
    };
  }

  return layer;
}

// Capa "máscara" del troquel: el mismo asset SUAJE.svg que nos pasaron
// (blanco por fuera de la línea de corte, hueco/transparente por dentro,
// donde va el arte) — a pedido: en vez de una máscara de capa "de verdad" en
// cada capa o un grupo, esta capa se coloca arriba de todas las demás y,
// como tapa en blanco sólido todo lo que queda fuera del troquel, hace el
// mismo efecto visual de corte que el recorte real del PNG. Blanco sólido
// en todo el canvas, con un agujero 100% transparente exactamente en la
// silueta del troquel.
function buildPsdTroquelMaskAsset(scale) {
  const w = TEMPLATE.canvas.width * scale, h = TEMPLATE.canvas.height * scale;
  const canvas = document.createElement("canvas");
  canvas.width = w; canvas.height = h;
  const ctx = canvas.getContext("2d");
  const scaleX = w / SUAJE_VIEWBOX.w, scaleY = h / SUAJE_VIEWBOX.h;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, w, h);
  ctx.save();
  ctx.scale(scaleX, scaleY);
  ctx.globalCompositeOperation = "destination-out"; // perfora el hueco (alpha 0) dentro del troquel
  ctx.fill(new Path2D(SUAJE_PATH_D));
  ctx.restore();
  return { name: "TROQUEL — máscara (blanco)", top: 0, left: 0, bottom: h, right: w, canvas };
}

// Capa-guía del troquel (solo la línea de corte, sin el tinte de aviso que
// se ve en pantalla) — puramente vectorial vía canvas, no necesita pasar por
// el pipeline de foreignObject.
function buildPsdTroquelGuideAsset(scale) {
  const w = TEMPLATE.canvas.width * scale, h = TEMPLATE.canvas.height * scale;
  const canvas = document.createElement("canvas");
  canvas.width = w; canvas.height = h;
  const ctx = canvas.getContext("2d");
  const scaleX = w / SUAJE_VIEWBOX.w, scaleY = h / SUAJE_VIEWBOX.h;
  ctx.save();
  ctx.scale(scaleX, scaleY);
  ctx.lineWidth = 0.3 * scale;
  ctx.strokeStyle = "#000000";
  ctx.stroke(new Path2D(SUAJE_PATH_D));
  ctx.restore();
  return { name: "LÍNEA DE TROQUEL (guía)", top: 0, left: 0, bottom: h, right: w, canvas };
}

// A pedido: el PSD exportado sale a más resolución física que el PNG/PDF —
// mismo tamaño impreso (12.4in x 6.2in), más píxeles por pulgada, para que
// aguante mejor si alguien escala el archivo o alguna capa después. Es un
// factor sobre EXPORT_DPI (300) exclusivo del PSD: 2 = 600dpi equivalente.
// Una capa de píxeles de Photoshop no puede tener "más resolución que las
// demás" por sí sola (eso es justo lo que hace un Objeto Inteligente, que
// se descartó por complejidad/riesgo) — así que sube el documento entero.
const PSD_EXPORT_SCALE = 2;

async function exportStagePSD(sourceStage, state, filename) {
  const scale = PSD_EXPORT_SCALE;
  const w = TEMPLATE.canvas.width * scale, h = TEMPLATE.canvas.height * scale;

  // 1. Cada capa de la plantilla, de abajo hacia arriba (mismo orden que
  //    TEMPLATE.layers = z-index de pantalla). "legal" sale como capa de
  //    TEXTO real y editable (ver buildPsdLegalTextLayer); el resto,
  //    rasterizada con su giro manual horneado (Photoshop no tiene una
  //    "matriz de colocación" para una capa de píxeles simple), recortada a
  //    rotatedLayerCropRect. A propósito, ninguna capa se recorta al
  //    troquel (son el "material fuente" editable) — el troquel va aparte,
  //    como máscara blanca + línea guía arriba de todo, y la imagen
  //    combinada final sí lo respeta, igual que el PNG.
  const layerAssets = [];
  for (const L of TEMPLATE.layers) {
    if (L.id === "legal") {
      const box = getLivePsdLayerBox(sourceStage, state, "legal");
      const textLayer = box && buildPsdLegalTextLayer(box, state.legal && state.legal.html, !!(state.legal && state.legal.outline), scale);
      if (textLayer) { layerAssets.push(textLayer); continue; }
    }
    const asset = await buildPsdLayerAsset(sourceStage, state, L.id, L.name, scale);
    if (asset) layerAssets.push(asset);
  }

  // 2. Arriba de todo: la máscara blanca del troquel y, por encima de esa,
  //    la línea de corte como guía fina — en ese orden para que la línea no
  //    quede tapada por el blanco de la máscara.
  layerAssets.push(buildPsdTroquelMaskAsset(scale));
  layerAssets.push(buildPsdTroquelGuideAsset(scale));

  // 3. Imagen combinada (la "miniatura" que usan lectores sin soporte de
  //    capas) — misma composición que el PNG: recortada al troquel, alpha
  //    real, a la misma resolución más alta que el resto del documento (el
  //    "Image Data" final de un PSD tiene que medir width x height exacto).
  const compositeCanvas = await buildExportCanvas(sourceStage, { scale });

  // 4. Escribir el PSD con ag-psd (vendor/ag-psd.min.js) — se encarga de
  //    canales, compresión, blending ranges y de la capa de texto (TySh).
  const psd = {
    width: w,
    height: h,
    children: layerAssets,
    canvas: compositeCanvas,
    imageResources: {
      resolutionInfo: {
        horizontalResolution: EXPORT_DPI * scale, horizontalResolutionUnit: "PPI", widthUnit: "Inches",
        verticalResolution: EXPORT_DPI * scale, verticalResolutionUnit: "PPI", heightUnit: "Inches",
      },
    },
  };
  const buffer = agPsd.writePsd(psd, { generateThumbnail: true });
  const blob = new Blob([buffer], { type: "image/vnd.adobe.photoshop" });
  triggerDownload(blob, filename || "render.psd");
}

// ============================================================================
// Cargar el texto legal directamente desde una celda de Excel (.xlsx),
// conservando negrita/cursiva/color REALES de la celda — sin pasar por el
// portapapeles del sistema operativo (que es justo lo que falla cuando la
// fuente no expone HTML real al copiar, ver el flujo de pegado más arriba).
//
// Un .xlsx es un zip con XML adentro (formato OOXML). Se lee con JSZip
// (vendor/jszip.min.js) y se parsean a mano las tres piezas relevantes:
//   - xl/sharedStrings.xml: el texto de una celda de tipo "s" vive acá, como
//     un <si> — que puede traer un solo <t> (texto plano) o varios <r> (un
//     "run" por cada tramo con su propio estilo: así es como Excel guarda
//     "seleccioné una palabra y le puse negrita/color con el mouse").
//   - xl/styles.xml: fonts[] + cellXfs[] — el estilo de fuente de TODA la
//     celda (cuando no hay runs, o como respaldo del color de cada run).
//   - xl/theme/theme1.xml: los 10 colores base del tema del workbook, para
//     resolver un color elegido desde la fila de "Colores del tema" del
//     selector de Excel (el caso más común en la práctica) — con el mismo
//     ajuste de brillo ("tint") que aplica Excel.
// Cada celda se convierte a un HTML con el mismo formato que ya entiende
// sanitizeLegalHtml/parseLegalRuns (<span style="font-weight:bold;...">),
// así que se reutiliza TODO el resto del pipeline (PNG/PDF/PSD) sin tocarlo.
const XLSX_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XLSX_NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const XLSX_NS_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XLSX_NS_OFFICE_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function xlsxParseXml(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function xlsxFlagPresent(parentEl, tag) {
  // OOXML: la sola PRESENCIA de <b/> ya significa "negrita" (val ausente =
  // true); val="0"/"false" es la única forma de decir "no".
  const el = parentEl.getElementsByTagNameNS(XLSX_NS_MAIN, tag)[0];
  if (!el) return false;
  const val = el.getAttribute("val");
  return val === null || val === "1" || val === "true";
}

// ---- color: rgb directo, o theme+tint resuelto contra la paleta del workbook ----
function xlsxRgbToHls(r, g, b) {
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  return [h, l, s];
}
function xlsxHlsToRgb(h, l, s) {
  if (s === 0) return [l, l, l];
  const hue2rgb = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue2rgb(p, q, h + 1 / 3), hue2rgb(p, q, h), hue2rgb(p, q, h - 1 / 3)];
}
// Mismo cálculo que usa Excel para variantes tipo "Énfasis 1, Claro 40%":
// ajusta la luminosidad en espacio HLS. tint 0/null = color sin cambios.
function xlsxApplyTint(hexRgb, tint) {
  if (!tint) return hexRgb;
  const r = parseInt(hexRgb.slice(0, 2), 16) / 255;
  const g = parseInt(hexRgb.slice(2, 4), 16) / 255;
  const b = parseInt(hexRgb.slice(4, 6), 16) / 255;
  let [h, l, s] = xlsxRgbToHls(r, g, b);
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  l = Math.max(0, Math.min(1, l));
  const [r2, g2, b2] = xlsxHlsToRgb(h, l, s);
  const toHex = v => Math.round(Math.max(0, Math.min(255, v * 255))).toString(16).padStart(2, "0").toUpperCase();
  return toHex(r2) + toHex(g2) + toHex(b2);
}

const XLSX_THEME_INDEX_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
function xlsxParseTheme(themeXml) {
  const scheme = themeXml.getElementsByTagNameNS(XLSX_NS_DRAWING, "clrScheme")[0];
  if (!scheme) return null;
  const colors = XLSX_THEME_INDEX_ORDER.map(tag => {
    const wrap = scheme.getElementsByTagNameNS(XLSX_NS_DRAWING, tag)[0];
    if (!wrap) return null;
    const srgb = wrap.getElementsByTagNameNS(XLSX_NS_DRAWING, "srgbClr")[0];
    if (srgb) return srgb.getAttribute("val");
    const sysClr = wrap.getElementsByTagNameNS(XLSX_NS_DRAWING, "sysClr")[0];
    if (sysClr) return sysClr.getAttribute("lastClr");
    return null;
  });
  return colors.some(c => c === null) ? null : colors;
}

// idx 0/1 (fondo1/texto1) SIN tint es el negro/blanco "automático" de una
// celda nunca tocada — no una elección real de color, se ignora (igual que
// "sin color explícito") para no pisar el color base de la plantilla.
function xlsxColorFromEl(colorEl, themeColors) {
  if (!colorEl) return null;
  const rgb = colorEl.getAttribute("rgb");
  if (rgb) {
    let r, g, b;
    if (rgb.length === 8) { r = rgb.slice(2, 4); g = rgb.slice(4, 6); b = rgb.slice(6, 8); }
    else if (rgb.length === 6) { r = rgb.slice(0, 2); g = rgb.slice(2, 4); b = rgb.slice(4, 6); }
    else return null;
    return `rgb(${parseInt(r, 16)},${parseInt(g, 16)},${parseInt(b, 16)})`;
  }
  const themeAttr = colorEl.getAttribute("theme");
  if (themeAttr !== null && themeColors) {
    const idx = parseInt(themeAttr, 10);
    const tint = parseFloat(colorEl.getAttribute("tint") || "0");
    if (isNaN(idx) || idx < 0 || idx >= themeColors.length) return null;
    if ((idx === 0 || idx === 1) && Math.abs(tint) < 1e-6) return null;
    const hex = xlsxApplyTint(themeColors[idx], tint);
    return `rgb(${parseInt(hex.slice(0, 2), 16)},${parseInt(hex.slice(2, 4), 16)},${parseInt(hex.slice(4, 6), 16)})`;
  }
  return null;
}

// ---- fonts.xml / cellXfs: el estilo de fuente de una celda COMPLETA ----
function xlsxParseStyles(stylesXml, themeColors) {
  const fontsEl = stylesXml.getElementsByTagNameNS(XLSX_NS_MAIN, "fonts")[0];
  const fonts = fontsEl ? Array.from(fontsEl.getElementsByTagNameNS(XLSX_NS_MAIN, "font")).map(f => ({
    bold: xlsxFlagPresent(f, "b"),
    italic: xlsxFlagPresent(f, "i"),
    color: xlsxColorFromEl(f.getElementsByTagNameNS(XLSX_NS_MAIN, "color")[0], themeColors),
  })) : [];
  const cellXfsEl = stylesXml.getElementsByTagNameNS(XLSX_NS_MAIN, "cellXfs")[0];
  const cellXfs = cellXfsEl ? Array.from(cellXfsEl.getElementsByTagNameNS(XLSX_NS_MAIN, "xf")).map(xf => {
    const fontId = xf.getAttribute("fontId");
    return fontId !== null ? parseInt(fontId, 10) : null;
  }) : [];
  return { fonts, cellXfs };
}

// ---- una celda de texto (<si> de sharedStrings, o <is> inline): texto
// plano (un solo <t>) o varios <r> con su propio estilo cada uno ----
function xlsxParseRunThemed(rEl, themeColors) {
  const rPr = rEl.getElementsByTagNameNS(XLSX_NS_MAIN, "rPr")[0];
  const tEl = rEl.getElementsByTagNameNS(XLSX_NS_MAIN, "t")[0];
  const text = tEl ? tEl.textContent : "";
  if (!rPr) return { text, bold: false, italic: false, color: null };
  return {
    text,
    bold: xlsxFlagPresent(rPr, "b"),
    italic: xlsxFlagPresent(rPr, "i"),
    color: xlsxColorFromEl(rPr.getElementsByTagNameNS(XLSX_NS_MAIN, "color")[0], themeColors),
  };
}
function xlsxParseSi(siEl, themeColors) {
  const runs = Array.from(siEl.getElementsByTagNameNS(XLSX_NS_MAIN, "r"));
  if (runs.length) return { runs: runs.map(r => xlsxParseRunThemed(r, themeColors)) };
  const tEl = siEl.getElementsByTagNameNS(XLSX_NS_MAIN, "t")[0];
  return { text: tEl ? tEl.textContent : "" };
}

// ---- une los runs de una celda con el estilo de fuente de la celda
// completa (respaldo de color cuando no hay rich text; negrita/cursiva/color
// cuando la celda es texto plano sin runs) y arma el HTML final ----
function xlsxCellToRuns(parsed, cellFont) {
  const cellColor = cellFont ? cellFont.color : null;
  const cellBold = !!(cellFont && cellFont.bold);
  const cellItalic = !!(cellFont && cellFont.italic);
  if (parsed.runs) {
    return parsed.runs.map(r => ({
      text: r.text,
      bold: r.bold,
      italic: r.italic,
      color: r.color || cellColor,
    }));
  }
  return [{ text: parsed.text || "", bold: cellBold, italic: cellItalic, color: cellColor }];
}
function xlsxRunsToHtml(runs) {
  return runs.map(r => {
    const lines = String(r.text || "").split(/\r\n|\r|\n/).map(escapeHtmlText).join("<br>");
    if (!lines) return "";
    const styles = [];
    if (r.bold) styles.push("font-weight:bold");
    if (r.italic) styles.push("font-style:italic");
    if (r.color) styles.push(`color:${r.color}`);
    return styles.length ? `<span style="${styles.join(";")}">${lines}</span>` : lines;
  }).join("");
}

function xlsxDecodeRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return { row: 0, col: 0 };
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(m[2], 10), col };
}
function xlsxColLetter(n) {
  let s = "";
  while (n > 0) { const rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// límite de la vista previa (celdas más allá de esto simplemente no se
// muestran para pegarles clic — cualquier hoja de promos real es mucho más
// chica que esto)
const XLSX_PREVIEW_MAX_ROW = 300;
const XLSX_PREVIEW_MAX_COL = 60;

function xlsxParseWorksheet(sheetXml, sharedStrings, styleInfo, themeColors) {
  const cells = {};
  let maxRow = 0, maxCol = 0;
  Array.from(sheetXml.getElementsByTagNameNS(XLSX_NS_MAIN, "row")).forEach(rowEl => {
    Array.from(rowEl.getElementsByTagNameNS(XLSX_NS_MAIN, "c")).forEach(cEl => {
      const ref = cEl.getAttribute("r");
      if (!ref) return;
      const { row, col } = xlsxDecodeRef(ref);
      if (row < 1 || col < 1 || row > XLSX_PREVIEW_MAX_ROW || col > XLSX_PREVIEW_MAX_COL) return;

      const type = cEl.getAttribute("t") || "n";
      const sIdx = cEl.getAttribute("s");
      const fontId = (sIdx !== null && styleInfo.cellXfs[+sIdx] != null) ? styleInfo.cellXfs[+sIdx] : null;
      const cellFont = fontId !== null ? styleInfo.fonts[fontId] : null;

      let parsed;
      if (type === "s") {
        const vEl = cEl.getElementsByTagNameNS(XLSX_NS_MAIN, "v")[0];
        const idx = vEl ? parseInt(vEl.textContent, 10) : NaN;
        parsed = sharedStrings[idx] || { text: "" };
      } else if (type === "inlineStr") {
        const isEl = cEl.getElementsByTagNameNS(XLSX_NS_MAIN, "is")[0];
        parsed = isEl ? xlsxParseSi(isEl, themeColors) : { text: "" };
      } else {
        const vEl = cEl.getElementsByTagNameNS(XLSX_NS_MAIN, "v")[0];
        parsed = { text: vEl ? vEl.textContent : "" };
      }

      const runs = xlsxCellToRuns(parsed, cellFont);
      const plainText = runs.map(r => r.text).join("");
      if (!plainText.trim()) return;
      const html = xlsxRunsToHtml(runs);
      cells[ref] = { text: plainText, html };
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    });
  });
  return { cells, maxRow, maxCol };
}

async function xlsxReadXmlFromZip(zip, path) {
  const f = zip.file(path);
  if (!f) return null;
  return xlsxParseXml(await f.async("text"));
}

// Punto de entrada: recibe el ArrayBuffer del .xlsx subido y devuelve
// { sheetNames, sheetsData } — sheetsData[name] = { cells, maxRow, maxCol }
// (ver xlsxParseWorksheet), listo para pintar la grilla de selección y sacar
// el HTML de la celda que el usuario elija.
async function xlsxLoadWorkbookForPicker(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const wbXml = await xlsxReadXmlFromZip(zip, "xl/workbook.xml");
  const relsXml = await xlsxReadXmlFromZip(zip, "xl/_rels/workbook.xml.rels");
  if (!wbXml || !relsXml) throw new Error("Archivo .xlsx inválido o dañado.");

  const relMap = {};
  Array.from(relsXml.getElementsByTagNameNS(XLSX_NS_RELS, "Relationship")).forEach(r => {
    relMap[r.getAttribute("Id")] = "xl/" + r.getAttribute("Target").replace(/^\/?(xl\/)?/, "");
  });

  const sheetRefs = Array.from(wbXml.getElementsByTagNameNS(XLSX_NS_MAIN, "sheet")).map(s => {
    const rId = s.getAttributeNS(XLSX_NS_OFFICE_RELS, "id") || s.getAttribute("r:id");
    return { name: s.getAttribute("name"), path: relMap[rId] };
  }).filter(s => s.path);
  if (!sheetRefs.length) throw new Error("El workbook no tiene hojas legibles.");

  const themeXml = await xlsxReadXmlFromZip(zip, "xl/theme/theme1.xml");
  const themeColors = themeXml ? xlsxParseTheme(themeXml) : null;

  const stylesXml = await xlsxReadXmlFromZip(zip, "xl/styles.xml");
  const styleInfo = stylesXml ? xlsxParseStyles(stylesXml, themeColors) : { fonts: [], cellXfs: [] };

  const sstXml = await xlsxReadXmlFromZip(zip, "xl/sharedStrings.xml");
  const sharedStrings = sstXml ? Array.from(sstXml.getElementsByTagNameNS(XLSX_NS_MAIN, "si")).map(si => xlsxParseSi(si, themeColors)) : [];

  const sheetsData = {};
  for (const s of sheetRefs) {
    const sheetXml = await xlsxReadXmlFromZip(zip, s.path);
    if (!sheetXml) continue;
    sheetsData[s.name] = xlsxParseWorksheet(sheetXml, sharedStrings, styleInfo, themeColors);
  }
  return { sheetNames: sheetRefs.map(s => s.name).filter(n => sheetsData[n]), sheetsData };
}

// Pinta la grilla de selección (tabla clicable) dentro de `container` a
// partir de un sheetData de xlsxLoadWorkbookForPicker. `onPick(ref, html)`
// se llama con la celda elegida y su HTML ya armado (bold/italic/color).
function xlsxRenderPreviewGrid(container, sheetData, onPick) {
  container.innerHTML = "";
  if (!sheetData.maxRow || !sheetData.maxCol) {
    const empty = document.createElement("div");
    empty.className = "xlsx-empty";
    empty.textContent = "Esta hoja no tiene celdas con texto.";
    container.appendChild(empty);
    return;
  }
  const table = document.createElement("table");
  table.className = "xlsx-grid";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headRow.appendChild(document.createElement("th"));
  for (let c = 1; c <= sheetData.maxCol; c++) {
    const th = document.createElement("th");
    th.textContent = xlsxColLetter(c);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (let r = 1; r <= sheetData.maxRow; r++) {
    const tr = document.createElement("tr");
    const rowTh = document.createElement("th");
    rowTh.textContent = String(r);
    tr.appendChild(rowTh);
    for (let c = 1; c <= sheetData.maxCol; c++) {
      const ref = xlsxColLetter(c) + r;
      const cell = sheetData.cells[ref];
      const td = document.createElement("td");
      if (cell) {
        td.textContent = cell.text.length > 60 ? cell.text.slice(0, 60) + "…" : cell.text;
        td.title = cell.text;
        td.dataset.ref = ref;
        td.addEventListener("click", () => onPick(ref, cell.html));
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  container.appendChild(table);
}
