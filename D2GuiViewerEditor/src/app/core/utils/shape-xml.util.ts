/**
 * Edycja pass-through XML kształtów (ADR-0056/0063): `data-docx-xml` niesie ORYGINALNY
 * OOXML (base64), z którego writer odtwarza grafikę przy zapisie. Przesunięcie/rozmiar
 * zmienione tylko w stylach GUI ginęłyby przy pierwszym zapisie — źródłem prawdy jest XML,
 * więc drag/resize musi aktualizować także jego geometrię.
 *
 * Kontrakt pozycji (ADR-0030 + reader ResolveAnchorPosition): X od lewej krawędzi STRONY,
 * Y od góry OBSZARU TREŚCI (= Y_page − marginTop). Zapis normalizuje obie osie do
 * relativeFrom="page" (offset = współrzędna page-relative w EMU) — reader odczytuje ją
 * identycznie w body i w pasmach (gałąź "page" nie zależy od kontekstu pasma).
 */

export const EMU_PER_PIXEL = 9525;

const NS = {
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  v: 'urn:schemas-microsoft-com:vml',
} as const;

/** base64 → XMLDocument (null przy uszkodzonym markerze — caller zostawia XML w spokoju). */
export function decodeShapeXml(b64: string): XMLDocument | null {
  try {
    const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    const xml = new TextDecoder().decode(bytes);
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
  } catch {
    return null;
  }
}

/** XMLDocument → base64 (UTF-8, kształt odporny na znaki spoza Latin-1 w nazwach docPr). */
export function encodeShapeXml(doc: XMLDocument): string {
  const xml = new XMLSerializer().serializeToString(doc);
  const bytes = new TextEncoder().encode(xml);
  let bin = '';
  bytes.forEach(b => (bin += String.fromCharCode(b)));
  return btoa(bin);
}

function firstNS(root: ParentNode, ns: string, local: string): Element | null {
  return (root as Element | XMLDocument).getElementsByTagNameNS?.(ns, local)?.[0] ?? null;
}

function setPositionAxis(doc: XMLDocument, anchor: Element, local: 'positionH' | 'positionV', pageEmu: number): void {
  let pos = anchor.getElementsByTagNameNS(NS.wp, local)[0] ?? null;
  if (!pos) {
    pos = doc.createElementNS(NS.wp, local);
    anchor.appendChild(pos);
  }
  pos.setAttribute('relativeFrom', 'page');
  // wp:align i wp:posOffset są alternatywami — po jawnym przesunięciu zostaje sam offset.
  Array.from(pos.children).forEach(c => c.remove());
  const off = doc.createElementNS(NS.wp, 'posOffset');
  off.textContent = String(Math.round(pageEmu));
  pos.appendChild(off);
}

/** Ustawia wartość osi w stylu VML (margin-left/margin-top/width/height, pt). */
function setVmlStyleProp(el: Element, prop: string, valuePt: number): void {
  const style = el.getAttribute('style') ?? '';
  const parts = style.split(';').map(s => s.trim()).filter(s => s && !s.startsWith(`${prop}:`));
  parts.push(`${prop}:${valuePt.toFixed(2)}pt`);
  el.setAttribute('style', parts.join(';'));
}

/**
 * Przesuwa kotwicę na współrzędne page-relative (EMU): każdy wp:anchor w dokumencie
 * (gałąź mc:Choice) + top-level element VML Fallbacku (żeby stary Word nie pokazał
 * kształtu w poprzednim miejscu). Zwraca false dla grafik inline (wp:inline nie ma
 * pozycji — przesuwanie nie ma czego aktualizować).
 */
export function setAnchorPositionPageEmu(doc: XMLDocument, xPageEmu: number, yPageEmu: number): boolean {
  const anchors = Array.from(doc.getElementsByTagNameNS(NS.wp, 'anchor'));
  if (anchors.length === 0) return false;
  anchors.forEach(anchor => {
    setPositionAxis(doc, anchor, 'positionH', xPageEmu);
    setPositionAxis(doc, anchor, 'positionV', yPageEmu);
  });
  // VML fallback: pozycja w inline style (pt) + mso-position-*-relative:page.
  for (const local of ['group', 'shape', 'rect', 'oval', 'line'] as const) {
    for (const el of Array.from(doc.getElementsByTagNameNS(NS.v, local))) {
      const style = el.getAttribute('style') ?? '';
      if (!style.includes('position:absolute')) continue;
      if ((el.parentNode as Element | null)?.namespaceURI === NS.v) continue; // tylko top-level
      setVmlStyleProp(el, 'margin-left', xPageEmu / EMU_PER_PIXEL * 0.75);
      setVmlStyleProp(el, 'margin-top', yPageEmu / EMU_PER_PIXEL * 0.75);
      const st = (el.getAttribute('style') ?? '').split(';').map(s => s.trim())
        .filter(s => s && !s.startsWith('mso-position-horizontal-relative')
          && !s.startsWith('mso-position-vertical-relative'));
      st.push('mso-position-horizontal-relative:page', 'mso-position-vertical-relative:page');
      el.setAttribute('style', st.join(';'));
    }
  }
  return true;
}

/**
 * Skaluje rozmiar grafiki: wp:extent (rozmiar w układzie strony), root-owe a:xfrm/a:ext
 * (grupa: dzieci skalują się przez stosunek ext/chExt — chExt celowo NIETKNIĘTE; pojedynczy
 * kształt: przestrzeń custGeom normalizuje się do viewBoxu) oraz width/height stylu VML
 * Fallbacku. Działa dla wp:anchor I wp:inline.
 */
export function scaleShapeExtent(doc: XMLDocument, fx: number, fy: number): boolean {
  let touched = false;
  for (const extent of Array.from(doc.getElementsByTagNameNS(NS.wp, 'extent'))) {
    const cx = parseInt(extent.getAttribute('cx') ?? '', 10);
    const cy = parseInt(extent.getAttribute('cy') ?? '', 10);
    if (Number.isFinite(cx)) extent.setAttribute('cx', String(Math.max(0, Math.round(cx * fx))));
    if (Number.isFinite(cy)) extent.setAttribute('cy', String(Math.max(0, Math.round(cy * fy))));
    touched = true;
  }
  // Root xfrm: grupa (wpg:grpSpPr) albo pojedynczy kształt (wps:spPr) NAJWYŻSZEGO poziomu.
  const rootProps = firstNS(doc, NS.wpg, 'grpSpPr') ?? firstNS(doc, NS.wps, 'spPr');
  const xfrm = rootProps ? firstNS(rootProps, NS.a, 'xfrm') : null;
  const ext = xfrm ? firstNS(xfrm, NS.a, 'ext') : null;
  if (ext) {
    const cx = parseInt(ext.getAttribute('cx') ?? '', 10);
    const cy = parseInt(ext.getAttribute('cy') ?? '', 10);
    if (Number.isFinite(cx)) ext.setAttribute('cx', String(Math.max(0, Math.round(cx * fx))));
    if (Number.isFinite(cy)) ext.setAttribute('cy', String(Math.max(0, Math.round(cy * fy))));
    touched = true;
  }
  for (const local of ['group', 'shape', 'rect', 'oval'] as const) {
    for (const el of Array.from(doc.getElementsByTagNameNS(NS.v, local))) {
      if ((el.parentNode as Element | null)?.namespaceURI === NS.v) continue;
      const style = el.getAttribute('style') ?? '';
      const w = /(?:^|;)\s*width:([0-9.]+)pt/.exec(style)?.[1];
      const h = /(?:^|;)\s*height:([0-9.]+)pt/.exec(style)?.[1];
      if (w) setVmlStyleProp(el, 'width', parseFloat(w) * fx);
      if (h) setVmlStyleProp(el, 'height', parseFloat(h) * fy);
    }
  }
  return touched;
}

/**
 * Skaluje PODGLĄD kształtu w DOM edytora: boks kontenera + inline geometrie dzieci
 * (divy/imgi grup mają absolutne px, SVG ma width/height w atrybutach; viewBox +
 * preserveAspectRatio=none rozciągają ścieżki, stroke jest non-scaling).
 */
export function rescaleShapePreview(shape: HTMLElement, fx: number, fy: number): void {
  const scalePx = (v: string, f: number): string | null => {
    const n = parseFloat(v);
    return Number.isFinite(n) ? `${(n * f).toFixed(2)}px` : null;
  };
  const applyTo = (el: HTMLElement): void => {
    if (el.style.left) el.style.left = scalePx(el.style.left, fx) ?? el.style.left;
    if (el.style.top) el.style.top = scalePx(el.style.top, fy) ?? el.style.top;
    if (el.style.width) el.style.width = scalePx(el.style.width, fx) ?? el.style.width;
    if (el.style.height) el.style.height = scalePx(el.style.height, fy) ?? el.style.height;
  };
  applyTo(shape);
  shape.querySelectorAll<HTMLElement>('div, img').forEach(applyTo);
  shape.querySelectorAll('svg').forEach(svg => {
    const w = parseFloat(svg.getAttribute('width') ?? '');
    const h = parseFloat(svg.getAttribute('height') ?? '');
    if (Number.isFinite(w)) svg.setAttribute('width', (w * fx).toFixed(2));
    if (Number.isFinite(h)) svg.setAttribute('height', (h * fy).toFixed(2));
  });
}
