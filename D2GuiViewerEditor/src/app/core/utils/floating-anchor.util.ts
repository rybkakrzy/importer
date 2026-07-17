/**
 * Model kotwicy elementów pływających (obrazy front/behind, pola tekstowe docx-textbox).
 *
 * Kotwica NIE jest sztucznym identyfikatorem — to relacja pozycyjna w DOM, ta sama,
 * którą serializuje backend (DocxToHtmlConverter / HtmlToDocxConverter):
 *  - pływający OBRAZ leży WEWNĄTRZ swojego akapitu-kotwicy (span.editor-image-wrapper w <p>),
 *  - POLE TEKSTOWE (div.docx-textbox) jest emitowane bezpośrednio PRZED akapitem-kotwicą
 *    (blokowy div w <p> byłby re-parentowany przez parser przeglądarki), a writer przypina
 *    je z powrotem do NASTĘPNEGO akapitu.
 * Dzięki temu relacja przeżywa zapis DOCX, undo/redo (snapshoty HTML) i ponowny render.
 */

export const EMU_PER_PX = 9525;

const ANCHOR_BLOCK_SELECTOR = 'p, h1, h2, h3, h4, h5, h6, li';

/**
 * Geometria pasma nagłówka/stopki potrzebna do przeliczenia kotwicy między układem KONTRAKTU
 * (data-x-emu/data-y-emu: X od lewej krawędzi STRONY, Y od góry OBSZARU TREŚCI = y_page − marginTop
 * — tak emituje reader i tak zapisuje writer: wp:anchor X=Page/Y=Margin) a układem PASMA
 * (absolut wewnątrz .header-display/.header-editor-content, których origin jest przesunięty o
 * padding-left pasma = margines lewy i o offset pasma od góry strony).
 */
export interface HfBandGeometry {
  band: 'header' | 'footer';
  marginLeftPx: number;
  marginTopPx: number;
  /** Górna krawędź KONTENERA pasma od góry strony: header = dystans nagłówka;
   *  footer = wysokość strony − dystans stopki − pasmo stopki (min-height, przybliżenie). */
  bandTopPx: number;
}

/**
 * Kontrakt → układ pasma: pozycja `left/top` absolutu wewnątrz kontenera nagłówka/stopki,
 * przy której obraz ląduje w tym samym miejscu strony co w MS Word. Wartości mogą być ujemne
 * (obiekt wystaje poza pasmo — np. logo sięgające nad dystans nagłówka); kontenery pasm mają
 * overflow visible, więc render jest poprawny.
 */
export function contractToBand(xPx: number, yPx: number, geo: HfBandGeometry): { leftPx: number; topPx: number } {
  return {
    leftPx: xPx - geo.marginLeftPx,
    topPx: yPx + geo.marginTopPx - geo.bandTopPx,
  };
}

/** Układ pasma → kontrakt (dokładna odwrotność contractToBand) — do zapisu data-x/y-emu po dragu. */
export function bandToContract(leftPx: number, topPx: number, geo: HfBandGeometry): { xPx: number; yPx: number } {
  return {
    xPx: leftPx + geo.marginLeftPx,
    yPx: topPx - geo.marginTopPx + geo.bandTopPx,
  };
}

/** Czy element jest pływający (pozycjonowany absolutnie względem strony). */
export function isFloatingElement(el: HTMLElement): boolean {
  const mode = el.dataset['posMode'] ?? '';
  return mode === 'front' || mode === 'behind' || el.style.position === 'absolute';
}

/**
 * Akapit, do którego zakotwiczony jest element pływający.
 * Obraz → najbliższy blokowy przodek; textbox → następny brat-akapit
 * (fallback: poprzedni brat, potem blokowy przodek — np. textbox w <li>).
 */
export function findAnchorParagraph(el: HTMLElement, editorRoot: HTMLElement): HTMLElement | null {
  if (el.classList.contains('docx-textbox')) {
    let next = el.nextElementSibling;
    while (next && !next.matches(ANCHOR_BLOCK_SELECTOR)) next = next.nextElementSibling;
    if (next) return next as HTMLElement;

    let prev = el.previousElementSibling;
    while (prev && !prev.matches(ANCHOR_BLOCK_SELECTOR)) prev = prev.previousElementSibling;
    if (prev) return prev as HTMLElement;

    const parentBlock = el.parentElement?.closest(ANCHOR_BLOCK_SELECTOR) as HTMLElement | null;
    return parentBlock && editorRoot.contains(parentBlock) ? parentBlock : null;
  }

  const block = el.closest(ANCHOR_BLOCK_SELECTOR) as HTMLElement | null;
  return block && editorRoot.contains(block) ? block : null;
}

/**
 * Pozycja znacznika kotwicy w px UKŁADU strony (bez skali zoomu) — znacznik jest
 * absolutnie pozycjonowanym dzieckiem .page, więc transform zoomu skaluje go razem
 * z treścią i nie wymaga przeliczeń przy zmianie zoomu/scrolla.
 * Recty wejściowe są we współrzędnych viewportu (getBoundingClientRect); dzielenie
 * przez aktualną skalę (szerokość rect / szerokość layoutu strony) sprowadza je do
 * przestrzeni layoutu.
 */
export function computeAnchorBadgePosition(
  paragraphRect: DOMRect,
  pageRect: DOMRect,
  pageLayoutWidthPx: number,
  badgeSizePx = 18,
): { leftPx: number; topPx: number } {
  const scale = pageLayoutWidthPx > 0 && pageRect.width > 0 ? pageRect.width / pageLayoutWidthPx : 1;
  const topPx = (paragraphRect.top - pageRect.top) / scale;
  const leftPx = (paragraphRect.left - pageRect.left) / scale - badgeSizePx - 4;
  return { leftPx: Math.max(2, Math.round(leftPx)), topPx: Math.max(0, Math.round(topPx)) };
}

/** Delta wskaźnika (px viewportu) → px układu treści przy zoomie `scale`. */
export function viewportDeltaToLayout(deltaPx: number, scale: number): number {
  return scale > 0 ? deltaPx / scale : deltaPx;
}

/** Czy zdarzenie trafiło w strefę krawędzi elementu (pas `edgePx` od brzegu, w px viewportu). */
export function isPointerOnEdge(
  clientX: number,
  clientY: number,
  rect: DOMRect,
  edgePx = 8,
): boolean {
  const inside =
    clientX >= rect.left && clientX <= rect.right && clientY >= rect.top && clientY <= rect.bottom;
  if (!inside) return false;
  return (
    clientX - rect.left <= edgePx ||
    rect.right - clientX <= edgePx ||
    clientY - rect.top <= edgePx ||
    rect.bottom - clientY <= edgePx
  );
}
