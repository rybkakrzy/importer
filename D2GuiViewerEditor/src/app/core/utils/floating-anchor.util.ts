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
