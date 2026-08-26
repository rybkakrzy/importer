/**
 * Odstępy „przed"/„po" akapitem — kontrakt z readerem/writerem (ADR-0107).
 *
 * Word domyślnie liczy odstęp między akapitami jako MAX(after A, before B) — dokładnie tak,
 * jak kolapsują marginesy rodzeństwa w CSS. Dlatego domyślnym nośnikiem odstępu „po" jest
 * `margin-bottom`. Wyjątek: dokument z flagą zgodności `doNotUseHTMLParagraphAutoSpacing`
 * (Word SUMUJE after + before) — reader oznacza go `data-para-spacing-sum="1"`, strony
 * `.editor-content` dostają klasę `para-spacing-sum`, a nośnikiem „po" jest wtedy
 * `padding-bottom` (nie kolapsuje). Writer czyta oba nośniki.
 *
 * Odstępy w jednostkach linii (`w:beforeLines`/`w:afterLines`) reader niesie markerami
 * `--w-before-lines` / `--w-after-lines` obok wartości w pt. Ręczna edycja wartości w pt
 * musi te markery zdjąć — inaczej writer odtworzyłby wersję liniową, która w Wordzie
 * WYGRYWA z wartością w pt.
 */

export const PARAGRAPH_SPACING_SUM_CLASS = 'para-spacing-sum';

/** Czy blok żyje w dokumencie z modelem SUMY odstępów (flaga zgodności Worda). */
export function isSumSpacingModel(el: Element): boolean {
  return !!el.closest('.editor-content')?.classList.contains(PARAGRAPH_SPACING_SUM_CLASS);
}

/** Ustawia odstęp „przed" (CSS length, np. `12pt`/`16px`/`0`). */
export function setParagraphSpaceBefore(el: HTMLElement, value: string): void {
  el.style.marginTop = value;
  el.style.removeProperty('--w-before-lines');
}

/**
 * Ustawia odstęp „po" na nośniku właściwym dla modelu dokumentu i czyści drugi nośnik,
 * żeby wartości nie sumowały się podwójnie (treść sprzed ADR-0107 / akapity z tłem).
 */
export function setParagraphSpaceAfter(el: HTMLElement, value: string): void {
  if (isSumSpacingModel(el)) {
    el.style.paddingBottom = value;
    el.style.marginBottom = '';
  } else {
    el.style.marginBottom = value;
    el.style.paddingBottom = '';
  }
  el.style.removeProperty('--w-after-lines');
}

/**
 * Odstęp „przed" w px: wartość ZADEKLAROWANA inline (jak dialog Worda — pokazuje 6 pt
 * także wtedy, gdy contextualSpacing znosi ją wizualnie), inaczej wartość efektywna
 * z kaskady (domyślny odstęp dokumentu).
 */
export function readParagraphSpaceBeforePx(el: HTMLElement): number {
  const inline = parseFloat(el.style.marginTop);
  if (Number.isFinite(inline)) return inline * inlineLengthToPx(el.style.marginTop, el);
  return parseFloat(getComputedStyle(el).marginTop) || 0;
}

/** Odstęp „po" w px — suma obu nośników (jeden z nich jest zawsze 0). */
export function readParagraphSpaceAfterPx(el: HTMLElement): number {
  const inlineMargin = el.style.marginBottom;
  const inlinePadding = el.style.paddingBottom;
  if (inlineMargin || inlinePadding) {
    return (
      (parseFloat(inlineMargin) || 0) * inlineLengthToPx(inlineMargin, el) +
      (parseFloat(inlinePadding) || 0) * inlineLengthToPx(inlinePadding, el)
    );
  }
  const cs = getComputedStyle(el);
  return (parseFloat(cs.marginBottom) || 0) + (parseFloat(cs.paddingBottom) || 0);
}

/** Mnożnik jednostki inline → px (reader pisze pt, dialog px; inne jednostki = 1). */
function inlineLengthToPx(value: string, el: HTMLElement): number {
  if (!value) return 1;
  if (value.endsWith('pt')) return 96 / 72;
  if (value.endsWith('em')) return parseFloat(getComputedStyle(el).fontSize) || 16;
  return 1;
}
