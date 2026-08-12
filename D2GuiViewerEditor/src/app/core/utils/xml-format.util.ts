import { XmlToken, tokenizeXml } from './xml-highlight.util';

export interface XmlViewLine {
  /** Numer wiersza liczony od 1. W trybie surowym odpowiada linii w pliku źródłowym. */
  number: number;
  depth: number;
  text: string;
  tokens: XmlToken[];
  /** Wiersz otwiera element, który ma zawartość — da się go zwinąć. */
  foldable: boolean;
  /** Numer wiersza zamykającego (dla wiersza zwijalnego). */
  foldEnd: number | null;
}

/** Powyżej tego rozmiaru podgląd nie formatuje ani nie koloruje — chodzi o responsywność UI. */
export const XML_FORMAT_LIMIT = 400_000;

/**
 * Wiersze surowego XML dokładnie tak, jak w pliku — numeracja zgadza się z numerem linii
 * zwracanym przez backend, więc podświetlenie elementu trafia we właściwe miejsce.
 */
export function toRawLines(xml: string): XmlViewLine[] {
  const colorize = xml.length <= XML_FORMAT_LIMIT;

  return xml.split('\n').map((text, index) => ({
    number: index + 1,
    depth: 0,
    text,
    tokens: colorize ? tokenizeXml(text) : [{ kind: 'text' as const, text }],
    foldable: false,
    foldEnd: null,
  }));
}

/**
 * Wiersze sformatowane: jeden znacznik na wiersz, wcięcie wg zagnieżdżenia, zakresy do zwijania.
 *
 * To skaner, nie parser — musi zadziałać także dla XML, którego parser by odrzucił, bo właśnie
 * takie dokumenty diagnozuje to narzędzie. Treść znaczników nie jest modyfikowana.
 */
export function toFormattedLines(xml: string): XmlViewLine[] {
  if (xml.length > XML_FORMAT_LIMIT) {
    return toRawLines(xml);
  }

  const lines: XmlViewLine[] = [];
  const openLines: number[] = [];
  let depth = 0;
  let index = 0;

  const push = (text: string, lineDepth: number, foldable: boolean): number => {
    lines.push({
      number: lines.length + 1,
      depth: lineDepth,
      text,
      tokens: tokenizeXml(text),
      foldable,
      foldEnd: null,
    });
    return lines.length - 1;
  };

  while (index < xml.length) {
    const tagStart = xml.indexOf('<', index);

    if (tagStart < 0) {
      appendText(xml.slice(index));
      break;
    }

    appendText(xml.slice(index, tagStart));

    const tagEnd = findTagEnd(xml, tagStart);

    if (tagEnd < 0) {
      push(xml.slice(tagStart), depth, false);
      break;
    }

    const tag = xml.slice(tagStart, tagEnd + 1);
    index = tagEnd + 1;

    if (isClosingTag(tag)) {
      depth = Math.max(0, depth - 1);
      const lineIndex = push(tag, depth, false);
      const openIndex = openLines.pop();

      if (openIndex !== undefined) {
        lines[openIndex].foldEnd = lines[lineIndex].number;
        lines[openIndex].foldable = lines[lineIndex].number > lines[openIndex].number + 1;
      }

      continue;
    }

    if (isSelfContainedTag(tag)) {
      push(tag, depth, false);
      continue;
    }

    openLines.push(push(tag, depth, false));
    depth++;
  }

  return lines;

  function appendText(text: string): void {
    const trimmed = text.trim();

    if (trimmed.length > 0) {
      push(trimmed, depth, false);
    }
  }
}

/** Numery wierszy ukrytych przez zwinięte zakresy. */
export function collectHiddenLines(lines: XmlViewLine[], collapsed: ReadonlySet<number>): Set<number> {
  const hidden = new Set<number>();

  for (const line of lines) {
    if (!collapsed.has(line.number) || line.foldEnd === null) {
      continue;
    }

    for (let number = line.number + 1; number <= line.foldEnd; number++) {
      hidden.add(number);
    }
  }

  return hidden;
}

/**
 * Koniec znacznika. Znak `>` wewnątrz wartości atrybutu nie kończy znacznika, więc skaner śledzi
 * cudzysłowy — inaczej atrybuty ze ścieżkami XPath rozjeżdżałyby podgląd.
 */
function findTagEnd(xml: string, tagStart: number): number {
  let quote: string | null = null;

  for (let index = tagStart + 1; index < xml.length; index++) {
    const character = xml[index];

    if (quote) {
      if (character === quote) {
        quote = null;
      }
      continue;
    }

    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }

    if (character === '>') {
      return index;
    }
  }

  return -1;
}

function isClosingTag(tag: string): boolean {
  return tag.startsWith('</');
}

/** Znacznik, który nie otwiera zagnieżdżenia: samozamykający, deklaracja, komentarz, CDATA, DOCTYPE. */
function isSelfContainedTag(tag: string): boolean {
  return tag.endsWith('/>') || tag.startsWith('<?') || tag.startsWith('<!');
}
