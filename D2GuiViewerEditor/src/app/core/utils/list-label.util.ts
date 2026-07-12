/**
 * Silnik etykiet list DOCX dla podglądu edytora (odpowiednik liczników readera
 * `DocxToHtmlConverter` — wspólna semantyka, patrz ADR-0036 wariant A).
 *
 * Przeglądarka nie renderuje poprawnie: szablonów wielopoziomowych (`%1.%2.%3` → „2.4.1."),
 * sufiksów lvlText innych niż kropka („1)", „Art. 1."), numeracji legal (w:isLgl),
 * kontynuacji między fragmentami po edycji (natywny `<ol start>` jest statyczny) ani
 * `startOverride`. Silnik liczy etykiety z kontraktu data-* na kontenerach ul/ol
 * (data-num-id / data-abstract-num-id / data-ilvl / data-num-fmt / data-lvl-text /
 * data-start / data-start-override / data-is-legal / data-lvl-restart) i zapisuje je
 * w `data-list-label` na `li` — render robi CSS `::before` (poza edytowalnym DOM,
 * nie kopiuje się z tekstem i nie trafia do zapisu).
 *
 * Semantyka Worda (identyczna z licznikami readera):
 * - licznik żyje per ABSTRAKT (instancje wspólnego abstraktu kontynuują numerację),
 * - startOverride restartuje licznik przy PIERWSZYM użyciu instancji,
 * - element na poziomie L restartuje poziomy głębsze (lvlRestart=0 wyłącza restart,
 *   lvlRestart=N ogranicza restart do użycia poziomów < N — wartość JEDNOBAZOWA),
 * - punktory też konsumują pozycję (restartują głębsze poziomy numerowane).
 */

/** Atrybuty listy odczytane z kontenera ul/ol (kontrakt data-* readera). */
export interface ListContainerAttrs {
  ordered: boolean;
  numId: string;
  /** Klucz licznika: abstrakt (wspólny dla instancji), fallback per-numId. */
  counterKey: string;
  level: number;
  fmt: string;
  lvlText: string | null;
  start: number;
  startOverride: number;
  isLegal: boolean;
  /** Surowa wartość w:lvlRestart (jednobazowa; 0 = nigdy; -1 = brak). */
  lvlRestart: number;
  suffix: 'tab' | 'space' | 'nothing';
  picBullet: boolean;
}

/** Odczyt kontraktu data-* z kontenera listy; null gdy lista nie pochodzi z DOCX. */
export function readListContainerAttrs(el: Element): ListContainerAttrs | null {
  const numId = el.getAttribute('data-num-id');
  if (!numId) return null;
  const abstractId = el.getAttribute('data-abstract-num-id');
  const intAttr = (name: string, fallback: number): number => {
    const v = parseInt(el.getAttribute(name) ?? '', 10);
    return Number.isFinite(v) ? v : fallback;
  };
  const suffixRaw = el.getAttribute('data-suffix');
  return {
    ordered: el.tagName.toLowerCase() === 'ol',
    numId,
    counterKey: abstractId !== null ? `abs:${abstractId}` : `num:${numId}`,
    level: Math.min(8, Math.max(0, intAttr('data-ilvl', 0))),
    fmt: el.getAttribute('data-num-fmt') ?? 'decimal',
    lvlText: el.getAttribute('data-lvl-text'),
    start: intAttr('data-start', 1),
    startOverride: intAttr('data-start-override', -1),
    isLegal: el.getAttribute('data-is-legal') === '1',
    lvlRestart: intAttr('data-lvl-restart', -1),
    suffix: suffixRaw === 'space' || suffixRaw === 'nothing' ? suffixRaw : 'tab',
    picBullet: el.getAttribute('data-pic-bullet') === '1',
  };
}

/** Formatuje wartość licznika wg tokenu w:numFmt (nazwy jak w data-num-fmt). */
export function formatListNumber(fmt: string, n: number): string {
  if (n < 1) n = 1;
  switch (fmt) {
    case 'decimalZero':
      return n < 10 ? `0${n}` : String(n);
    case 'lowerLetter':
      return toWordLetter(n);
    case 'upperLetter':
      return toWordLetter(n).toUpperCase();
    case 'lowerRoman':
      return toRoman(n).toLowerCase();
    case 'upperRoman':
      return toRoman(n);
    default:
      return String(n);
  }
}

/** Litery jak w Wordzie: a..z, potem POWTÓRZONA litera (27=aa, 28=bb, 53=aaa). */
function toWordLetter(n: number): string {
  const repeat = Math.ceil(n / 26);
  const letter = String.fromCharCode(97 + ((n - 1) % 26));
  return letter.repeat(repeat);
}

function toRoman(n: number): string {
  const map: Array<[number, string]> = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
    [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
    [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let result = '';
  for (const [value, symbol] of map) {
    while (n >= value) {
      result += symbol;
      n -= value;
    }
  }
  return result || 'I';
}

interface LevelSpec {
  fmt: string;
  start: number;
  lvlRestart: number;
}

/**
 * Stanowy silnik liczników — jedno przejście po dokumencie w kolejności DOM.
 * Po edycji tworzony od zera (deterministyczny, O(n) po elementach list).
 */
export class ListLabelEngine {
  /** counterKey → wartości liczników per poziom (undefined = poziom nieużyty). */
  private readonly counters = new Map<string, Array<number | undefined>>();
  /** numId-y, których startOverride już zastosowano (reset tylko przy 1. użyciu instancji). */
  private readonly appliedOverrides = new Set<string>();
  /** (counterKey, level) → definicja poziomu (pierwsza napotkana wygrywa — jak w writerze). */
  private readonly specs = new Map<string, LevelSpec>();

  /**
   * Konsumuje kolejny element listy i zwraca etykietę do wyświetlenia
   * (null dla punktorów/formatów bullet|none — te renderuje istniejący mechanizm).
   */
  next(attrs: ListContainerAttrs): string | null {
    this.registerSpec(attrs);
    this.applyStartOverrideOnFirstUse(attrs);

    const row = this.countersFor(attrs.counterKey);
    const spec = this.specFor(attrs.counterKey, attrs.level);
    row[attrs.level] = row[attrs.level] === undefined ? spec.start : row[attrs.level]! + 1;

    // Element na poziomie L restartuje poziomy głębsze (semantyka Worda):
    // lvlRestart=0 → nigdy; lvlRestart=N (jednobazowe) → restart tylko po poziomach < N.
    for (let deeper = attrs.level + 1; deeper <= 8; deeper++) {
      if (row[deeper] === undefined) continue;
      const deeperRestart = this.specs.get(`${attrs.counterKey}|${deeper}`)?.lvlRestart ?? -1;
      if (deeperRestart === 0) continue;
      if (deeperRestart > 0 && attrs.level >= deeperRestart) continue;
      row[deeper] = undefined;
    }

    if (attrs.fmt === 'bullet' || attrs.fmt === 'none' || attrs.picBullet) return null;
    return this.buildLabel(attrs, row);
  }

  private buildLabel(attrs: ListContainerAttrs, row: Array<number | undefined>): string {
    const template =
      attrs.lvlText && attrs.lvlText.includes('%') ? attrs.lvlText : `%${attrs.level + 1}.`;
    return template.replace(/%(\d)/g, (_m, digit: string) => {
      const refLevel = parseInt(digit, 10) - 1;
      const refSpec = this.specFor(attrs.counterKey, refLevel);
      const value = row[refLevel] ?? refSpec.start;
      // w:isLgl: wszystkie poziomy w etykiecie formatowane jako decimal.
      const fmt = attrs.isLegal ? 'decimal' : refLevel === attrs.level ? attrs.fmt : refSpec.fmt;
      return formatListNumber(fmt, value);
    });
  }

  private registerSpec(attrs: ListContainerAttrs): void {
    const key = `${attrs.counterKey}|${attrs.level}`;
    if (!this.specs.has(key)) {
      this.specs.set(key, { fmt: attrs.fmt, start: attrs.start, lvlRestart: attrs.lvlRestart });
    }
  }

  private specFor(counterKey: string, level: number): LevelSpec {
    return this.specs.get(`${counterKey}|${level}`) ?? { fmt: 'decimal', start: 1, lvlRestart: -1 };
  }

  private applyStartOverrideOnFirstUse(attrs: ListContainerAttrs): void {
    if (attrs.startOverride <= 0 || this.appliedOverrides.has(attrs.numId)) return;
    this.appliedOverrides.add(attrs.numId);
    const row = this.countersFor(attrs.counterKey);
    row[attrs.level] = attrs.startOverride - 1;
  }

  private countersFor(counterKey: string): Array<number | undefined> {
    let row = this.counters.get(counterKey);
    if (!row) {
      row = new Array<number | undefined>(9).fill(undefined);
      this.counters.set(counterKey, row);
    }
    return row;
  }
}

/**
 * Przelicza etykiety wszystkich list DOCX w podanych kontenerach (strony edytora
 * W KOLEJNOŚCI DOKUMENTU — kontynuacja działa przez granice stron i fragmentów).
 *
 * Efekt na DOM: `li` list numerowanych z data-* dostaje `data-list-label`
 * (+ `data-list-suffix` dla suffix ≠ tab); punktory i listy bez data-* są nietykane
 * (renderuje je istniejący mechanizm: natywny marker / span.list-marker).
 * Atrybuty są czysto prezentacyjne — serializacja zapisu musi je zdejmować.
 */
export function applyListLabels(roots: Iterable<HTMLElement>): void {
  const engine = new ListLabelEngine();
  for (const root of roots) {
    // querySelectorAll zwraca kolejność dokumentu — zagnieżdżone li trafiają
    // we właściwym miejscu między elementy rodzica (jak konsumpcja w readerze).
    const items = root.querySelectorAll<HTMLLIElement>(
      'ol[data-num-id] > li, ul[data-num-id] > li',
    );
    items.forEach(li => {
      const container = li.parentElement;
      const attrs = container ? readListContainerAttrs(container) : null;
      if (!attrs) return;
      const label = engine.next(attrs);
      if (label !== null) {
        li.setAttribute('data-list-label', label);
        if (attrs.suffix !== 'tab') li.setAttribute('data-list-suffix', attrs.suffix);
        else li.removeAttribute('data-list-suffix');
      } else {
        li.removeAttribute('data-list-label');
        li.removeAttribute('data-list-suffix');
      }
    });
  }
}

/** Zdejmuje prezentacyjne atrybuty etykiet przed serializacją do zapisu. */
export function stripListLabelAttributes(root: Element): void {
  root.querySelectorAll('[data-list-label], [data-list-suffix]').forEach(el => {
    el.removeAttribute('data-list-label');
    el.removeAttribute('data-list-suffix');
  });
}

/**
 * Uzupełnia brakujące znaczniki `span.list-marker` w listach punktowanych z DOCX.
 *
 * Reader emituje marker span (własny symbol, marker tekstowy „TODO:", obraz punktatora)
 * per `li` — ale `li` utworzony ENTEREM w edytorze go nie ma (Chrome dzieli element bez
 * klonowania pierwszego dziecka), więc nowy element listy tracił widoczny punktator.
 * Wzorzec klonujemy z rodzeństwa tego samego kontenera; TYLKO dodajemy (usuwanie węzłów
 * w contenteditable grozi utratą kotwicy kursora), a eksport i tak pomija marker spany.
 */
export function ensureBulletMarkers(roots: Iterable<HTMLElement>): void {
  for (const root of roots) {
    root.querySelectorAll<HTMLElement>('ul[data-num-id], ol[data-num-id]').forEach(container => {
      const items = Array.from(container.children).filter(
        (c): c is HTMLLIElement => c.tagName === 'LI',
      );
      const template = items
        .map(li => li.querySelector<HTMLElement>(':scope > span.list-marker'))
        .find(span => span !== null);
      if (!template) return;
      for (const li of items) {
        if (li.querySelector(':scope > span.list-marker')) continue;
        li.insertBefore(template.cloneNode(true), li.firstChild);
      }
    });
  }
}
