import { TableBorderScope, TableBorderSettings } from '../../models/table-style.model';

/**
 * Czyste operacje na obramowaniach tabeli (poziom DOM). Linie są utrwalane jako
 * style inline na komórkach, więc przeżywają zapis HTML → DOCX i ponowne otwarcie.
 * Funkcje są bezstanowe i testowalne w jsdom. Treść komórek pozostaje nietknięta —
 * modyfikujemy wyłącznie właściwości `border*`, nigdy `innerHTML`.
 */

export const DEFAULT_CELL_BORDER_COLOR = '#cccccc';
const CELL_PADDING = '8px';
const CELL_MIN_WIDTH = '30px';

function cellsOf(table: HTMLTableElement): HTMLTableCellElement[] {
  return Array.from(table.querySelectorAll('td, th')) as HTMLTableCellElement[];
}

function ensureCellBox(cell: HTMLTableCellElement): void {
  if (!cell.style.padding) cell.style.padding = CELL_PADDING;
  if (!cell.style.minWidth) cell.style.minWidth = CELL_MIN_WIDTH;
}

function borderCss(border: TableBorderSettings): string {
  if (border.style === 'none' || border.width <= 0) return 'none';
  return `${border.width}px ${border.style} ${border.color}`;
}

/**
 * Stosuje linię w wybranym zakresie do **dowolnego zbioru komórek** (cel: cała
 * tabela / pojedyncza komórka / zaznaczony fragment). Krawędzie „outer/inner/top…"
 * liczone są względem prostokąta opisanego na przekazanych komórkach, dzięki czemu
 * np. „zewnętrzne" rysuje obrys zaznaczenia, a nie całej tabeli. Większość zakresów
 * jest addytywna (dokłada krawędzie) — jak przyciski obramowań w MS Word; `all`
 * ustawia pełne obramowanie komórki, `none` czyści wszystkie krawędzie.
 * Pozycje (wiersz/kolumna) bierzemy z DOM (`rowIndex`/`cellIndex`).
 */
export function applyBorderToCells(
  cells: HTMLTableCellElement[],
  scope: TableBorderScope,
  border: TableBorderSettings
): void {
  if (cells.length === 0) return;
  const css = borderCss(border);

  let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
  const positioned = cells.map(cell => {
    const r = (cell.parentElement as HTMLTableRowElement).rowIndex;
    const c = cell.cellIndex;
    minR = Math.min(minR, r); maxR = Math.max(maxR, r);
    minC = Math.min(minC, c); maxC = Math.max(maxC, c);
    return { cell, r, c };
  });

  positioned.forEach(({ cell, r, c }) => {
    ensureCellBox(cell);
    switch (scope) {
      case 'all':
        cell.style.border = css;
        break;
      case 'none':
        cell.style.border = 'none';
        break;
      case 'outer':
        if (r === minR) cell.style.borderTop = css;
        if (r === maxR) cell.style.borderBottom = css;
        if (c === minC) cell.style.borderLeft = css;
        if (c === maxC) cell.style.borderRight = css;
        break;
      case 'inner':
        if (r !== maxR) cell.style.borderBottom = css;
        if (c !== maxC) cell.style.borderRight = css;
        break;
      case 'inner-horizontal':
        if (r !== maxR) cell.style.borderBottom = css;
        break;
      case 'inner-vertical':
        if (c !== maxC) cell.style.borderRight = css;
        break;
      case 'top':
        if (r === minR) cell.style.borderTop = css;
        break;
      case 'bottom':
        if (r === maxR) cell.style.borderBottom = css;
        break;
      case 'left':
        if (c === minC) cell.style.borderLeft = css;
        break;
      case 'right':
        if (c === maxC) cell.style.borderRight = css;
        break;
    }
  });
}

/** Wariant dla całej tabeli (zakres liczony względem pełnej siatki). */
export function applyBorderScope(
  table: HTMLTableElement,
  scope: TableBorderScope,
  border: TableBorderSettings
): void {
  applyBorderToCells(cellsOf(table), scope, border);
}

export type BorderTargetKind = 'cell' | 'range' | 'row' | 'column' | 'table' | 'none';

export interface BorderTargetInfo {
  kind: BorderTargetKind;
  /** Liczba wierszy objętych prostokątem zaznaczenia. */
  rows: number;
  /** Liczba kolumn objętych prostokątem zaznaczenia. */
  cols: number;
}

/**
 * Rozpoznaje intencję użytkownika z zestawu komórek (auto-zakres obramowania):
 * pełna szerokość i wysokość → cała tabela; pełna wysokość → kolumna(y); pełna
 * szerokość → wiersz(e); jedna komórka → komórka; w innym razie → zakres komórek.
 * Pozycje czytane z DOM (`rowIndex`/`cellIndex`). Tabele z mocno scalonymi
 * komórkami mogą dać przybliżony wynik (etykieta), ale samo rysowanie obramowania
 * i tak działa na przekazanym zbiorze komórek.
 */
export function classifyBorderTarget(
  table: HTMLTableElement,
  cells: HTMLTableCellElement[]
): BorderTargetInfo {
  if (cells.length === 0) return { kind: 'none', rows: 0, cols: 0 };

  const totalRows = table.rows.length;
  const totalCols = Math.max(...Array.from(table.rows).map(r => r.cells.length), 0);

  let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
  cells.forEach(cell => {
    const r = (cell.parentElement as HTMLTableRowElement).rowIndex;
    const c = cell.cellIndex;
    minR = Math.min(minR, r); maxR = Math.max(maxR, r);
    minC = Math.min(minC, c); maxC = Math.max(maxC, c);
  });

  const rows = maxR - minR + 1;
  const cols = maxC - minC + 1;
  const fullWidth = minC === 0 && maxC === totalCols - 1;
  const fullHeight = minR === 0 && maxR === totalRows - 1;

  if (fullWidth && fullHeight) return { kind: 'table', rows, cols };
  if (fullHeight) return { kind: 'column', rows, cols };
  if (fullWidth) return { kind: 'row', rows, cols };
  if (cells.length === 1) return { kind: 'cell', rows: 1, cols: 1 };
  return { kind: 'range', rows, cols };
}

/**
 * Przywraca domyślne obramowanie tabeli (jak po wstawieniu) — pełna, neutralna
 * siatka 1px. Zachowuje treść; nie rusza tła ani innych stylów komórek.
 */
export function restoreDefaultTableBorders(table: HTMLTableElement): void {
  cellsOf(table).forEach(cell => {
    ensureCellBox(cell);
    cell.style.border = `1px solid ${DEFAULT_CELL_BORDER_COLOR}`;
  });
}
