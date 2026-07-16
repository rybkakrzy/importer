/**
 * Logical table-grid model + <colgroup> synchronisation.
 *
 * The DOCX round-trip treats <colgroup> as the source of truth for w:tblGrid
 * (Word's column widths). Editor operations (column/table resize) mutate inline
 * cell widths; this module keeps the logical grid and the colgroup consistent so
 * a save does not return with stale geometry.
 *
 * Why a logical grid instead of td.cellIndex: a cell's position inside its <tr>
 * is NOT its column in Word's grid once the table has horizontal merges (colspan),
 * vertical merges (a rowspan leaves the <tr> below with fewer <td>s), or
 * gridBefore/gridAfter spacer cells (<td data-grid-spacer>). Operating by cellIndex
 * moves a different column in every row and corrupts merged/indented tables.
 * Everything here works in logical columns.
 */

/** A single rendered cell mapped onto the logical column grid. */
export interface TableGridCell {
  readonly el: HTMLTableCellElement;
  /** DOM row index of the cell's <tr>. */
  readonly rowIndex: number;
  /** Logical grid column of the cell's left edge. */
  readonly colStart: number;
  /** Number of logical columns the cell occupies (>= 1). */
  readonly colSpan: number;
  /** Number of rows the cell occupies (>= 1). */
  readonly rowSpan: number;
  /** True for gridBefore/gridAfter distance cells (no OOXML cell of their own). */
  readonly isSpacer: boolean;
}

export interface TableGrid {
  readonly columnCount: number;
  readonly cells: readonly TableGridCell[];
  readonly rows: readonly HTMLTableRowElement[];
}

/** The column immediately left of a dragged boundary, or the whole-table right edge. */
export type ColumnEdge = 'left' | 'right';

function clampSpan(value: number): number {
  return Number.isFinite(value) && value > 1 ? Math.floor(value) : 1;
}

/**
 * Builds the logical grid via a column-occupancy sweep. `occupancy[c]` counts how
 * many further rows column `c` is still covered by a rowspan started above, so the
 * next row skips those columns before placing its own cells (standard HTML table
 * layout, ECMA-376 vMerge maps to rowspan the same way).
 */
export function buildTableGrid(table: HTMLTableElement): TableGrid {
  const rows = Array.from(table.rows);
  const cells: TableGridCell[] = [];
  const occupancy: number[] = [];
  let columnCount = 0;

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    let col = 0;
    for (const el of Array.from(rows[rowIndex].cells)) {
      while ((occupancy[col] ?? 0) > 0) col++;

      const colSpan = clampSpan(el.colSpan);
      const rowSpan = clampSpan(el.rowSpan);
      cells.push({
        el,
        rowIndex,
        colStart: col,
        colSpan,
        rowSpan,
        isSpacer: el.hasAttribute('data-grid-spacer'),
      });

      for (let c = col; c < col + colSpan; c++) occupancy[c] = rowSpan;
      col += colSpan;
    }

    columnCount = Math.max(columnCount, col, occupancy.length);
    for (let c = 0; c < occupancy.length; c++) {
      if (occupancy[c] > 0) occupancy[c]--;
    }
  }

  return { columnCount, cells, rows };
}

export function findGridCell(grid: TableGrid, el: HTMLTableCellElement): TableGridCell | null {
  return grid.cells.find(c => c.el === el) ?? null;
}

/**
 * Logical column left of the boundary the pointer is dragging. Right edge → the
 * cell's last column; left edge → the column before its first. `null` means there
 * is no boundary (left edge of the very first column).
 */
export function columnBoundaryForCell(
  grid: TableGrid,
  el: HTMLTableCellElement,
  edge: ColumnEdge
): number | null {
  const cell = findGridCell(grid, el);
  if (!cell) return null;
  if (edge === 'right') return cell.colStart + cell.colSpan - 1;
  return cell.colStart > 0 ? cell.colStart - 1 : null;
}

function parsePx(value: string | null | undefined): number {
  if (!value) return 0;
  const match = /(-?\d+(?:\.\d+)?)px/.exec(value);
  return match ? Math.max(0, parseFloat(match[1])) : 0;
}

function colgroupCols(table: HTMLTableElement): HTMLElement[] {
  return Array.from(table.querySelectorAll(':scope > colgroup > col')) as HTMLElement[];
}

/**
 * Per-logical-column widths (px). The colgroup is authoritative (it round-trips
 * w:tblGrid); columns it does not size are measured from a single-span cell that
 * starts there. Columns whose width is unknown stay 0 and are left untouched by
 * {@link applyColumnWidths} so the layout is never forced to width:0.
 */
export function readColumnWidths(table: HTMLTableElement, grid: TableGrid): number[] {
  const widths = new Array<number>(grid.columnCount).fill(0);

  colgroupCols(table).forEach((col, i) => {
    if (i < widths.length) {
      const px = parsePx(col.style.width);
      if (px > 0) widths[i] = px;
    }
  });

  for (const cell of grid.cells) {
    if (cell.colSpan !== 1 || cell.isSpacer || widths[cell.colStart] > 0) continue;
    const measured = Math.round(cell.el.getBoundingClientRect().width);
    if (measured > 0) widths[cell.colStart] = measured;
  }

  return widths;
}

/**
 * Sets each cell's inline width to the sum of the logical columns it spans, keeping
 * every row aligned to the same grid. Cells over columns with unknown (0) width are
 * left as-is rather than collapsed.
 */
export function applyColumnWidths(table: HTMLTableElement, grid: TableGrid, widths: number[]): void {
  for (const cell of grid.cells) {
    let sum = 0;
    let known = true;
    for (let c = cell.colStart; c < cell.colStart + cell.colSpan; c++) {
      const w = widths[c] ?? 0;
      if (w > 0) sum += w;
      else {
        known = false;
        break;
      }
    }
    if (known && sum > 0) cell.el.style.width = `${Math.round(sum)}px`;
  }

  const total = widths.reduce((s, w) => s + (w > 0 ? w : 0), 0);
  if (total > 0) table.style.width = `${Math.round(total)}px`;
}

function ensureColgroup(table: HTMLTableElement, columnCount: number): HTMLElement {
  let colgroup = table.querySelector(':scope > colgroup') as HTMLElement | null;
  if (!colgroup) {
    colgroup = document.createElement('colgroup');
    table.insertBefore(colgroup, table.firstChild);
  }
  while (colgroup.children.length > columnCount) colgroup.removeChild(colgroup.lastChild!);
  while (colgroup.children.length < columnCount) colgroup.appendChild(document.createElement('col'));
  return colgroup;
}

/**
 * Writes resize results into the colgroup (the exported w:tblGrid). A column whose
 * width actually changed drops its exact data-w-tw (imported twips) and falls back
 * to px, so the grid does not drift; untouched columns keep their precise twips.
 * Works for fully-merged tables too, because widths come from the logical model
 * rather than measuring a "clean" row that may not exist.
 */
export function writeColgroupWidths(
  table: HTMLTableElement,
  grid: TableGrid,
  startWidths: number[],
  newWidths: number[]
): void {
  const cols = Array.from(ensureColgroup(table, grid.columnCount).children) as HTMLElement[];
  for (let i = 0; i < grid.columnCount; i++) {
    const col = cols[i];
    if (!col) continue;
    const nw = Math.round(newWidths[i] ?? 0);
    if (nw <= 0) continue;
    const changed = Math.round(startWidths[i] ?? 0) !== nw;
    if (changed) {
      col.style.width = `${nw}px`;
      col.removeAttribute('data-w-tw');
    } else if (parsePx(col.style.width) <= 0) {
      col.style.width = `${nw}px`;
    }
  }
}

/**
 * Rebuilds <colgroup> from the current layout after edits that changed the column
 * count (insert/delete column). Column count comes from the logical grid; widths of
 * a clean, unmerged reference row are measured. A measurement of 0 (jsdom, or a
 * table detached from layout) never overwrites an existing width.
 */
export function syncTableColgroup(table: HTMLTableElement): void {
  const grid = buildTableGrid(table);
  if (grid.columnCount <= 0) return;

  ensureColgroup(table, grid.columnCount);
  const cols = colgroupCols(table);

  // A reference row spans all columns with single-column, non-spacer cells; its
  // measured widths map 1:1 onto grid columns. Heavily merged tables have none, so
  // only the <col> count is corrected and existing widths are kept.
  const byRow = new Map<number, TableGridCell[]>();
  for (const cell of grid.cells) {
    const bucket = byRow.get(cell.rowIndex);
    if (bucket) bucket.push(cell);
    else byRow.set(cell.rowIndex, [cell]);
  }
  let reference: TableGridCell[] | null = null;
  for (const rowCells of byRow.values()) {
    if (
      rowCells.length === grid.columnCount &&
      rowCells.every(c => c.colSpan === 1 && !c.isSpacer)
    ) {
      reference = rowCells.slice().sort((a, b) => a.colStart - b.colStart);
      break;
    }
  }
  if (!reference) return;

  reference.forEach(cell => {
    const measured = Math.round(cell.el.getBoundingClientRect().width);
    if (measured <= 0) return;
    const col = cols[cell.colStart];
    if (!col) return;
    const newWidth = `${measured}px`;
    if (col.style.width !== newWidth) {
      col.style.width = newWidth;
      col.removeAttribute('data-w-tw');
    }
  });
}
