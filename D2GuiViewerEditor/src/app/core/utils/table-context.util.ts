/**
 * Czysta klasyfikacja położenia selekcji DOM względem edytora i tabeli.
 * Wydzielona z `DocumentEditorComponent.detectTableContext`, by była testowalna
 * i by jednoznacznie rozróżnić trzy przypadki istotne dla cyklu życia aktywnej
 * tabeli:
 *  - `outside-editor` — selekcja poza treścią edytora (np. klik w panel boczny /
 *    toolbar / inne UI). Wywołujący NIE powinien czyścić aktywnej tabeli — to
 *    interakcja z konfiguracją, nie przejście w dokumencie.
 *  - `in-table`       — karetka w komórce tabeli → ustaw aktywną tabelę.
 *  - `outside-table`  — karetka w treści edytora, ale poza tabelą (np. inny akapit)
 *    → wyczyść aktywną tabelę i wizualne zaznaczenie.
 */
export type SelectionTablePlacement = 'outside-editor' | 'in-table' | 'outside-table';

export interface TableContextResult {
  placement: SelectionTablePlacement;
  cell: HTMLTableCellElement | null;
  table: HTMLTableElement | null;
}

export function resolveTableContext(
  anchorNode: Node | null | undefined,
  editorEl: HTMLElement | null | undefined
): TableContextResult {
  if (!anchorNode || !editorEl) {
    return { placement: 'outside-editor', cell: null, table: null };
  }
  const node = anchorNode instanceof HTMLElement ? anchorNode : anchorNode.parentElement;
  if (!node || !editorEl.contains(node)) {
    return { placement: 'outside-editor', cell: null, table: null };
  }
  const cell = node.closest('td, th') as HTMLTableCellElement | null;
  const table = node.closest('table') as HTMLTableElement | null;
  if (cell && table) {
    return { placement: 'in-table', cell, table };
  }
  return { placement: 'outside-table', cell: null, table: null };
}
