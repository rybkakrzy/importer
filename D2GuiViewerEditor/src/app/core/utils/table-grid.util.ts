/**
 * Synchronizacja <colgroup> z rzeczywistą siatką tabeli.
 *
 * <colgroup> jest dla eksportu HTML→DOCX źródłem prawdy o w:tblGrid (szerokości kolumn
 * Worda). Operacje edytora (resize kolumny/tabeli, wstawienie/usunięcie kolumny) zmieniają
 * inline'owe szerokości komórek, ale nie aktualizowały colgroup — po zapisie dokument
 * wracał ze STARĄ geometrią kolumn. Ten util odtwarza colgroup z aktualnego układu.
 */

/**
 * Przebudowuje <colgroup> tabeli na podstawie bieżącego układu:
 * - liczba <col> = liczba kolumn siatki (max sumy colSpan po wierszach),
 * - szerokości mierzone z pierwszego wiersza bez scaleń poziomych (colSpan=1 w każdej komórce).
 * Gdy taki wiersz nie istnieje (mocno scalona tabela) — dopasowywana jest tylko liczba <col>,
 * istniejące szerokości zostają (lepsza stara geometria niż zgadywana).
 * Pomiar 0 (np. jsdom / tabela odpięta od layoutu) nie nadpisuje istniejącej szerokości.
 */
export function syncTableColgroup(table: HTMLTableElement): void {
  const rows = Array.from(table.rows);
  if (rows.length === 0) return;

  const gridCount = Math.max(
    ...rows.map(r => Array.from(r.cells).reduce((s, c) => s + (c.colSpan || 1), 0))
  );
  if (gridCount <= 0) return;

  let colgroup = table.querySelector(':scope > colgroup') as HTMLTableColElement | null;
  if (!colgroup) {
    colgroup = document.createElement('colgroup') as HTMLTableColElement;
    table.insertBefore(colgroup, table.firstChild);
  }

  // Dopasuj liczbę <col> do siatki.
  while (colgroup.children.length > gridCount) colgroup.removeChild(colgroup.lastChild!);
  while (colgroup.children.length < gridCount) colgroup.appendChild(document.createElement('col'));

  const baseRow = rows.find(
    r =>
      r.cells.length === gridCount &&
      Array.from(r.cells).every(c => (c.colSpan || 1) === 1)
  );
  if (!baseRow) return;

  Array.from(baseRow.cells).forEach((cell, i) => {
    const w = Math.round(cell.getBoundingClientRect().width);
    if (w > 0) (colgroup!.children[i] as HTMLElement).style.width = `${w}px`;
  });
}
