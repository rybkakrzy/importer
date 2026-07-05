import { describe, expect, it } from 'vitest';
import { syncTableColgroup } from './table-grid.util';

/**
 * Synchronizacja <colgroup> z siatką tabeli — po operacjach edytora (wstawienie/usunięcie
 * kolumny, resize) colgroup jest źródłem prawdy dla eksportowanego w:tblGrid.
 * W jsdom getBoundingClientRect zwraca 0, więc testujemy strukturę (liczbę <col>
 * i zachowanie istniejących szerokości), nie pomiary.
 */
describe('syncTableColgroup', () => {
  function makeTable(html: string): HTMLTableElement {
    const host = document.createElement('div');
    host.innerHTML = html;
    return host.querySelector('table') as HTMLTableElement;
  }

  it('dokłada brakujące <col> po wstawieniu kolumny', () => {
    const table = makeTable(
      '<table><colgroup><col style="width:100px;"><col style="width:200px;"></colgroup>' +
        '<tr><td>a</td><td>b</td><td>NOWA</td></tr><tr><td>c</td><td>d</td><td>e</td></tr></table>'
    );
    syncTableColgroup(table);
    const cols = table.querySelectorAll('colgroup > col');
    expect(cols.length).toBe(3);
    // Pomiar 0 (jsdom) nie może wyzerować istniejących szerokości.
    expect((cols[0] as HTMLElement).style.width).toBe('100px');
    expect((cols[1] as HTMLElement).style.width).toBe('200px');
  });

  it('usuwa nadmiarowe <col> po usunięciu kolumny', () => {
    const table = makeTable(
      '<table><colgroup><col><col><col></colgroup>' +
        '<tr><td>a</td><td>b</td></tr></table>'
    );
    syncTableColgroup(table);
    expect(table.querySelectorAll('colgroup > col').length).toBe(2);
  });

  it('tworzy colgroup, gdy tabela go nie miała', () => {
    const table = makeTable('<table><tr><td>a</td><td>b</td></tr></table>');
    syncTableColgroup(table);
    const colgroup = table.querySelector(':scope > colgroup');
    expect(colgroup).not.toBeNull();
    expect(colgroup!.children.length).toBe(2);
    expect(table.firstElementChild).toBe(colgroup);
  });

  it('liczy kolumny siatki z uwzględnieniem colspan', () => {
    const table = makeTable(
      '<table><tr><td colspan="2">ab</td><td>c</td></tr><tr><td>a</td><td>b</td><td>c</td></tr></table>'
    );
    syncTableColgroup(table);
    expect(table.querySelectorAll('colgroup > col').length).toBe(3);
  });

  it('mocno scalona tabela (bez wiersza 1:1) zachowuje istniejące szerokości', () => {
    const table = makeTable(
      '<table><colgroup><col style="width:80px;"><col style="width:80px;"></colgroup>' +
        '<tr><td colspan="2">ab</td></tr></table>'
    );
    syncTableColgroup(table);
    const cols = table.querySelectorAll('colgroup > col');
    expect(cols.length).toBe(2);
    expect((cols[0] as HTMLElement).style.width).toBe('80px');
  });
});
