import { resolveTableContext } from './table-context.util';

function buildEditorWithTable(): { editor: HTMLElement; cellText: Node; paragraphText: Node } {
  const editor = document.createElement('div');
  const table = document.createElement('table');
  const tr = table.insertRow();
  const td = tr.insertCell();
  td.textContent = 'hello';
  editor.appendChild(table);

  const p = document.createElement('p');
  p.textContent = 'outside';
  editor.appendChild(p);

  document.body.appendChild(editor);
  return { editor, cellText: td.firstChild!, paragraphText: p.firstChild! };
}

describe('resolveTableContext', () => {
  it('treats a null anchor (no DOM selection) as outside-editor — context is preserved by caller', () => {
    const editor = document.createElement('div');
    const r = resolveTableContext(null, editor);
    expect(r.placement).toBe('outside-editor');
    expect(r.table).toBeNull();
  });

  it('treats an anchor outside the editor (e.g. side panel button) as outside-editor', () => {
    const { editor } = buildEditorWithTable();
    const panelButton = document.createElement('button');
    document.body.appendChild(panelButton);

    const r = resolveTableContext(panelButton, editor);
    // This is the core fix: clicking the panel must NOT clear the active table.
    expect(r.placement).toBe('outside-editor');
  });

  it('detects a caret inside a table cell as in-table and returns the cell + table', () => {
    const { editor, cellText } = buildEditorWithTable();
    const r = resolveTableContext(cellText, editor);
    expect(r.placement).toBe('in-table');
    expect(r.cell?.tagName).toBe('TD');
    expect(r.table?.tagName).toBe('TABLE');
  });

  it('detects a caret in editor content but outside any table as outside-table', () => {
    const { editor, paragraphText } = buildEditorWithTable();
    const r = resolveTableContext(paragraphText, editor);
    expect(r.placement).toBe('outside-table');
    expect(r.table).toBeNull();
  });

  it('returns outside-editor when there is no editor element', () => {
    expect(resolveTableContext(document.createElement('span'), null).placement).toBe('outside-editor');
  });
});
