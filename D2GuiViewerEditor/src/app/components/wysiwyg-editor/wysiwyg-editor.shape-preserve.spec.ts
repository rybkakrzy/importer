import { TestBed, ComponentFixture } from '@angular/core/testing';
import { ElementRef } from '@angular/core';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * ADR-0056: pass-through grafik XML. Kształty DrawingML/VML/OLE/grupy niosą oryginalny
 * OOXML w atrybucie data-docx-xml (base64; + części relacji w data-docx-rels) i są
 * renderowane jako niedotykalny podgląd (contenteditable=false z readera). GUI jest
 * warstwą passthrough: serializacja (getContent → writer) MUSI oddać markery 1:1 —
 * bez nich writer nie odtworzy grafiki i autosave trwale usunie ją z v2.
 *
 * Atrybut base64 jest odporny na normalizację innerHTML (w przeciwieństwie do inline
 * SVG podglądu), więc strzeżemy tu przejścia przez getContent i atomowości w paginacji.
 */
describe('WysiwygEditorComponent — pass-through grafik XML (ADR-0056)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    if (typeof (document as any).queryCommandState !== 'function') {
      (document as any).queryCommandState = () => false;
    }
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  const hostEls: HTMLElement[] = [];
  afterEach(() => {
    hostEls.splice(0).forEach(el => el.remove());
  });

  function editorsWith(...htmls: string[]): HTMLDivElement[] {
    const editors = htmls.map(html => {
      const editor = document.createElement('div');
      editor.innerHTML = html;
      document.body.appendChild(editor);
      hostEls.push(editor);
      return editor;
    });
    (component as any).pageEditorRefs = {
      toArray: () => editors.map(e => new ElementRef(e)),
    };
    return editors;
  }

  // Przykładowy marker (base64 "<w:drawing/>" — treść bez znaczenia dla GUI, kontrakt to opakowanie).
  const XML_MARKER = 'PHc6ZHJhd2luZy8+';

  it('getContent oddaje div.docx-shape z data-docx-xml i contenteditable=false 1:1', () => {
    (component as any)._captureDocumentDefaults('<p>x</p>');
    editorsWith(
      `<p>przed</p>` +
        `<div class="docx-shape docx-custgeom" data-shape="custom" contenteditable="false" ` +
        `data-docx-xml="${XML_MARKER}" style="width:100px;height:50px;">` +
        `<svg viewBox="0 0 100 50"><path d="M0 0 L100 0 Z" fill="#000066"/></svg></div>` +
        `<p>po</p>`
    );

    const content = component.getContent();
    expect(content).toContain(`data-docx-xml="${XML_MARKER}"`);
    expect(content).toContain('contenteditable="false"');
    expect(content).toContain('docx-custgeom');
  });

  it('getContent oddaje span.docx-preserved (niewidoczny placeholder) wraz z data-docx-rels', () => {
    (component as any)._captureDocumentDefaults('<p>x</p>');
    editorsWith(
      `<p>tekst<span class="docx-preserved" data-preserved="vml" contenteditable="false" ` +
        `data-docx-xml="${XML_MARKER}" data-docx-rels="eyJySWQxIjp7fX0=" ` +
        `style="display:inline-block;width:80px;height:40px;"></span>dalej</p>`
    );

    const content = component.getContent();
    expect(content).toContain('docx-preserved');
    expect(content).toContain(`data-docx-xml="${XML_MARKER}"`);
    expect(content).toContain('data-docx-rels="eyJySWQxIjp7fX0="');
  });

  it('paginacja nie tnie bloku z markerem pass-through (_splitBlockAtBudget → null)', () => {
    fixture.detectChanges();
    const block = document.createElement('p');
    block.innerHTML =
      `abc<span class="docx-preserved" data-docx-xml="${XML_MARKER}"></span>def`;
    const measurer = document.createElement('div');
    document.body.appendChild(measurer);
    try {
      const result = (component as unknown as {
        _splitBlockAtBudget(b: HTMLElement, budget: number, m: HTMLElement, lh: number): unknown;
      })._splitBlockAtBudget(block, 100, measurer, 10);
      expect(result).toBeNull();
    } finally {
      measurer.remove();
    }
  });

  it('skompilowany CSS chroni podgląd kształtu przed zapadnięciem (max-width:none)', () => {
    fixture.detectChanges();
    const css = Array.from(document.querySelectorAll('style'))
      .map(s => s.textContent ?? '')
      .join('\n');
    // Ta sama klasa buga co .docx-tab-seg img: procentowy max-width w pozycjonowanym
    // kontenerze zapada obraz/SVG do ~0px.
    expect(css).toMatch(/docx-shape[^}]*\bsvg\b[^}]*/);
    expect(css).toMatch(/max-width:\s*none/);
  });
});
