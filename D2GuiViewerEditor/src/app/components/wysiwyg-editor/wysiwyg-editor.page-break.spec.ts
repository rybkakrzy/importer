import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Item 8 — a manual page break must be a SEMANTIC, non-printing marker, never a
 * graphic in the content. These tests pin the two guarantees we control here:
 *  - insertPageBreak emits a bare `div.page-break` with NO inline graphic styling
 *    (no border/dashed line/margin) — the old code hardcoded a dashed border.
 *  - the visible hint is opt-in via the "formatting marks" mode, off by default.
 *
 * The OOXML export guarantee (div.page-break → w:br type=page, not a drawing) is
 * covered by the backend converter tests; here we only assert the editor never
 * injects a graphic.
 */
describe('WysiwygEditorComponent — page break is a semantic marker (item 8)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  it('inserts a bare page-break marker with no inline graphic styling', () => {
    let inserted = '';
    (component as any).insertHtml = (html: string) => {
      inserted = html;
    };

    component.insertPageBreak();

    expect(inserted).toContain('class="page-break"');
    // Regression: the marker must NOT carry a visible graphic (border/dashed/margin).
    expect(inserted).not.toMatch(/border/i);
    expect(inserted).not.toMatch(/dashed/i);
    expect(inserted).not.toMatch(/margin/i);
    // It should stay non-editable so it cannot be focused like an image.
    expect(inserted).toContain('contenteditable="false"');
  });

  it('hides the break hint by default and reveals it only in formatting-marks mode', () => {
    expect(component.showFormattingMarks()).toBe(false);

    component.toggleFormattingMarks();
    expect(component.showFormattingMarks()).toBe(true);

    component.toggleFormattingMarks(false);
    expect(component.showFormattingMarks()).toBe(false);
  });
});
