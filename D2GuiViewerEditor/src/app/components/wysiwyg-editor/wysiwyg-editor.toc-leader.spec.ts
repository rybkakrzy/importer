import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Spis treści jak w MS Word: wpis „tytuł …… numer strony" renderuje się jako linia flex,
 * w której tab z wypełniaczem (w:leader) to span.docx-tab-leader (inline flex:1 z readera),
 * a ZNAKI wypełniacza (kropki/podkreślenia) maluje pseudo-element ::after ze skompilowanego
 * SCSS komponentu — realne glify w foncie akapitu, przycięte do szerokości luki.
 *
 * jsdom nie layoutuje flexa ani pseudo-elementów, więc wizualny wygląd pinuje się w headless
 * Chrome (skill `verify`). Tu strzeżemy: (1) skompilowany CSS zawiera reguły malujące
 * wypełniacz, (2) paginacja nie tnie linii z wypełniaczem w środku (jednoliniowy flex).
 */
describe('WysiwygEditorComponent — wypełniacz tabulatora spisu treści (docx-tab-leader)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  function collectStyleText(): string {
    return Array.from(document.querySelectorAll('style'))
      .map((s) => s.textContent ?? '')
      .join('\n');
  }

  it('skompilowany CSS maluje kropki dla data-leader="dot" (i pozostałe wypełniacze)', () => {
    const css = collectStyleText();
    expect(css).toMatch(/docx-tab-leader\[data-leader=(?:'|")?dot(?:'|")?\]::after/);
    expect(css).toMatch(/docx-tab-leader\[data-leader=(?:'|")?underscore(?:'|")?\]::after/);
    expect(css).toMatch(/docx-tab-leader\[data-leader=(?:'|")?hyphen(?:'|")?\]::after/);
    // Zawartość kropkowa istnieje (content z ciągiem kropek).
    expect(css).toMatch(/content:\s*(?:'|")\.{10,}/);
  });

  it('nie tnie linii z wypełniaczem między stronami (_splitBlockAtBudget → null)', () => {
    const block = document.createElement('p');
    block.setAttribute('style', 'display:flex;align-items:baseline;');
    block.innerHTML =
      '<span class="docx-tab-text">Rozdział pierwszy</span>' +
      '<span class="docx-tab-leader" data-leader="dot" style="flex:1 1 0;">\t</span>' +
      '<span class="docx-tab-text">2</span>';
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
});
