import { TestBed, ComponentFixture } from '@angular/core/testing';
import { EditorToolbarComponent } from './editor-toolbar';
import { FontProviderService } from '../../services/font-provider.service';
import { EditorState } from '../../models/document.model';

/**
 * Item 6 — the font control must never blank on focus, must be typeable and
 * searchable, must not overwrite the document font until confirmed, and must
 * show a mixed state. Item 7 — the main toolbar draws its list from the shared
 * FontProviderService (so the corporate font is available here too).
 */

class StubFontProvider extends FontProviderService {
  raw = "'CorporateSans', 'Calibri', 'Segoe UI', Arial, sans-serif";
  protected override readCorporateFontRaw(): string {
    return this.raw;
  }
}

function fakeInput(value = ''): HTMLInputElement {
  const el = document.createElement('input');
  el.value = value;
  el.select = () => {};
  el.blur = () => {};
  return el;
}

function stateWithFont(fontFamily: string, fontMixed = false): EditorState {
  return {
    isModified: false,
    canUndo: false,
    canRedo: false,
    wordCount: 0,
    fontMixed,
    currentFormatting: {
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      subscript: false,
      superscript: false,
    },
    currentStyle: { fontFamily },
  };
}

describe('EditorToolbarComponent — font combobox (items 6 & 7)', () => {
  let fixture: ComponentFixture<EditorToolbarComponent>;
  let component: EditorToolbarComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [EditorToolbarComponent],
      providers: [{ provide: FontProviderService, useClass: StubFontProvider }],
    }).compileComponents();
    fixture = TestBed.createComponent(EditorToolbarComponent);
    component = fixture.componentInstance;
  });

  it('lists the corporate font from the shared provider (also in the mini-toolbar)', () => {
    const provider = TestBed.inject(FontProviderService);
    (provider as StubFontProvider).refresh();
    expect(component.fontFamilies()).toContain('CorporateSans');
    // The mini-toolbar binds the very same signal, so the font is shared.
    expect(component.fontFamilies()).toBe(provider.displayNames());
  });

  it('clears the field on focus so the datalist shows the full list', () => {
    component.editorState = stateWithFont('Arial');
    expect(component.selectedFontFamily()).toBe('Arial');
    expect(component.fontInputValue()).toBe('Arial');

    const input = fakeInput('Arial');
    component.onFontFocus({ target: input } as unknown as FocusEvent);
    // The visible text is cleared: otherwise the native datalist filters its
    // options down to the current font and no other font can be picked.
    expect(input.value).toBe('');
    // The effective font (the signal) is preserved for restore on blur.
    expect(component.fontInputValue()).toBe('Arial');
  });

  it('restores the current font on blur when nothing was picked', () => {
    component.editorState = stateWithFont('Arial');
    const input = fakeInput('Arial');
    component.onFontFocus({ target: input } as unknown as FocusEvent);
    expect(input.value).toBe('');

    // User clicks away without choosing anything.
    component.onFontBlur({ target: input } as unknown as FocusEvent);
    expect(input.value).toBe('Arial');
    expect(component.selectedFontFamily()).toBe('Arial');
  });

  it('commits a typed font on Enter and emits it normalised', () => {
    let emitted: string | null = null;
    component.fontFamilyChange.subscribe((v) => (emitted = v));

    const input = fakeInput('times new roman');
    component.onFontKeydown({
      key: 'Enter',
      preventDefault: () => {},
      target: input,
    } as unknown as KeyboardEvent);

    expect(emitted).toBe('Times New Roman');
    expect(component.selectedFontFamily()).toBe('Times New Roman');
    expect(input.value).toBe('Times New Roman');
  });

  it('does not overwrite the font on an empty commit', () => {
    component.editorState = stateWithFont('Arial');
    let emitted = false;
    component.fontFamilyChange.subscribe(() => (emitted = true));

    const input = fakeInput('   ');
    component.onFontCommit({ target: input } as unknown as Event);

    expect(emitted).toBe(false);
    expect(component.selectedFontFamily()).toBe('Arial');
    expect(input.value).toBe('Arial'); // restored
  });

  it('restores the current font on Escape without emitting', () => {
    component.editorState = stateWithFont('Arial');
    let emitted = false;
    component.fontFamilyChange.subscribe(() => (emitted = true));

    const input = fakeInput('Comic Sans MS'); // user typed but cancels
    component.onFontKeydown({
      key: 'Escape',
      preventDefault: () => {},
      target: input,
    } as unknown as KeyboardEvent);

    expect(emitted).toBe(false);
    expect(input.value).toBe('Arial');
    expect(component.selectedFontFamily()).toBe('Arial');
  });

  it('shows a blank (mixed) state for a multi-font selection', () => {
    component.editorState = stateWithFont('Arial', true);
    expect(component.fontMixed()).toBe(true);
    expect(component.fontInputValue()).toBe('');
  });

  // ── Własny dropdown (max-height + scroll) i placeholder = aktywny krój ──

  it('placeholder pokazuje AKTYWNY krój (nie statyczne „Czcionka"); mieszana selekcja → „—"', () => {
    component.editorState = stateWithFont('Arial');
    expect(component.fontPlaceholder()).toBe('Arial');

    component.editorState = stateWithFont('Arial', true);
    expect(component.fontPlaceholder()).toBe('—');
  });

  it('focus otwiera dropdown z PEŁNĄ listą (filtr wyzerowany), blur zamyka', () => {
    component.editorState = stateWithFont('Arial');
    const input = fakeInput('Arial');

    component.onFontFocus({ target: input } as unknown as FocusEvent);
    expect(component.fontDropdownOpen()).toBe(true);
    expect(component.filteredFonts()).toBe(component.fontFamilies());

    component.onFontBlur({ target: input } as unknown as FocusEvent);
    expect(component.fontDropdownOpen()).toBe(false);
  });

  it('wpisywanie filtruje listę (case-insensitive substring)', () => {
    (TestBed.inject(FontProviderService) as StubFontProvider).refresh();
    const input = fakeInput('');
    component.onFontFocus({ target: input } as unknown as FocusEvent);

    input.value = 'corp';
    component.onFontInput({ target: input } as unknown as Event);

    expect(component.filteredFonts()).toEqual(['CorporateSans']);
  });

  it('mousedown na opcji commituje czcionkę i zamyka dropdown (przed blur inputa)', () => {
    let emitted: string | null = null;
    component.fontFamilyChange.subscribe((v) => (emitted = v as string));
    const input = fakeInput('');
    component.onFontFocus({ target: input } as unknown as FocusEvent);

    let prevented = false;
    component.onFontOptionMouseDown({
      preventDefault: () => (prevented = true),
      target: document.createElement('button'),
    } as unknown as MouseEvent, 'CorporateSans');

    expect(prevented).toBe(true);
    expect(component.fontDropdownOpen()).toBe(false);
    expect(emitted!).toBe('CorporateSans');
    expect(component.selectedFontFamily()).toBe('CorporateSans');
  });

  it('Enter z filtrem pasującym do jednej pozycji wybiera ją (jak klik)', () => {
    (TestBed.inject(FontProviderService) as StubFontProvider).refresh();
    let emitted: string | null = null;
    component.fontFamilyChange.subscribe((v) => (emitted = v as string));
    const input = fakeInput('');
    component.onFontFocus({ target: input } as unknown as FocusEvent);
    input.value = 'corp';
    component.onFontInput({ target: input } as unknown as Event);

    component.onFontKeydown({
      key: 'Enter',
      preventDefault: () => {},
      target: input,
    } as unknown as KeyboardEvent);

    expect(emitted!).toBe('CorporateSans');
    expect(component.fontDropdownOpen()).toBe(false);
  });
});
