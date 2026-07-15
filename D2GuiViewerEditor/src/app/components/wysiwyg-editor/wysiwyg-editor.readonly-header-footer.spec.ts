import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regresja (bug 13677237, runda 2 — feedback QA): w dokumencie chronionym
 * (tryb tylko do odczytu) klik w pasmo nagłówka/stopki otwierał tryb edycji
 * (panel „Nagłówek i stopka", wstaw obraz / numery stron / usuń nagłówek),
 * a hover podświetlał pasma sugerując edytowalność.
 *
 * Kontrakt: przy `readOnly=true` startEditingHeader/Footer są no-op
 * (editingSection zostaje 'body', zero emisji editingSectionChange),
 * a wrapper dostaje klasę `read-only`, która wyłącza kursor i podświetlenie
 * hover pasm (reguły w wysiwyg-editor.scss).
 */
describe('WysiwygEditorComponent — read-only blokuje edycję nagłówka/stopki', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  it('startEditingHeader jest no-op w read-only (bez trybu edycji i bez emisji)', () => {
    component.readOnly = true;
    const emitted: string[] = [];
    component.editingSectionChange.subscribe((v) => emitted.push(v));

    component.startEditingHeader();

    expect(component.editingSection()).toBe('body');
    expect(emitted.length).toBe(0);
  });

  it('startEditingFooter jest no-op w read-only', () => {
    component.readOnly = true;
    const emitted: string[] = [];
    component.editingSectionChange.subscribe((v) => emitted.push(v));

    component.startEditingFooter();

    expect(component.editingSection()).toBe('body');
    expect(emitted.length).toBe(0);
  });

  it('bez read-only wejście w edycję nagłówka działa (kontrola)', () => {
    component.readOnly = false;

    component.startEditingHeader();

    expect(component.editingSection()).toBe('header');
  });

  it('klik w pasmo nagłówka w wyrenderowanym DOM nie otwiera edycji w read-only', () => {
    component.readOnly = true;
    fixture.detectChanges();

    const header = fixture.nativeElement.querySelector('.page-header') as HTMLElement;
    expect(header).toBeTruthy();
    header.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(component.editingSection()).toBe('body');
  });

  it('wrapper niesie klasę read-only sterowaną inputem (hook stylów hover)', () => {
    component.readOnly = true;
    fixture.detectChanges();
    const wrapper = fixture.nativeElement.querySelector('.wysiwyg-editor-wrapper') as HTMLElement;
    expect(wrapper.classList.contains('read-only')).toBe(true);

    component.readOnly = false;
    fixture.detectChanges();
    expect(wrapper.classList.contains('read-only')).toBe(false);
  });

  it('skompilowany CSS zawiera reguły wygaszające hover pasm w read-only', () => {
    fixture.detectChanges();
    const css = Array.from(document.querySelectorAll('style'))
      .map((s) => s.textContent ?? '')
      .join('\n');
    expect(css).toMatch(/read-only[^{}]*\.page-header/);
    expect(css).toMatch(/read-only[^{}]*\.page-footer:hover/);
  });
});
