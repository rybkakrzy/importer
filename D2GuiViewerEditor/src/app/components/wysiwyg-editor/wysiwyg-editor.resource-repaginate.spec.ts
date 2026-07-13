import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regresja (UAT): „ten sam dokument ma inną liczbę stron w DOC2 Viewer i w MS Word".
 *
 * Jedną z potwierdzonych przyczyn rozjazdu jest liczenie paginacji na metrykach dostępnych
 * ZANIM doładują się web-fonty i obrazy: pierwszy pomiar bloków korzysta z czcionki fallback
 * (inne wznoszenie/opadanie/interlinia), więc liczba stron może odbiegać od finalnego układu.
 * Fix: jednorazowa KOREKCYJNA repaginacja po `document.fonts.ready` + doładowaniu obrazów.
 *
 * Te testy pinują kontrakt tej korekty: (1) po ustabilizowaniu zasobów leci DOKŁADNIE jeden
 * przebieg, (2) świeższy import anuluje oczekiwanie starszego (token generacji), (3) zniszczony
 * komponent nie repaginuje, (4) oczekiwanie na obraz kończy się na jego zdarzeniu load.
 */
describe('WysiwygEditorComponent — korekcyjna repaginacja po załadowaniu zasobów', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;
  const bodyEditors: HTMLElement[] = [];

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  afterEach(() => {
    bodyEditors.splice(0).forEach(el => el.remove());
    vi.useRealTimers();
  });

  function setPageEditor(html: string): HTMLDivElement {
    const editor = document.createElement('div');
    editor.innerHTML = html;
    document.body.appendChild(editor);
    bodyEditors.push(editor);
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: editor }] };
    return editor;
  }

  /** Przepuszcza kolejkę mikrozadań (łańcuch Promise.all w _repaginateAfterResources). */
  const flushMicrotasks = () => new Promise(resolve => setTimeout(resolve, 0));

  it('po ustabilizowaniu zasobów wykonuje DOKŁADNIE jeden przebieg repaginacji', async () => {
    setPageEditor('<p>Treść</p>');
    const flush = vi.fn();
    (component as any)._flushPaginateNow = flush;

    (component as any)._repaginateAfterResources();
    expect(flush).not.toHaveBeenCalled(); // czeka na zasoby, nie leci synchronicznie

    await flushMicrotasks();
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it('nowszy import anuluje oczekiwanie starszego (token generacji) — jeden przebieg', async () => {
    setPageEditor('<p>Treść</p>');
    const flush = vi.fn();
    (component as any)._flushPaginateNow = flush;

    (component as any)._repaginateAfterResources(); // gen 1
    (component as any)._repaginateAfterResources(); // gen 2 — unieważnia gen 1

    await flushMicrotasks();
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it('zniszczony komponent nie repaginuje po doładowaniu zasobów', async () => {
    setPageEditor('<p>Treść</p>');
    const flush = vi.fn();
    (component as any)._flushPaginateNow = flush;

    (component as any)._repaginateAfterResources();
    (component as any)._isDestroyed = true;

    await flushMicrotasks();
    expect(flush).not.toHaveBeenCalled();
  });

  it('_pendingImagesSettled czeka na obraz jeszcze-ładujący się i kończy na load', async () => {
    const editor = setPageEditor('<p><img alt="x"></p>');
    const img = editor.querySelector('img') as HTMLImageElement;
    // jsdom nie odpala realnego ładowania — wymuś stan „w toku".
    Object.defineProperty(img, 'complete', { value: false, configurable: true });

    let settled = false;
    (component as any)._pendingImagesSettled().then(() => (settled = true));

    await flushMicrotasks();
    expect(settled).toBe(false); // wciąż czeka na obraz

    img.dispatchEvent(new Event('load'));
    await flushMicrotasks();
    expect(settled).toBe(true);
  });

  it('_pendingImagesSettled rozwiązuje się od razu, gdy brak ładujących się obrazów', async () => {
    setPageEditor('<p>Bez obrazów</p>');
    let settled = false;
    (component as any)._pendingImagesSettled().then(() => (settled = true));

    await flushMicrotasks();
    expect(settled).toBe(true);
  });
});
