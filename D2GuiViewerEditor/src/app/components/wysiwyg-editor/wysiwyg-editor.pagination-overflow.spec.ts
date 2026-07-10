import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Regresja: „ENTER przy dolnej krawędzi strony rozciągał kartkę zamiast utworzyć nową stronę".
 *
 * Root cause: bloki były mierzone w GOŁYM <div> poza kontekstem stylów `.editor-content`
 * (pusty <p> = 0 px, brak reguły `p:empty::before`), a `getBoundingClientRect().height`
 * nie zawiera marginesów bloków — suma pomiarów zaniżała realną wysokość treści, więc
 * `_repaginateNow` nie widziała przepełnienia i wszystkie bloki zostawały na jednej stronie
 * (`.page` ma tylko min-height → rosła w pion). Dodatkowo czysty debounce 250 ms resetował
 * się na każdym `input`, więc PRZYTRZYMANY Enter odsuwał repaginację w nieskończoność.
 *
 * Te testy pinują: (1) measurer w kontekście `.editor-content`, (2) pomiar wysokości
 * z marginesami (delty pozycji), (3) tworzenie nowej strony i przeniesienie przepełnionego
 * bloku (w tym pustego akapitu) bez utraty/duplikacji treści, (4) max-wait debounce'a.
 */
describe('WysiwygEditorComponent — przepełnienie strony tworzy nową stronę (stała wysokość kartki)', () => {
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

  function pageWith(html: string): HTMLDivElement {
    const editor = document.createElement('div');
    editor.innerHTML = html;
    document.body.appendChild(editor);
    bodyEditors.push(editor);
    (component as any).pageEditorRefs = { toArray: () => [{ nativeElement: editor }] };
    return editor;
  }

  /** Stub pomiaru: stała wysokość na blok (jsdom nie ma layoutu). */
  function stubBlockHeights(heightPerBlock: number): void {
    (component as any)._measureBlockRunHeights = (_m: HTMLElement, blocks: HTMLElement[]) =>
      blocks.map(() => heightPerBlock);
  }

  describe('measurer paginacji', () => {
    it('renderuje bloki w kontekście stylów strony (klasa editor-content, bez własnych paddingów)', () => {
      const cs = getComputedStyle(document.body);
      const measurer = (component as any)._createBlockMeasurer(cs, 700) as HTMLElement;

      expect(measurer.classList.contains('editor-content')).toBe(true);
      expect(measurer.style.padding).toBe('0px');
      expect(measurer.style.position).toBe('absolute');
      expect(measurer.style.width).toBe('700px');
    });

    it('_measureBlockRunHeights liczy wysokość konsumowaną Z marginesami (delty pozycji, nie rect.height)', () => {
      // Syntetyczny layout pionowy: top elementu = suma (data-h + data-mb + data-mt)
      // poprzednich rodzeństw + własny data-mt. Symuluje przeglądarkowy układ blokowy.
      const orig = Element.prototype.getBoundingClientRect;
      const num = (el: Element | null, a: string) =>
        el ? parseFloat(el.getAttribute(a) ?? '0') || 0 : 0;
      Element.prototype.getBoundingClientRect = function (this: Element): DOMRect {
        let top = 0;
        let sib = this.previousElementSibling;
        while (sib) {
          top += num(sib, 'data-h') + num(sib, 'data-mb') + num(sib, 'data-mt');
          sib = sib.previousElementSibling;
        }
        top += num(this, 'data-mt');
        const height = num(this, 'data-h');
        return {
          top, height, bottom: top + height,
          left: 0, right: 0, width: 0, x: 0, y: top,
          toJSON: () => ({}),
        } as DOMRect;
      };
      try {
        const measurer = document.createElement('div');
        document.body.appendChild(measurer);
        bodyEditors.push(measurer);

        const p1 = document.createElement('p');
        p1.setAttribute('data-h', '19');
        p1.setAttribute('data-mb', '10');
        p1.textContent = 'a';
        const p2 = document.createElement('p');
        p2.setAttribute('data-h', '19');
        p2.setAttribute('data-mb', '10');
        p2.textContent = 'b';

        const heights = (component as any)._measureBlockRunHeights(measurer, [p1, p2]) as number[];

        // 19 px treści + 10 px margin-bottom = 29 px konsumowane na stronie.
        // Stary pomiar (rect.height) zwracał 19 — kartka rosła o 10 px na akapit.
        expect(heights).toEqual([29, 29]);
      } finally {
        Element.prototype.getBoundingClientRect = orig;
      }
    });
  });

  describe('_repaginateNow — podział przy przepełnieniu', () => {
    it('tworzy nową stronę, gdy bloki nie mieszczą się na jednej (strona nie rośnie w pion)', () => {
      // 5 bloków × 700 px przy dostępnej wysokości A4 (~930 px) → dokładnie 1 blok na stronę.
      pageWith('<p>1</p><p>2</p><p>3</p><p>4</p><p>5</p>');
      stubBlockHeights(700);

      (component as any)._repaginateNow();

      const pages = component.pageContents();
      expect(pages.length).toBe(5);
      pages.forEach((html, i) => expect(html).toContain(`<p>${i + 1}</p>`));
    });

    it('zachowuje kolejność i nie duplikuje/nie gubi treści przy podziale', () => {
      pageWith('<p>alfa</p><p>beta</p><p>gamma</p><p>delta</p>');
      stubBlockHeights(700);

      (component as any)._repaginateNow();

      const joined = component.pageContents().join('');
      expect(joined).toBe('<p>alfa</p><p>beta</p><p>gamma</p><p>delta</p>');
    });

    it('przenosi przepełniony PUSTY akapit (po ENTER) na nową stronę zamiast rozciągać bieżącą', () => {
      pageWith('<p>tekst na dole strony</p><p></p>');
      stubBlockHeights(700);

      (component as any)._repaginateNow();

      const pages = component.pageContents();
      expect(pages.length).toBe(2);
      expect(pages[0]).toBe('<p>tekst na dole strony</p>');
      expect(pages[1]).toBe('<p></p>');
    });

    it('nie tworzy nadmiarowych stron, gdy treść się mieści', () => {
      pageWith('<p>a</p><p>b</p><p>c</p>');
      stubBlockHeights(10);

      (component as any)._repaginateNow();

      expect(component.pageContents().length).toBe(1);
    });

    it('po usunięciu bloków treść wraca na jedną stronę (repaginacja scala w górę)', () => {
      pageWith('<p>zostaje</p>');
      stubBlockHeights(10);
      // Symulacja stanu po wcześniejszym podziale na 2 strony:
      component.pageContents.set(['<p>zostaje</p>', '<p></p>']);

      (component as any)._repaginateNow();

      expect(component.pageContents().length).toBe(1);
    });
  });

  describe('_schedulePaginate — max-wait debounce', () => {
    it('repaginuje najpóźniej po 600 ms mimo ciągłego strumienia input (przytrzymany ENTER)', () => {
      vi.useFakeTimers();
      const spy = vi.fn();
      (component as any)._repaginateNow = spy;

      // Zdarzenia co 100 ms przez 1 s — czysty debounce 250 ms nigdy by nie odpalił.
      for (let t = 0; t < 10; t++) {
        (component as any)._schedulePaginate('held-enter');
        vi.advanceTimersByTime(100);
      }

      expect(spy.mock.calls.length).toBeGreaterThanOrEqual(1);
    });

    it('pojedyncza zmiana repaginuje po zwykłym debounce 250 ms', () => {
      vi.useFakeTimers();
      const spy = vi.fn();
      (component as any)._repaginateNow = spy;

      (component as any)._schedulePaginate('single');
      vi.advanceTimersByTime(249);
      expect(spy).not.toHaveBeenCalled();
      vi.advanceTimersByTime(1);
      expect(spy).toHaveBeenCalledTimes(1);
    });
  });

  /**
   * Regresja: „ENTER na dole strony rozciąga kartkę na ~250 ms (debounce), a nowa strona
   * pojawia się dopiero po ręcznym scrollu". Fix: ENTER wymusza repaginację NATYCHMIAST
   * (rAF po wstawieniu akapitu, z pominięciem debounce'a), a `_restoreGlobalCaret` przewija
   * kursor do widoku — więc przelana treść od razu widać bez scrollowania.
   */
  describe('ENTER — natychmiastowa repaginacja + przewinięcie kursora do widoku', () => {
    it('_flushPaginateNow repaginuje od razu i kasuje oczekujący debounce', () => {
      vi.useFakeTimers();
      const spy = vi.fn();
      (component as any)._repaginateNow = spy;

      (component as any)._schedulePaginate('input'); // ustawia timer debounce'a
      (component as any)._flushPaginateNow();

      expect(spy).toHaveBeenCalledTimes(1);
      expect((component as any)._paginateTimer).toBeNull();

      // Skasowany debounce nie odpali repaginacji drugi raz.
      vi.advanceTimersByTime(600);
      expect(spy).toHaveBeenCalledTimes(1);
    });

    it('ENTER (keydown) wymusza natychmiastowy flush paginacji', () => {
      const flush = vi.fn();
      (component as any)._flushPaginateSoon = flush;

      (component as any).handleKeyboard(new KeyboardEvent('keydown', { key: 'Enter' }));

      expect(flush).toHaveBeenCalledTimes(1);
    });

    it('Ctrl+Enter NIE wymusza flushu (nie koliduje ze skrótami)', () => {
      const flush = vi.fn();
      (component as any)._flushPaginateSoon = flush;

      (component as any).handleKeyboard(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true }));

      expect(flush).not.toHaveBeenCalled();
    });

    it('_flushPaginateSoon koalescuje wiele ENTER w jedną repaginację na klatkę', () => {
      const rafCbs: FrameRequestCallback[] = [];
      const rafSpy = vi.spyOn(window, 'requestAnimationFrame')
        .mockImplementation((cb: FrameRequestCallback) => { rafCbs.push(cb); return rafCbs.length; });
      const flushNow = vi.fn();
      (component as any)._flushPaginateNow = flushNow;

      (component as any)._flushPaginateSoon();
      (component as any)._flushPaginateSoon();
      (component as any)._flushPaginateSoon();

      expect(rafSpy).toHaveBeenCalledTimes(1); // trzy ENTER → jeden rAF
      rafCbs[0](0);
      expect(flushNow).toHaveBeenCalledTimes(1);

      rafSpy.mockRestore();
    });

    it('_restoreGlobalCaret przewija kursor do widoku (scrollIntoView block:nearest)', () => {
      const editor = pageWith('<p>alfa</p><p>beta</p>');
      const scrollSpy = vi.fn();
      (Array.from(editor.children) as HTMLElement[]).forEach(b => { b.scrollIntoView = scrollSpy; });

      // Kursor w 2. bloku (treść przelana na kolejną stronę po ENTER).
      (component as any)._restoreGlobalCaret({ block: 1, offset: 0 });

      expect(scrollSpy).toHaveBeenCalledWith({ block: 'nearest', inline: 'nearest' });
    });
  });

  /**
   * Regresja: „w trakcie pisania kursor na ułamek sekundy skacze na początek dokumentu,
   * a niektóre znaki lądują w złym miejscu".
   *
   * Root cause: sygnał pageContents jest celowo przestarzały między repaginacjami
   * (onPageInput go nie aktualizuje), więc stary check `identical` (porównanie nowego
   * rozkładu Z SYGNAŁEM) był przy pisaniu zawsze false — każda repaginacja (max-wait 600 ms)
   * rebindowała [innerHTML] wszystkich stron mimo że żaden blok nie zmienił strony.
   * Podmiana DOM kasowała selekcję (karetka spadała na początek contenteditable), a znaki
   * wpisane w oknie rebind→restore lądowały w złym miejscu lub ginęły ze starym DOM.
   * Fix: porównanie z ŻYWYM DOM — bez zmiany rozkładu bloków zero rebindów.
   */
  describe('_repaginateNow — brak rebindu stron, gdy rozkład bloków się nie zmienił', () => {
    it('pisanie bez przelania strony NIE wymienia DOM (pageContents nietknięte)', () => {
      const editor = pageWith('<p>alfa</p><p>beta</p>');
      stubBlockHeights(10); // wszystko mieści się na jednej stronie
      component.pageContents.set(['<p>alfa</p><p>beta</p>']);
      const setSpy = vi.spyOn(component.pageContents, 'set');

      // Symulacja pisania: żywy DOM odjeżdża od sygnału, ale bloki zostają na stronie.
      (editor.children[0] as HTMLElement).textContent = 'alfaX';
      (component as any)._repaginateNow();

      expect(setSpy).not.toHaveBeenCalled();
      // Sygnał celowo przestarzały (DOM jest źródłem prawdy do czasu realnego przelania).
      expect(component.pageContents()).toEqual(['<p>alfa</p><p>beta</p>']);
    });

    it('przelanie bloku na nową stronę nadal rebinduje strony', () => {
      const editor = pageWith('<p>alfa</p><p>beta</p>');
      stubBlockHeights(700); // 2 bloki nie mieszczą się na jednej stronie
      component.pageContents.set(['<p>alfa</p><p>beta</p>']);

      (editor.children[0] as HTMLElement).textContent = 'alfaX';
      (component as any)._repaginateNow();

      expect(component.pageContents()).toEqual(['<p>alfaX</p>', '<p>beta</p>']);
    });

    it('opróżnienie strony (usunięcie całej treści) rebinduje syntetyczny pusty akapit', () => {
      pageWith(''); // użytkownik skasował wszystko — contenteditable bez bloków
      stubBlockHeights(10);
      component.pageContents.set(['<p>stare</p>']);

      (component as any)._repaginateNow();

      expect(component.pageContents()).toEqual(['<p></p>']);
    });

    it('rozjazd liczby stron sygnału z DOM wymusza rebind (scalanie w górę po usunięciu)', () => {
      pageWith('<p>zostaje</p>');
      stubBlockHeights(10);
      component.pageContents.set(['<p>zostaje</p>', '<p></p>']);

      (component as any)._repaginateNow();

      expect(component.pageContents()).toEqual(['<p>zostaje</p>']);
    });
  });
});
