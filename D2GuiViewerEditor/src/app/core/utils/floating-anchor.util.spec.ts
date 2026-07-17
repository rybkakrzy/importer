import {
  HfBandGeometry,
  bandToContract,
  computeAnchorBadgePosition,
  contractToBand,
  findAnchorParagraph,
  isFloatingElement,
  isPointerOnEdge,
  viewportDeltaToLayout,
} from './floating-anchor.util';

function rect(left: number, top: number, width: number, height: number): DOMRect {
  return {
    left, top, width, height,
    right: left + width,
    bottom: top + height,
    x: left, y: top,
    toJSON: () => ({}),
  } as DOMRect;
}

describe('floating-anchor.util — model kotwicy elementów pływających', () => {
  describe('isFloatingElement', () => {
    it('rozpoznaje data-pos-mode front/behind oraz position:absolute', () => {
      const el = document.createElement('div');
      expect(isFloatingElement(el)).toBe(false);

      el.dataset['posMode'] = 'front';
      expect(isFloatingElement(el)).toBe(true);

      el.dataset['posMode'] = 'behind';
      expect(isFloatingElement(el)).toBe(true);

      delete el.dataset['posMode'];
      el.style.position = 'absolute';
      expect(isFloatingElement(el)).toBe(true);
    });
  });

  describe('findAnchorParagraph', () => {
    it('obraz pływający → akapit zawierający (model: wrapper w <p>)', () => {
      const root = document.createElement('div');
      root.innerHTML = '<p id="a">tekst<span class="editor-image-wrapper"><img/></span></p><p id="b">inny</p>';
      const wrapper = root.querySelector('.editor-image-wrapper') as HTMLElement;

      const anchor = findAnchorParagraph(wrapper, root);

      expect(anchor?.id).toBe('a');
    });

    it('textbox → NASTĘPNY akapit (model: div hoistowany przed akapit-kotwicę)', () => {
      const root = document.createElement('div');
      root.innerHTML =
        '<p id="przed">przed</p>'
        + '<div class="docx-textbox" data-textbox="1"><p>treść pola</p></div>'
        + '<p id="kotwica">akapit kotwica</p>';
      const textbox = root.querySelector('.docx-textbox') as HTMLElement;

      const anchor = findAnchorParagraph(textbox, root);

      expect(anchor?.id).toBe('kotwica');
    });

    it('textbox pomija braci nie-akapity szukając kotwicy w przód', () => {
      const root = document.createElement('div');
      root.innerHTML =
        '<div class="docx-textbox"><p>pole</p></div>'
        + '<div class="docx-section-break"></div>'
        + '<p id="kotwica">akapit</p>';
      const textbox = root.querySelector('.docx-textbox') as HTMLElement;

      expect(findAnchorParagraph(textbox, root)?.id).toBe('kotwica');
    });

    it('textbox jako ostatni element → fallback do POPRZEDNIEGO akapitu', () => {
      const root = document.createElement('div');
      root.innerHTML = '<p id="poprzedni">tekst</p><div class="docx-textbox"><p>pole</p></div>';
      const textbox = root.querySelector('.docx-textbox') as HTMLElement;

      expect(findAnchorParagraph(textbox, root)?.id).toBe('poprzedni');
    });

    it('textbox w <li> (nie hoistowany) → element listy jako kotwica', () => {
      const root = document.createElement('div');
      root.innerHTML = '<ul><li id="li1">punkt<div class="docx-textbox"><p>pole</p></div></li></ul>';
      const textbox = root.querySelector('.docx-textbox') as HTMLElement;

      expect(findAnchorParagraph(textbox, root)?.id).toBe('li1');
    });

    it('zwraca null, gdy nie ma żadnego akapitu-kandydata', () => {
      const root = document.createElement('div');
      root.innerHTML = '<div class="docx-textbox"><p>pole</p></div>';
      const textbox = root.querySelector('.docx-textbox') as HTMLElement;

      expect(findAnchorParagraph(textbox, root)).toBeNull();
    });
  });

  describe('computeAnchorBadgePosition', () => {
    it('zoom 100%: pozycja = offset akapitu względem strony, ikona w lewym marginesie', () => {
      const pos = computeAnchorBadgePosition(rect(150, 300, 500, 20), rect(100, 100, 794, 1122), 794, 18);

      expect(pos.topPx).toBe(200);
      // 150-100=50 od lewej strony; 50 - 18 - 4 = 28
      expect(pos.leftPx).toBe(28);
    });

    it('zoom 200% (transform scale): współrzędne wracają do px układu strony', () => {
      // Strona 794 px layoutu renderowana jako 1588 px (scale=2); akapit 100 px od góry układu.
      const pos = computeAnchorBadgePosition(rect(300, 400, 1000, 40), rect(200, 200, 1588, 2244), 794, 18);

      expect(pos.topPx).toBe(100);   // (400-200)/2
      expect(pos.leftPx).toBe(28);   // (300-200)/2 - 22
    });

    it('nie wyjeżdża poza stronę (clamp od lewej/góry)', () => {
      const pos = computeAnchorBadgePosition(rect(0, -50, 100, 10), rect(0, 0, 794, 1122), 794, 18);

      expect(pos.leftPx).toBeGreaterThanOrEqual(2);
      expect(pos.topPx).toBeGreaterThanOrEqual(0);
    });
  });

  describe('viewportDeltaToLayout', () => {
    it('dzieli deltę przez skalę zoomu (drag pod zoomem nie ucieka kursorowi)', () => {
      expect(viewportDeltaToLayout(100, 2)).toBe(50);
      expect(viewportDeltaToLayout(100, 0.5)).toBe(200);
      expect(viewportDeltaToLayout(100, 1)).toBe(100);
    });

    it('skala 0/ujemna → delta bez zmian (ochrona przed dzieleniem przez zero)', () => {
      expect(viewportDeltaToLayout(100, 0)).toBe(100);
    });
  });

  describe('isPointerOnEdge', () => {
    const box = rect(100, 100, 200, 100);

    it('pas 8px przy krawędzi → true (uchwyt przenoszenia)', () => {
      expect(isPointerOnEdge(104, 150, box)).toBe(true);   // lewa
      expect(isPointerOnEdge(296, 150, box)).toBe(true);   // prawa
      expect(isPointerOnEdge(200, 104, box)).toBe(true);   // góra
      expect(isPointerOnEdge(200, 196, box)).toBe(true);   // dół
    });

    it('środek pola → false (klik zostaje dla edycji tekstu)', () => {
      expect(isPointerOnEdge(200, 150, box)).toBe(false);
    });

    it('poza polem → false', () => {
      expect(isPointerOnEdge(50, 50, box)).toBe(false);
    });
  });

  describe('contractToBand / bandToContract — kotwice w pasmach nagłówka/stopki', () => {
    // A4, marginesy 2.5cm (~94px), dystans nagłówka 1.25cm (~47px), pasmo 1.25cm (~47px).
    const headerGeo: HfBandGeometry = {
      band: 'header', marginLeftPx: 94, marginTopPx: 94, bandTopPx: 47,
    };
    // Stopka: strona 1122px, dystans 47px, pasmo 47px → góra pasma = 1122−47−47 = 1028.
    const footerGeo: HfBandGeometry = {
      band: 'footer', marginLeftPx: 94, marginTopPx: 94, bandTopPx: 1028,
    };

    it('header: kontrakt (od strony/obszaru treści) → układ pasma', () => {
      // Logo przy lewej krawędzi strony (x=20), u góry strony (y_page=30 → kontrakt y=30−94=−64).
      const { leftPx, topPx } = contractToBand(20, -64, headerGeo);
      expect(leftPx).toBe(20 - 94);     // ujemne — wystaje w lewy margines przed pasmo
      expect(topPx).toBe(-64 + 94 - 47); // = −17 → wystaje nad kontener pasma (jak w Wordzie)
    });

    it('footer: kontrakt → układ pasma (kotwica przy dole strony)', () => {
      // y_page = 1050 → kontrakt y = 1050−94 = 956.
      const { leftPx, topPx } = contractToBand(94, 956, footerGeo);
      expect(leftPx).toBe(0);
      expect(topPx).toBe(956 + 94 - 1028); // = 22 px od góry pasma stopki
    });

    it('bandToContract jest dokładną odwrotnością (round-trip bez dryfu)', () => {
      for (const geo of [headerGeo, footerGeo]) {
        const { leftPx, topPx } = contractToBand(123, 456, geo);
        const { xPx, yPx } = bandToContract(leftPx, topPx, geo);
        expect(xPx).toBe(123);
        expect(yPx).toBe(456);
      }
    });

    it('nie clampuje wartości ujemnych — obiekt może wystawać poza pasmo', () => {
      const { topPx } = contractToBand(0, -94, headerGeo);
      expect(topPx).toBeLessThan(0);
    });
  });
});
