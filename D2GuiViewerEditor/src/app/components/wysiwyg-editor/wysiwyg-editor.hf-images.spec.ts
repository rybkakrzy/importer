import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';
import { contractToBand } from '../../core/utils/floating-anchor.util';

/**
 * Zakotwiczone (pływające, wp:anchor) obrazy w NAGŁÓWKU/STOPCE — pozycjonowanie jak w MS Word.
 * Kontrakt data-x/y-emu: X od lewej krawędzi STRONY, Y od góry OBSZARU TREŚCI (y_page − marginTop);
 * kontener pasma ma inny origin (padding-left = margines, offset od góry strony), więc:
 *  - tryb WYŚWIETLANIA (.header-display, [innerHTML]) dostaje inline absolut przeliczony
 *    z kontraktu (_positionBandAnchors) — wcześniej goły <img> renderował się w flow,
 *  - tryb EDYCJI przelicza styl wrappera (contractToBand), a data-emu zostają w kontrakcie,
 *  - commit pasma zdejmuje edycyjne wrappery (model czysty jak w body).
 */
describe('WysiwygEditorComponent — obrazy kotwiczone w nagłówku/stopce', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  const EMU_PER_PX = 9525;
  const IMG =
    `<img src="data:image/png;base64,x" data-pos-mode="front" ` +
    `data-x-emu="${20 * EMU_PER_PX}" data-y-emu="${-30 * EMU_PER_PX}" ` +
    `style="width:120px;height:40px;">`;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [WysiwygEditorComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  function positionBandAnchors(html: string, band: 'header' | 'footer'): string {
    return (component as unknown as {
      _positionBandAnchors(h: string, i: number, b: 'header' | 'footer'): string;
    })._positionBandAnchors(html, 0, band);
  }

  function bandGeoFor(band: 'header' | 'footer') {
    return (component as unknown as {
      _bandGeoFor(i: number, b: 'header' | 'footer'): import('../../core/utils/floating-anchor.util').HfBandGeometry;
    })._bandGeoFor(0, band);
  }

  it('display: obraz kotwiczony w nagłówku dostaje inline absolut przeliczony z kontraktu', () => {
    const out = positionBandAnchors(`<p>Tekst${IMG}</p>`, 'header');

    const tpl = document.createElement('template');
    tpl.innerHTML = out;
    const img = tpl.content.querySelector('img') as HTMLImageElement;
    expect(img.style.position).toBe('absolute');

    const geo = bandGeoFor('header');
    const expected = contractToBand(20, -30, geo);
    expect(parseInt(img.style.left, 10)).toBe(Math.round(expected.leftPx));
    expect(parseInt(img.style.top, 10)).toBe(Math.round(expected.topPx));
    // front = nad tekstem.
    expect(img.style.zIndex).toBe('10');
    // Kontrakt na data-* nietknięty (transformacja jest jednokierunkowa, tylko prezentacja).
    expect(img.getAttribute('data-x-emu')).toBe(String(20 * EMU_PER_PX));
    expect(img.getAttribute('data-y-emu')).toBe(String(-30 * EMU_PER_PX));
  });

  it('display: behind dostaje z-index -1 (pod tekstem, jak w Wordzie)', () => {
    const html = IMG.replace('data-pos-mode="front"', 'data-pos-mode="behind"');
    const out = positionBandAnchors(`<p>${html}</p>`, 'header');
    const tpl = document.createElement('template');
    tpl.innerHTML = out;
    expect((tpl.content.querySelector('img') as HTMLImageElement).style.zIndex).toBe('-1');
  });

  it('display: HTML bez obrazów kotwiczonych przechodzi nietknięty (zero kosztu)', () => {
    const html = '<p>Zwykły nagłówek <img src="x" style="width:10px;"></p>';
    expect(positionBandAnchors(html, 'header')).toBe(html);
  });

  it('display: kotwiczony kształt (docx-shape) w stopce dostaje left/top przeliczone z kontraktu', () => {
    // Reader emituje inline absolut we współrzędnych KONTRAKTU (X od lewej krawędzi strony,
    // Y od góry obszaru treści) — bez przeliczenia origin pasma kształt ląduje poza kartką
    // (zgłoszenie: logo-kwadrat w stopce z left:675/top:906 renderowany daleko pod stroną).
    const shape =
      '<div class="docx-shape docx-custgeom" data-shape="custom" ' +
      'style="position:absolute;left:675px;top:906px;z-index:1;width:75px;height:75px;">' +
      '<svg viewBox="0 0 10 10"><path d="M0 0 L10 0 Z"></path></svg></div>';
    const out = positionBandAnchors(`<p>Stopka</p>${shape}`, 'footer');

    const tpl = document.createElement('template');
    tpl.innerHTML = out;
    const el = tpl.content.querySelector('.docx-shape') as HTMLElement;

    const expected = contractToBand(675, 906, bandGeoFor('footer'));
    expect(parseInt(el.style.left, 10)).toBe(Math.round(expected.leftPx));
    expect(parseInt(el.style.top, 10)).toBe(Math.round(expected.topPx));
    // Rozmiar i zawartość SVG nietknięte.
    expect(el.style.width).toBe('75px');
    expect(el.querySelector('svg')).not.toBeNull();
  });

  it('display: statyczny kształt (inline-block, bez kotwicy) przechodzi nietknięty', () => {
    const html =
      '<p>Stopka</p><div class="docx-shape docx-line" data-shape="line" ' +
      'style="display:inline-block;width:100px;height:2px;background:#000;"></div>';
    expect(positionBandAnchors(html, 'footer')).toBe(html);
  });

  it('geometria pasm: stopka liczy górę pasma od dołu strony, nagłówek od dystansu', () => {
    const header = bandGeoFor('header');
    const footer = bandGeoFor('footer');
    expect(header.bandTopPx).toBe(component.headerOffsetPx(0));
    expect(footer.bandTopPx).toBeCloseTo(
      component.pageHeightPx(0) - component.footerOffsetPx(0) - component.footerBandPx(0), 3);
    // Marginesy wspólne z bindingami pasma.
    expect(header.marginLeftPx).toBe(component.pageMarginPx(0, 'left'));
    expect(header.marginTopPx).toBe(component.pageMarginPx(0, 'top'));
  });

  it('edycja: wrapExistingImages w paśmie ustawia styl w układzie pasma, a data-emu zostają kontraktem', () => {
    const bandEl = document.createElement('div');
    bandEl.className = 'header-editor-content';
    bandEl.innerHTML = `<p>${IMG}</p>`;
    document.body.appendChild(bandEl);
    try {
      (component as unknown as { wrapExistingImages(c: HTMLElement): void }).wrapExistingImages(bandEl);

      const wrapper = bandEl.querySelector('.editor-image-wrapper') as HTMLElement;
      const img = bandEl.querySelector('img') as HTMLImageElement;
      expect(wrapper).not.toBeNull();
      expect(wrapper.style.position).toBe('absolute');

      const expected = contractToBand(20, -30, bandGeoFor('header'));
      expect(parseInt(wrapper.style.left, 10)).toBe(Math.round(expected.leftPx));
      expect(parseInt(wrapper.style.top, 10)).toBe(Math.round(expected.topPx));
      // Kontrakt bez zmian — zapis DOCX nie może przesunąć obrazu o origin pasma.
      expect(img.getAttribute('data-x-emu')).toBe(String(20 * EMU_PER_PX));
      expect(img.getAttribute('data-y-emu')).toBe(String(-30 * EMU_PER_PX));
    } finally {
      bandEl.remove();
    }
  });

  it('edycja w BODY: wrapExistingImages zachowuje dotychczasowy kontrakt 1:1 (regresja)', () => {
    const bodyEl = document.createElement('div');
    bodyEl.className = 'editor-content';
    bodyEl.innerHTML = `<p>${IMG}</p>`;
    document.body.appendChild(bodyEl);
    try {
      (component as unknown as { wrapExistingImages(c: HTMLElement): void }).wrapExistingImages(bodyEl);
      const wrapper = bodyEl.querySelector('.editor-image-wrapper') as HTMLElement;
      expect(parseInt(wrapper.style.left, 10)).toBe(20);
      expect(parseInt(wrapper.style.top, 10)).toBe(-30);
    } finally {
      bodyEl.remove();
    }
  });

  it('commit pasma: model dostaje czysty HTML bez wrapperów, <img> zachowuje data-emu', () => {
    component.headerContent = { html: '<p>H</p>', height: 1.25 };

    // Surowy innerHTML edycji: wrapper z uchwytami i edycyjnym absolutem w układzie pasma.
    const edited =
      '<p>Tekst<span class="editor-image-wrapper" contenteditable="false" draggable="true" ' +
      'data-pos-mode="front" style="position:absolute;left:5px;top:7px;">' +
      IMG +
      '<span class="image-resize-handle resize-handle-right"></span></span></p>';
    component.onHeaderInput({ target: { innerHTML: edited } } as unknown as Event);

    const saved = (component as unknown as { _headerHtml(): string })._headerHtml();
    expect(saved).not.toContain('editor-image-wrapper');
    expect(saved).not.toContain('image-resize-handle');
    expect(saved).not.toContain('contenteditable');
    expect(saved).toContain(`data-x-emu="${20 * EMU_PER_PX}"`);
    expect(saved).toContain(`data-y-emu="${-30 * EMU_PER_PX}"`);
    // Edycyjny absolut wrappera nie może zostać na <img> — display pozycjonuje z kontraktu.
    expect(saved).not.toContain('position:absolute');
  });
});
