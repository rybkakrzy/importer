import {
  decodeShapeXml,
  encodeShapeXml,
  setAnchorPositionPageEmu,
  scaleShapeExtent,
  rescaleShapePreview,
  EMU_PER_PIXEL,
} from './shape-xml.util';

/**
 * ADR-0063: drag/resize kształtu pass-through musi aktualizować ORYGINALNY XML markera
 * (writer odtwarza grafikę z data-docx-xml — zmiany tylko w stylach GUI giną przy zapisie).
 */
describe('shape-xml.util', () => {
  const ANCHOR_XML =
    '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>111</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="714515" cy="714528"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    '<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="714515" cy="714528"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="714515" cy="714528"/></a:xfrm></wpg:grpSpPr>' +
    '</wpg:wgp></a:graphicData></a:graphic>' +
    '</wp:anchor></w:drawing>';

  function toB64(xml: string): string {
    return btoa(String.fromCharCode(...new TextEncoder().encode(xml)));
  }

  it('decode/encode: round-trip zachowuje treść (w tym znaki spoza ASCII)', () => {
    const xml = '<x name="żółć">zażółć</x>';
    const doc = decodeShapeXml(toB64(xml))!;
    expect(doc).not.toBeNull();
    const back = decodeShapeXml(encodeShapeXml(doc))!;
    expect(back.documentElement.getAttribute('name')).toBe('żółć');
    expect(back.documentElement.textContent).toBe('zażółć');
  });

  it('decode: uszkodzone base64/XML → null (marker zostaje nietknięty)', () => {
    expect(decodeShapeXml('!!!nie-base64!!!')).toBeNull();
    expect(decodeShapeXml(toB64('<niedomkniete'))).toBeNull();
  });

  it('setAnchorPositionPageEmu: normalizuje obie osie do page + posOffset (align znika)', () => {
    const doc = decodeShapeXml(toB64(ANCHOR_XML))!;

    expect(setAnchorPositionPageEmu(doc, 6130290, 9729977)).toBe(true);

    const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
    const posH = doc.getElementsByTagNameNS(wp, 'positionH')[0];
    const posV = doc.getElementsByTagNameNS(wp, 'positionV')[0];
    expect(posH.getAttribute('relativeFrom')).toBe('page');
    expect(posV.getAttribute('relativeFrom')).toBe('page');
    expect(posH.getElementsByTagNameNS(wp, 'posOffset')[0].textContent).toBe('6130290');
    expect(posV.getElementsByTagNameNS(wp, 'posOffset')[0].textContent).toBe('9729977');
    expect(posH.getElementsByTagNameNS(wp, 'align').length).toBe(0);
  });

  it('setAnchorPositionPageEmu: wp:inline → false (grafika w przepływie nie ma pozycji)', () => {
    const inline = ANCHOR_XML.replace(/wp:anchor/g, 'wp:inline');
    const doc = decodeShapeXml(toB64(inline))!;
    expect(setAnchorPositionPageEmu(doc, 1, 2)).toBe(false);
  });

  it('scaleShapeExtent: skaluje wp:extent i root a:ext, chExt NIETKNIĘTE (dzieci skalują się przez stosunek)', () => {
    const doc = decodeShapeXml(toB64(ANCHOR_XML))!;

    expect(scaleShapeExtent(doc, 2, 2)).toBe(true);

    const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const extent = doc.getElementsByTagNameNS(wp, 'extent')[0];
    expect(extent.getAttribute('cx')).toBe(String(714515 * 2));
    expect(extent.getAttribute('cy')).toBe(String(714528 * 2));
    const ext = doc.getElementsByTagNameNS(a, 'ext')[0];
    expect(ext.getAttribute('cx')).toBe(String(714515 * 2));
    const chExt = doc.getElementsByTagNameNS(a, 'chExt')[0];
    expect(chExt.getAttribute('cx')).toBe('714515');
  });

  it('rescaleShapePreview: kontener, dzieci i atrybuty SVG skalowane wspólnym faktorem', () => {
    const shape = document.createElement('div');
    shape.setAttribute('style', 'position:absolute;left:643px;top:927px;width:75px;height:75px;');
    shape.innerHTML =
      '<div style="position:absolute;left:24.42px;top:69.3px;width:7.03px;height:5.71px;">' +
      '<svg width="7.035" height="5.712" viewBox="0 0 67004 54405"></svg></div>';

    rescaleShapePreview(shape, 2, 2);

    expect(parseFloat(shape.style.width)).toBe(150);
    const child = shape.querySelector('div') as HTMLElement;
    expect(child.style.left).toBe('48.84px');
    expect(child.style.width).toBe('14.06px');
    const svg = shape.querySelector('svg')!;
    expect(svg.getAttribute('width')).toBe('14.07');
    expect(svg.getAttribute('viewBox')).toBe('0 0 67004 54405');
  });

  it('EMU_PER_PIXEL zgodny z resztą aplikacji', () => {
    expect(EMU_PER_PIXEL).toBe(9525);
  });
});
