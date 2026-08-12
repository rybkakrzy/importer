import { XML_HIGHLIGHT_LIMIT, tokenizeXml } from './xml-highlight.util';

/**
 * Tokenizer podglądu XML: kolorowanie nie może zgubić ANI JEDNEGO znaku (użytkownik czyta surowy
 * XML, nie jego interpretację) i musi zadziałać także dla XML, którego parser by nie przyjął —
 * bo właśnie takie dokumenty diagnozuje walidator struktury.
 */
describe('tokenizeXml', () => {
  function joined(xml: string): string {
    return tokenizeXml(xml)
      .map((token) => token.text)
      .join('');
  }

  it('zachowuje pełną treść wejściowego XML', () => {
    const xml = '<w:p><w:r w:rsidR="00A1"><w:t xml:space="preserve">Tekst &amp; więcej</w:t></w:r></w:p>';

    expect(joined(xml)).toBe(xml);
  });

  it('rozdziela nazwę elementu, atrybuty i wartości', () => {
    const tokens = tokenizeXml('<wp:anchor behindDoc="1"/>');

    expect(tokens.find((token) => token.kind === 'name')?.text).toBe('wp:anchor');
    expect(tokens.find((token) => token.kind === 'attribute')?.text).toBe('behindDoc');
    expect(tokens.find((token) => token.kind === 'value')?.text).toBe('"1"');
  });

  it('rozpoznaje komentarz jako jeden token', () => {
    const tokens = tokenizeXml('<a/><!-- uwaga --><b/>');

    expect(tokens.filter((token) => token.kind === 'comment').map((token) => token.text)).toEqual([
      '<!-- uwaga -->',
    ]);
  });

  it('nie gubi treści przy niedomkniętym znaczniku', () => {
    const xml = '<w:p><w:r';

    expect(joined(xml)).toBe(xml);
  });

  it('rezygnuje z kolorowania powyżej limitu rozmiaru', () => {
    const huge = `<a>${'x'.repeat(XML_HIGHLIGHT_LIMIT)}</a>`;

    const tokens = tokenizeXml(huge);

    expect(tokens).toHaveLength(1);
    expect(tokens[0].kind).toBe('text');
    expect(tokens[0].text).toBe(huge);
  });
});
