import { describe, it, expect } from 'vitest';
import {
  applyListLabels,
  ensureBulletMarkers,
  formatListNumber,
  stripListLabelAttributes,
} from './list-label.util';

/**
 * Silnik etykiet list DOCX — odpowiednik liczników readera (ADR-0036 wariant A).
 * Scenariusze lustrzane do backendowych `ListNumberingFidelityTests`: wspólna
 * semantyka liczenia musi obowiązywać w podglądzie i w eksporcie.
 */
describe('list-label.util — formatListNumber', () => {
  it('formatuje decimal / decimalZero', () => {
    expect(formatListNumber('decimal', 7)).toBe('7');
    expect(formatListNumber('decimalZero', 5)).toBe('05');
    expect(formatListNumber('decimalZero', 12)).toBe('12');
  });

  it('formatuje litery jak Word (po z POWTÓRZONA litera: 27=aa, 53=aaa)', () => {
    expect(formatListNumber('lowerLetter', 1)).toBe('a');
    expect(formatListNumber('lowerLetter', 26)).toBe('z');
    expect(formatListNumber('lowerLetter', 27)).toBe('aa');
    expect(formatListNumber('lowerLetter', 53)).toBe('aaa');
    expect(formatListNumber('upperLetter', 2)).toBe('B');
  });

  it('formatuje liczby rzymskie', () => {
    expect(formatListNumber('upperRoman', 4)).toBe('IV');
    expect(formatListNumber('upperRoman', 9)).toBe('IX');
    expect(formatListNumber('upperRoman', 1999)).toBe('MCMXCIX');
    expect(formatListNumber('lowerRoman', 3)).toBe('iii');
  });
});

function labelsOf(root: HTMLElement): Array<string | null> {
  return Array.from(root.querySelectorAll('li')).map(li => li.getAttribute('data-list-label'));
}

function build(html: string): HTMLElement {
  const root = document.createElement('div');
  root.innerHTML = html;
  return root;
}

describe('list-label.util — applyListLabels', () => {
  it('prosta lista numerowana: 1. 2. 3.', () => {
    const root = build(
      '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1.">' +
        '<li>a</li><li>b</li><li>c</li></ol>',
    );
    applyListLabels([root]);
    expect(labelsOf(root)).toEqual(['1.', '2.', '3.']);
  });

  it('szablon wielopoziomowy %1.%2) z formatem rodzica i restartem po powrocie', () => {
    const nested =
      '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="1" data-num-fmt="lowerLetter" data-lvl-text="%1.%2)">' +
      '<li>x</li><li>y</li></ol>';
    const root = build(
      '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1.">' +
        `<li>jeden${nested}</li><li>dwa${nested}</li></ol>`,
    );
    applyListLabels([root]);
    // Kolejność dokumentu: 1., 1.a), 1.b), 2., 2.a), 2.b) — %1 formatowane decimal
    // (format POZIOMU odwołania, nie bieżącego), głębszy poziom restartuje po powrocie.
    expect(labelsOf(root)).toEqual(['1.', '1.a)', '1.b)', '2.', '2.a)', '2.b)']);
  });

  it('kontynuacja między fragmentami tej samej listy (wspólny data-num-id) i przez strony', () => {
    const frag = (items: string) =>
      '<ol data-num-id="7" data-abstract-num-id="3" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1.">' +
      `${items}</ol>`;
    const page1 = build(frag('<li>a</li><li>b</li>') + '<p>przerwa</p>');
    const page2 = build(frag('<li>c</li>'));
    applyListLabels([page1, page2]);
    expect(labelsOf(page1)).toEqual(['1.', '2.']);
    expect(labelsOf(page2)).toEqual(['3.']);
  });

  it('wspólny abstrakt bez override kontynuuje; data-start-override restartuje', () => {
    const ol = (numId: string, extra = '') =>
      `<ol data-num-id="${numId}" data-abstract-num-id="5" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1."${extra}>` +
      '<li>x</li></ol>';
    const continued = build(ol('1') + ol('2'));
    applyListLabels([continued]);
    expect(labelsOf(continued)).toEqual(['1.', '2.']);

    const restarted = build(ol('1') + ol('2', ' data-start-override="1"'));
    applyListLabels([restarted]);
    expect(labelsOf(restarted)).toEqual(['1.', '1.']);
  });

  it('data-start ustawia wartość początkową definicji', () => {
    const root = build(
      '<ol data-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1." data-start="5">' +
        '<li>a</li><li>b</li></ol>',
    );
    applyListLabels([root]);
    expect(labelsOf(root)).toEqual(['5.', '6.']);
  });

  it('isLegal formatuje wszystkie odwołania jako decimal', () => {
    const nested =
      '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="1" data-num-fmt="decimal" data-lvl-text="%1.%2." data-is-legal="1">' +
      '<li>x</li></ol>';
    const root = build(
      '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="0" data-num-fmt="upperRoman" data-lvl-text="%1.">' +
        `<li>jeden${nested}</li></ol>`,
    );
    applyListLabels([root]);
    // Bez isLgl byłoby "I.1." — legal wymusza "1.1.".
    expect(labelsOf(root)).toEqual(['I.', '1.1.']);
  });

  it('lvlRestart=0: poziom głębszy NIE restartuje po powrocie na płytszy', () => {
    const nested = (items: string) =>
      '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="1" data-num-fmt="decimal" data-lvl-text="%2." data-lvl-restart="0">' +
      `${items}</ol>`;
    const root = build(
      '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1.">' +
        `<li>jeden${nested('<li>x</li>')}</li><li>dwa${nested('<li>y</li>')}</li></ol>`,
    );
    applyListLabels([root]);
    expect(labelsOf(root)).toEqual(['1.', '1.', '2.', '2.']);
  });

  it('punktory nie dostają etykiety, ale restartują głębsze poziomy numerowane', () => {
    const nested = (items: string) =>
      '<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="1" data-num-fmt="decimal" data-lvl-text="%2.">' +
      `${items}</ol>`;
    const root = build(
      '<ul data-num-id="1" data-abstract-num-id="1" data-ilvl="0" data-num-fmt="bullet" data-lvl-text="•">' +
        `<li>punkt${nested('<li>x</li><li>y</li>')}</li><li>punkt${nested('<li>z</li>')}</li></ul>`,
    );
    applyListLabels([root]);
    const labels = labelsOf(root);
    expect(labels[0]).toBeNull();
    expect(labels[1]).toBe('1.');
    expect(labels[2]).toBe('2.');
    expect(labels[3]).toBeNull();
    expect(labels[4]).toBe('1.');
  });

  it('suffix ≠ tab jedzie do data-list-suffix; listy bez data-num-id nietykane', () => {
    const root = build(
      '<ol data-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1." data-suffix="nothing"><li>a</li></ol>' +
        '<ol><li>edytorowa</li></ol>',
    );
    applyListLabels([root]);
    const items = Array.from(root.querySelectorAll('li'));
    expect(items[0].getAttribute('data-list-suffix')).toBe('nothing');
    expect(items[1].hasAttribute('data-list-label')).toBe(false);
  });

  it('przeliczenie jest idempotentne i aktualizuje etykiety po edycji DOM', () => {
    const root = build(
      '<ol data-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1.">' +
        '<li>a</li><li>b</li></ol>',
    );
    applyListLabels([root]);
    // Symulacja edycji: nowy li na początku listy.
    const ol = root.querySelector('ol')!;
    const li = document.createElement('li');
    li.textContent = 'nowy';
    ol.insertBefore(li, ol.firstChild);
    applyListLabels([root]);
    expect(labelsOf(root)).toEqual(['1.', '2.', '3.']);
  });

  it('schematy wielopoziomowe z wymagań: 1.1.1. / 1.1.1) / A.1.a. / I.A.1. / § 1.1.1', () => {
    const scheme = (fmts: [string, string, string], tpls: [string, string, string]) =>
      build(
        `<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="0" data-num-fmt="${fmts[0]}" data-lvl-text="${tpls[0]}"><li>a` +
          `<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="1" data-num-fmt="${fmts[1]}" data-lvl-text="${tpls[1]}"><li>b` +
          `<ol data-num-id="1" data-abstract-num-id="1" data-ilvl="2" data-num-fmt="${fmts[2]}" data-lvl-text="${tpls[2]}"><li>c</li></ol></li></ol></li></ol>`,
      );

    const cases: Array<{
      fmts: [string, string, string];
      tpls: [string, string, string];
      expected: string[];
    }> = [
      {
        fmts: ['decimal', 'decimal', 'decimal'],
        tpls: ['%1.', '%1.%2.', '%1.%2.%3.'],
        expected: ['1.', '1.1.', '1.1.1.'],
      },
      {
        fmts: ['decimal', 'decimal', 'decimal'],
        tpls: ['%1)', '%1.%2)', '%1.%2.%3)'],
        expected: ['1)', '1.1)', '1.1.1)'],
      },
      {
        fmts: ['upperLetter', 'decimal', 'lowerLetter'],
        tpls: ['%1.', '%1.%2.', '%1.%2.%3.'],
        expected: ['A.', 'A.1.', 'A.1.a.'],
      },
      {
        fmts: ['upperRoman', 'upperLetter', 'decimal'],
        tpls: ['%1.', '%1.%2.', '%1.%2.%3.'],
        expected: ['I.', 'I.A.', 'I.A.1.'],
      },
      {
        fmts: ['decimal', 'decimal', 'decimal'],
        tpls: ['§ %1', '§ %1.%2', '§ %1.%2.%3'],
        expected: ['§ 1', '§ 1.1', '§ 1.1.1'],
      },
    ];

    for (const c of cases) {
      const root = scheme(c.fmts, c.tpls);
      applyListLabels([root]);
      expect(labelsOf(root)).toEqual(c.expected);
    }
  });

  it('numeratory z literałami w szablonie: (1) [1] Krok 1 Pkt 1', () => {
    const cases: Array<{ tpl: string; expected: string[] }> = [
      { tpl: '(%1)', expected: ['(1)', '(2)'] },
      { tpl: '[%1]', expected: ['[1]', '[2]'] },
      { tpl: 'Krok %1', expected: ['Krok 1', 'Krok 2'] },
      { tpl: 'Pkt %1', expected: ['Pkt 1', 'Pkt 2'] },
    ];
    for (const c of cases) {
      const root = build(
        `<ol data-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="${c.tpl}">` +
          '<li>a</li><li>b</li></ol>',
      );
      applyListLabels([root]);
      expect(labelsOf(root)).toEqual(c.expected);
    }
  });

  it('ensureBulletMarkers uzupełnia marker span w li utworzonym Enterem', () => {
    // Drugi li = świeżo utworzony Enterem (Chrome nie klonuje marker spana).
    const root = build(
      '<ul data-num-id="1" data-ilvl="0" data-num-fmt="bullet" data-lvl-text="TODO:">' +
        '<li><span class="list-marker" style="min-width:1.2em;">TODO:</span>pierwszy</li>' +
        '<li>nowy po Enterze</li></ul>',
    );
    ensureBulletMarkers([root]);

    const items = Array.from(root.querySelectorAll('li'));
    const marker = items[1].querySelector(':scope > span.list-marker');
    expect(marker).not.toBeNull();
    expect(marker!.textContent).toBe('TODO:');
    expect(marker).toBe(items[1].firstChild);
    // Istniejący marker nietknięty (bez duplikatów).
    expect(items[0].querySelectorAll(':scope > span.list-marker').length).toBe(1);

    // Idempotencja: kolejny przebieg niczego nie dokłada.
    ensureBulletMarkers([root]);
    expect(items[1].querySelectorAll(':scope > span.list-marker').length).toBe(1);
  });

  it('ensureBulletMarkers nie rusza list bez marker spanów ani list spoza DOCX', () => {
    const root = build(
      '<ol data-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1."><li>a</li><li>b</li></ol>' +
        '<ul><li>edytorowa</li></ul>',
    );
    ensureBulletMarkers([root]);
    expect(root.querySelectorAll('span.list-marker').length).toBe(0);
  });

  it('stripListLabelAttributes zdejmuje atrybuty prezentacyjne', () => {
    const root = build(
      '<ol data-num-id="1" data-ilvl="0" data-num-fmt="decimal" data-lvl-text="%1." data-suffix="space"><li>a</li></ol>',
    );
    applyListLabels([root]);
    expect(root.querySelectorAll('[data-list-label]').length).toBe(1);
    stripListLabelAttributes(root);
    expect(root.querySelectorAll('[data-list-label]').length).toBe(0);
    expect(root.querySelectorAll('[data-list-suffix]').length).toBe(0);
    // Kontrakt writera (data-num-id itd.) zostaje nietknięty.
    expect(root.querySelector('ol')!.getAttribute('data-num-id')).toBe('1');
  });
});
