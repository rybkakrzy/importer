import { RulerComponent, RulerColumnSegment } from './ruler';
import { CSS_PX_PER_CM } from '../../core/utils/units.util';

/**
 * Linijka per kolumna (jak MS Word): przy kursorze w sekcji wielokolumnowej pozioma
 * linijka pokazuje szare strefy na marginesy strony ORAZ odstępy między kolumnami,
 * a uchwyty wcięć są pozycjonowane i ograniczane względem AKTYWNEJ kolumny.
 */
describe('RulerComponent — segmenty kolumn', () => {
  const PX = CSS_PX_PER_CM;

  /** A4 pion, marginesy 2 cm, 2 kolumny po 7.9 cm z odstępem 1.2 cm. */
  const twoColumns: RulerColumnSegment[] = [
    { startCm: 2, widthCm: 7.9 },
    { startCm: 11.1, widthCm: 7.9 }
  ];

  function createRuler(overrides?: Partial<RulerComponent>): RulerComponent {
    const r = new RulerComponent();
    r.mode = 'horizontal';
    r.margins = { top: 2, bottom: 2, left: 2, right: 2 };
    r.blockIndent = { start: 0, end: 0 };
    r.columnSegments = twoColumns;
    r.activeColumnIndex = 0;
    Object.assign(r, overrides);
    return r;
  }

  // ---------- pozycje uchwytów ----------

  it('uchwyty w 1. kolumnie: start na krawędzi kolumny, end na jej prawej krawędzi', () => {
    const r = createRuler();
    expect(r.startMarginPx).toBeCloseTo(2 * PX, 1);
    expect(r.endMarginPx).toBeCloseTo((21 - 9.9) * PX, 1); // 21 − (2 + 7.9)
  });

  it('uchwyty w 2. kolumnie: start = startCm 2. segmentu', () => {
    const r = createRuler({ activeColumnIndex: 1 });
    expect(r.startMarginPx).toBeCloseTo(11.1 * PX, 1);
    expect(r.endMarginPx).toBeCloseTo((21 - 19) * PX, 1); // 21 − (11.1 + 7.9)
  });

  it('wcięcie bloku odkłada się od krawędzi aktywnej kolumny, nie strony', () => {
    const r = createRuler({ activeColumnIndex: 1, blockIndent: { start: 1, end: 0.5 } });
    expect(r.startMarginPx).toBeCloseTo((11.1 + 1) * PX, 1);
    expect(r.endMarginPx).toBeCloseTo((21 - 19 + 0.5) * PX, 1);
  });

  it('null / 1 segment → zachowanie jak dotychczas (marginesy strony)', () => {
    const single = createRuler({ columnSegments: [{ startCm: 2, widthCm: 17 }] });
    expect(single.activeSegment).toBeNull();
    expect(single.startMarginPx).toBeCloseTo(2 * PX, 1);

    const none = createRuler({ columnSegments: null });
    expect(none.activeSegment).toBeNull();
    expect(none.endMarginPx).toBeCloseTo(2 * PX, 1);
  });

  it('activeColumnIndex poza zakresem jest przycinany do ostatniego segmentu', () => {
    const r = createRuler({ activeColumnIndex: 7 });
    expect(r.activeSegment!.start).toBeCloseTo(11.1, 3);
  });

  // ---------- szare strefy ----------

  it('hGrayZones: marginesy strony + odstęp między kolumnami (geometria, bez wcięcia)', () => {
    const r = createRuler({ blockIndent: { start: 1.5, end: 0 } });
    const zones = r.hGrayZones;
    expect(zones.length).toBe(3);
    // lewy margines strony
    expect(zones[0].left).toBeCloseTo(0, 1);
    expect(zones[0].width).toBeCloseTo(2 * PX, 1);
    // odstęp między kolumnami (9.9 → 11.1 cm) — wcięcie bloku NIE przesuwa stref
    expect(zones[1].left).toBeCloseTo(9.9 * PX, 1);
    expect(zones[1].width).toBeCloseTo(1.2 * PX, 1);
    // prawy margines strony — do końca osi (oś jest zaokrąglana do pełnych px)
    expect(zones[2].left).toBeCloseTo(19 * PX, 1);
    expect(zones[2].width).toBeCloseTo(r.axisPxScaled - 19 * PX, 1);
  });

  it('hGrayZones bez kolumn: dwie strefy podążające za uchwytami (dotychczasowe zachowanie)', () => {
    const r = createRuler({ columnSegments: null, blockIndent: { start: 1, end: 0 } });
    const zones = r.hGrayZones;
    expect(zones.length).toBe(2);
    expect(zones[0].width).toBeCloseTo(3 * PX, 1); // margines 2 + wcięcie 1
    expect(zones[1].width).toBeCloseTo(2 * PX, 1);
  });

  it('hGrayZones respektuje zoom', () => {
    const r = createRuler({ zoomLevel: 150 });
    const zones = r.hGrayZones;
    expect(zones[1].left).toBeCloseTo(9.9 * PX * 1.5, 1);
  });

  // ---------- drag wcięcia w kolumnie ----------

  function drag(r: RulerComponent, side: 'start' | 'end', fromClientX: number, toClientX: number): void {
    const down = new MouseEvent('mousedown', { clientX: fromClientX, bubbles: true });
    side === 'start' ? r.onStartHandleDown(down) : r.onEndHandleDown(down);
    r.onMouseMove(new MouseEvent('mousemove', { clientX: toClientX, bubbles: true }));
    r.onMouseUp();
  }

  it('drag start-uchwytu w 2. kolumnie emituje wcięcie względem tej kolumny', () => {
    const r = createRuler({ activeColumnIndex: 1 });
    let emitted: { start?: number; end?: number } | null = null;
    r.blockIndentChange.subscribe(v => (emitted = v));

    // uchwyt stoi na 11.1 cm; przeciągamy o +1 cm
    drag(r, 'start', 11.1 * PX, 12.1 * PX);
    expect(emitted).not.toBeNull();
    expect(emitted!.start).toBeCloseTo(1, 2);
    expect(emitted!.end).toBeUndefined();
  });

  it('drag start-uchwytu nie wychodzi w lewo poza swoją kolumnę (wcięcie ≥ 0)', () => {
    const r = createRuler({ activeColumnIndex: 1 });
    let emitted: { start?: number; end?: number } | null = null;
    r.blockIndentChange.subscribe(v => (emitted = v));

    // próba wyciągnięcia uchwytu do 5 cm (środek 1. kolumny)
    drag(r, 'start', 11.1 * PX, 5 * PX);
    expect(emitted!.start).toBeCloseTo(0, 2);
  });

  it('drag end-uchwytu w 1. kolumnie nie zjada całej kolumny (zostaje minimum treści)', () => {
    const r = createRuler({ activeColumnIndex: 0 });
    let emitted: { start?: number; end?: number } | null = null;
    r.blockIndentChange.subscribe(v => (emitted = v));

    // end-uchwyt 1. kolumny stoi na 9.9 cm od lewej; ciągniemy do 2.2 cm (prawie start kolumny)
    drag(r, 'end', 9.9 * PX, 2.2 * PX);
    expect(emitted!.end).toBeDefined();
    // wcięcie z prawej ≤ szerokość kolumny − 1 cm treści
    expect(emitted!.end!).toBeLessThanOrEqual(7.9 - 1 + 0.01);
    expect(emitted!.end!).toBeGreaterThan(0);
  });

  it('drag bez kolumn działa jak dotychczas (wcięcie może być ujemne do −margines+0.1)', () => {
    const r = createRuler({ columnSegments: null });
    let emitted: { start?: number; end?: number } | null = null;
    r.blockIndentChange.subscribe(v => (emitted = v));

    drag(r, 'start', 2 * PX, 0);
    expect(emitted!.start).toBeCloseTo(-1.9, 2); // clamp 0.1 cm od krawędzi kartki
  });
});
