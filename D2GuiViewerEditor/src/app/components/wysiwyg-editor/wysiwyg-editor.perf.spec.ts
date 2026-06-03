import { TestBed, ComponentFixture } from '@angular/core/testing';
import { WysiwygEditorComponent } from './wysiwyg-editor';

/**
 * Lightweight, REPORT-ONLY performance harness for the editor's CPU-bound serialization paths
 * (setContent split, getContent merge/serialize, split-table merge). It logs timings and only
 * guards against catastrophic regressions with a lenient upper bound — it is not a strict gate.
 *
 * Layout-dependent timing (repaginate by height, real render) needs a real browser and is out of
 * scope for jsdom; measure that with a browser harness (Playwright) if/when added. See
 * `.ai/BENCHMARKS.md`.
 */
describe('WysiwygEditorComponent — perf harness (report-only)', () => {
  let fixture: ComponentFixture<WysiwygEditorComponent>;
  let component: WysiwygEditorComponent;

  // Generous ceiling: these are jsdom string/DOM ops; flag only pathological slowdowns.
  const CEILING_MS = 4000;

  beforeEach(async () => {
    await TestBed.configureTestingModule({ imports: [WysiwygEditorComponent] }).compileComponents();
    fixture = TestBed.createComponent(WysiwygEditorComponent);
    component = fixture.componentInstance;
  });

  function paragraphs(n: number): string {
    let s = '';
    for (let i = 0; i < n; i++) s += `<p>Akapit ${i} — przykładowa treść dokumentu o umiarkowanej długości.</p>`;
    return s;
  }

  function bigTable(rows: number): string {
    let s = '<table><colgroup><col style="width:200px;"/><col style="width:300px;"/></colgroup><tbody>';
    for (let r = 0; r < rows; r++) s += `<tr><td>Komórka ${r}A</td><td>Komórka ${r}B</td></tr>`;
    return s + '</tbody></table>';
  }

  function mockPages(...htmls: string[]) {
    const refs = htmls.map(h => {
      const el = document.createElement('div');
      el.innerHTML = h;
      return { nativeElement: el };
    });
    (component as any).pageEditorRefs = { toArray: () => refs };
  }

  function measure(label: string, fn: () => void): number {
    const t0 = performance.now();
    fn();
    const dt = performance.now() - t0;
    // eslint-disable-next-line no-console
    console.log(`[perf] ${label}: ${dt.toFixed(1)} ms`);
    return dt;
  }

  it('getContent — duża zawartość (1 strona, 800 akapitów)', () => {
    mockPages(paragraphs(800));
    let out = '';
    const dt = measure('getContent 800 paragraphs', () => { out = component.getContent(); });
    expect(out.length).toBeGreaterThan(0);
    expect(dt).toBeLessThan(CEILING_MS);
  });

  it('getContent — duża tabela (500 wierszy)', () => {
    mockPages(bigTable(500));
    let out = '';
    const dt = measure('getContent table 500 rows', () => { out = component.getContent(); });
    expect(out).toContain('Komórka 499B');
    expect(dt).toBeLessThan(CEILING_MS);
  });

  it('getContent — scalanie fragmentów split-table (20 fragmentów × 25 wierszy)', () => {
    const frags: string[] = [];
    for (let f = 0; f < 20; f++) {
      let t = '<table data-split-table-id="st-perf"><tbody>';
      for (let r = 0; r < 25; r++) t += `<tr><td>F${f}R${r}</td></tr>`;
      frags.push(t + '</tbody></table>');
    }
    mockPages(...frags);
    let out = '';
    const dt = measure('getContent merge 20 split-table fragments', () => { out = component.getContent(); });
    const tmp = document.createElement('div');
    tmp.innerHTML = out;
    expect(tmp.querySelectorAll('table').length).toBe(1); // merged into one
    expect(dt).toBeLessThan(CEILING_MS);
  });

  it('setContent — split dużego HTML z page-breakami', () => {
    let html = '';
    for (let i = 0; i < 50; i++) html += paragraphs(20) + '<div class="page-break"></div>';
    const dt = measure('setContent 50 pages', () => { component.setContent(html); });
    expect((component as any).pageContents().length).toBeGreaterThan(0);
    expect(dt).toBeLessThan(CEILING_MS);
  });
});
