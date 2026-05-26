import { TestBed, ComponentFixture } from '@angular/core/testing';
import { DocumentClassificationBadgeComponent } from './document-classification-badge';

describe('DocumentClassificationBadgeComponent', () => {
  let fixture: ComponentFixture<DocumentClassificationBadgeComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentClassificationBadgeComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(DocumentClassificationBadgeComponent);
  });

  function badge(): HTMLElement | null {
    return fixture.nativeElement.querySelector('.classification-badge');
  }

  async function setClassification(value: string | null | undefined): Promise<void> {
    fixture.componentRef.setInput('classification', value);
    await fixture.whenStable();
    fixture.detectChanges();
  }

  it('renders the value and a dictionary level class for known C-levels', async () => {
    await setClassification('C2');
    const el = badge();
    expect(el).not.toBeNull();
    expect(el!.textContent).toContain('C2');
    expect(el!.classList).toContain('lvl-c2');
  });

  it('shows unknown values verbatim with a neutral style (no allow-list)', async () => {
    await setClassification('TOP-SECRET');
    const el = badge();
    expect(el!.textContent).toContain('TOP-SECRET');
    expect(el!.classList).toContain('lvl-unknown');
  });

  it('renders nothing when classification is missing/null/empty/whitespace', async () => {
    for (const v of [undefined, null, '', '   '] as const) {
      await setClassification(v);
      expect(badge()).toBeNull();
    }
  });

  it('normalizes surrounding whitespace for a present value', async () => {
    await setClassification('  C4  ');
    const el = badge();
    expect(el).not.toBeNull();
    expect(el!.textContent).toContain('C4');
    expect(el!.classList).toContain('lvl-c4');
  });
});
