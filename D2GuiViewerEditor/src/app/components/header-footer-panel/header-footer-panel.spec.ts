import { TestBed, ComponentFixture } from '@angular/core/testing';
import { HeaderFooterPanelComponent } from './header-footer-panel';

/**
 * Smoke + behaviour tests for the side panel that replaces the old floating
 * header/footer toolbar. The component is stateless — it forwards each action
 * to the parent — so the tests focus on labels (context awareness) and that
 * every button emits the right output.
 */
describe('HeaderFooterPanelComponent', () => {
  let fixture: ComponentFixture<HeaderFooterPanelComponent>;
  let component: HeaderFooterPanelComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [HeaderFooterPanelComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(HeaderFooterPanelComponent);
    component = fixture.componentInstance;
  });

  function render() { fixture.detectChanges(); }
  function text(): string { return fixture.nativeElement.textContent ?? ''; }
  function btn(label: RegExp): HTMLButtonElement | null {
    const all = Array.from(fixture.nativeElement.querySelectorAll('button')) as HTMLButtonElement[];
    return all.find(b => label.test((b.textContent ?? '').trim())) ?? null;
  }

  it('shows the "Nagłówek" context label when editing the header', () => {
    component.section = 'header';
    render();
    expect(text()).toContain('Edytujesz:');
    expect(text()).toContain('Nagłówek');
  });

  it('shows the "Stopka" context label and footer-specific actions when editing the footer', () => {
    component.section = 'footer';
    render();
    expect(text()).toContain('Stopka');
    expect(btn(/Format stopki/)).not.toBeNull();
    expect(btn(/Usuń stopkę/)).not.toBeNull();
  });

  it('emits close when the "Zamknij nagłówek i stopkę" button is clicked', () => {
    component.section = 'header';
    render();
    let closed = 0;
    component.close.subscribe(() => closed++);

    btn(/^Zamknij nagłówek i stopkę$/)!.click();

    expect(closed).toBe(1);
  });

  it('emits toggleDifferentFirstPage when the checkbox changes', () => {
    component.section = 'header';
    render();
    let toggled = 0;
    component.toggleDifferentFirstPage.subscribe(() => toggled++);

    const checkbox = fixture.nativeElement.querySelector('input[type=checkbox]') as HTMLInputElement;
    checkbox.checked = true;
    checkbox.dispatchEvent(new Event('change'));

    expect(toggled).toBe(1);
  });

  it('routes every action button to its dedicated output', () => {
    component.section = 'header';
    render();
    const seen: string[] = [];
    component.insertImage.subscribe(() => seen.push('insertImage'));
    component.insertPageNumbers.subscribe(() => seen.push('insertPageNumbers'));
    component.openFormatDialog.subscribe(() => seen.push('openFormatDialog'));
    component.removeSection.subscribe(() => seen.push('removeSection'));

    btn(/Wstaw obraz/)!.click();
    btn(/Numery stron/)!.click();
    btn(/Format nagłówka/)!.click();
    btn(/Usuń nagłówek/)!.click();

    expect(seen).toEqual(['insertImage', 'insertPageNumbers', 'openFormatDialog', 'removeSection']);
  });

  it('reflects the differentFirstPage input on the checkbox', () => {
    component.section = 'header';
    component.differentFirstPage = true;
    render();
    const checkbox = fixture.nativeElement.querySelector('input[type=checkbox]') as HTMLInputElement;
    expect(checkbox.checked).toBe(true);
  });
});
