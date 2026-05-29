import { TestBed, ComponentFixture } from '@angular/core/testing';
import { ImagePropertiesPanelComponent, ImageSelectionState } from './image-properties-panel';

/**
 * Stateless side panel: every input mutation forwards an event to the parent. Tests
 * focus on event emission, snapshot mirroring and label-vs-state correctness.
 */
describe('ImagePropertiesPanelComponent', () => {
  let fixture: ComponentFixture<ImagePropertiesPanelComponent>;
  let component: ImagePropertiesPanelComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ImagePropertiesPanelComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(ImagePropertiesPanelComponent);
    component = fixture.componentInstance;
  });

  function render() { fixture.detectChanges(); }
  function text(): string { return fixture.nativeElement.textContent ?? ''; }
  function btn(label: RegExp): HTMLButtonElement | null {
    const all = Array.from(fixture.nativeElement.querySelectorAll('button')) as HTMLButtonElement[];
    // Normalise &nbsp; (U+00A0) → regular space so tests can use natural patterns
    // without leaking the encoded form.
    return all.find(b => label.test((b.textContent ?? '').replace(/ /g, ' ').trim())) ?? null;
  }
  function input(selector: string): HTMLInputElement {
    return fixture.nativeElement.querySelector(selector);
  }

  const sample: ImageSelectionState = {
    widthPx: 240,
    heightPx: 160,
    aspectRatio: 240 / 160,
    alignment: 'center',
    positionMode: 'inline',
  };

  it('shows the empty hint when no image is selected', () => {
    render();
    expect(text()).toContain('Zaznacz obraz');
  });

  it('mirrors width/height inputs from the state snapshot', () => {
    component.state = sample;
    render();
    // Read the signal directly — the rendered <input value=...> attribute depends on
    // ngModel's microtask schedule and is brittle in vitest with signal-driven CD.
    expect((component as any)._widthInput()).toBe(240);
    expect((component as any)._heightInput()).toBe(160);
  });

  it('emits widthChange via applyWidth() with the current signal value', () => {
    component.state = sample;
    render();
    const widths: number[] = [];
    component.widthChange.subscribe(v => widths.push(v));

    (component as any)._widthInput.set(300);
    (component as any).applyWidth();

    expect(widths).toEqual([300]);
  });

  it('clamps negative / non-finite width to 16 px (the editor floor)', () => {
    component.state = sample;
    render();
    const widths: number[] = [];
    component.widthChange.subscribe(v => widths.push(v));

    (component as any)._widthInput.set(-50);
    (component as any).applyWidth();

    expect(widths).toEqual([16]);
  });

  it('emits lockAspectChange when the checkbox toggles', () => {
    component.state = sample;
    render();
    let last: boolean | null = null;
    component.lockAspectChange.subscribe(v => last = v);

    const checkbox = fixture.nativeElement.querySelector('input[type=checkbox]') as HTMLInputElement;
    checkbox.checked = false;
    checkbox.dispatchEvent(new Event('change'));

    expect(last).toBe(false);
  });

  it('highlights the active alignment button and emits on click', () => {
    component.state = { ...sample, alignment: 'right' };
    render();
    const lastValues: (string | null)[] = [];
    component.alignmentChange.subscribe(v => lastValues.push(v));

    btn(/^L$/)!.click();
    btn(/^C$/)!.click();
    btn(/^—$/)!.click();

    expect(lastValues).toEqual(['left', 'center', null]);
    // The 'R' button should reflect the input state as active.
    const r = btn(/^R$/)!;
    expect(r.classList.contains('is-active')).toBe(true);
  });

  it('emits removeImage and resetAspect via their dedicated buttons', () => {
    component.state = sample;
    render();
    let removeCalls = 0;
    let resetCalls = 0;
    component.removeImage.subscribe(() => removeCalls++);
    component.resetAspect.subscribe(() => resetCalls++);

    btn(/Usuń obraz/)!.click();
    btn(/Przywróć proporcje/)!.click();

    expect(removeCalls).toBe(1);
    expect(resetCalls).toBe(1);
  });

  it('emits close when the X header button is clicked (deselect)', () => {
    component.state = sample;
    render();
    let closed = 0;
    component.close.subscribe(() => closed++);

    btn(/^×$/)!.click();

    expect(closed).toBe(1);
  });

  it('shows position-mode buttons + active state for the current mode', () => {
    component.state = { ...sample, positionMode: 'behind' };
    render();
    const behindBtn = btn(/^Za$/)!;
    expect(behindBtn.classList.contains('is-active')).toBe(true);
    const inlineBtn = btn(/W tekście/)!;
    expect(inlineBtn.classList.contains('is-active')).toBe(false);
  });

  it('emits positionModeChange when one of the three buttons is clicked', () => {
    component.state = sample;
    render();
    const seen: string[] = [];
    component.positionModeChange.subscribe(v => seen.push(v));

    btn(/^Przed$/)!.click();
    btn(/^Za$/)!.click();
    btn(/W tekście/)!.click();

    expect(seen).toEqual(['front', 'behind', 'inline']);
  });

  it('points out the square/tight roadmap explicitly in the panel UI', () => {
    component.state = sample;
    render();
    expect(text()).toContain('Square/tight wrap');
    expect(text()).toContain('w przygotowaniu');
  });
});
