import { TestBed } from '@angular/core/testing';
import { ButtonComponent } from './button';

describe('ButtonComponent', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ButtonComponent],
    }).compileComponents();
  });

  it('should create', () => {
    const fixture = TestBed.createComponent(ButtonComponent);
    expect(fixture.componentInstance).toBeTruthy();
  });

  it('should render slot content', async () => {
    const fixture = TestBed.createComponent(ButtonComponent);
    await fixture.whenStable();
    const el = fixture.nativeElement as HTMLElement;
    expect(el.querySelector('button')).not.toBeNull();
  });

  it('should be disabled when disabled input is true', () => {
    const fixture = TestBed.createComponent(ButtonComponent);
    fixture.componentInstance.disabled = true;
    fixture.detectChanges();
    const btn = fixture.nativeElement.querySelector('button') as HTMLButtonElement;
    expect(btn.disabled).toBe(true);
  });

  it('should emit clicked event on click', () => {
    const fixture = TestBed.createComponent(ButtonComponent);
    fixture.detectChanges();
    const spy = vi.spyOn(fixture.componentInstance.clicked, 'emit');
    const btn = fixture.nativeElement.querySelector('button') as HTMLButtonElement;
    btn.click();
    expect(spy).toHaveBeenCalled();
  });

  it('should show loadingText and disable button when loading', () => {
    const fixture = TestBed.createComponent(ButtonComponent);
    fixture.componentInstance.loading = true;
    fixture.componentInstance.loadingText = 'Proszę czekać...';
    fixture.detectChanges();
    const btn = fixture.nativeElement.querySelector('button') as HTMLButtonElement;
    expect(btn.disabled).toBe(true);
    expect(btn.textContent?.trim()).toBe('Proszę czekać...');
  });

  it('should apply correct variant class', () => {
    const fixture = TestBed.createComponent(ButtonComponent);
    fixture.componentInstance.variant = 'success';
    fixture.detectChanges();
    const btn = fixture.nativeElement.querySelector('button') as HTMLButtonElement;
    expect(btn.classList).toContain('lib-btn--success');
  });
});
