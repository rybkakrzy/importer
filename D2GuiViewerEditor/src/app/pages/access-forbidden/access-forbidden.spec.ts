import { TestBed } from '@angular/core/testing';
import { provideRouter } from '@angular/router';
import { AccessForbiddenComponent } from './access-forbidden';

describe('AccessForbiddenComponent', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AccessForbiddenComponent],
      providers: [provideRouter([])],
    }).compileComponents();
  });

  it('shows the "Brak uprawnień" message and a link home', () => {
    const fixture = TestBed.createComponent(AccessForbiddenComponent);
    fixture.detectChanges();
    const text = (fixture.nativeElement as HTMLElement).textContent ?? '';
    expect(text).toContain('Brak uprawnień');
    expect((fixture.nativeElement as HTMLElement).querySelector('a[href="/"]')).not.toBeNull();
  });
});
