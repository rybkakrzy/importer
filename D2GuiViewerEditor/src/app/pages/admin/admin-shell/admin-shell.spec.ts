import { TestBed, ComponentFixture } from '@angular/core/testing';
import { provideRouter } from '@angular/router';
import { AdminShellComponent } from './admin-shell';

describe('AdminShellComponent — nazewnictwo nawigacji admina', () => {
  let fixture: ComponentFixture<AdminShellComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AdminShellComponent],
      providers: [provideRouter([])],
    }).compileComponents();
    fixture = TestBed.createComponent(AdminShellComponent);
    fixture.detectChanges();
  });

  function navText(): string {
    return (fixture.nativeElement.querySelector('.sidebar-nav') as HTMLElement).textContent ?? '';
  }

  it('pokazuje pozycję "Pliki do wysłania"', () => {
    expect(navText()).toContain('Pliki do wysłania');
  });

  it('nie używa już nazwy "Wysyłki" ani "Workers file"', () => {
    expect(navText()).not.toContain('Wysyłki');
    expect(navText()).not.toContain('Workers');
  });

  it('zachowuje link do trasy deliveries (bez zmiany technicznej ścieżki)', () => {
    const link = fixture.nativeElement.querySelector('a[routerLink="deliveries"]');
    expect(link).not.toBeNull();
  });
});
