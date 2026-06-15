import { TestBed } from '@angular/core/testing';
import { provideRouter } from '@angular/router';
import { DocumentAccessDeniedComponent } from './document-access-denied';

describe('DocumentAccessDeniedComponent', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocumentAccessDeniedComponent],
      providers: [provideRouter([])],
    }).compileComponents();
  });

  it('creates', () => {
    const fixture = TestBed.createComponent(DocumentAccessDeniedComponent);
    expect(fixture.componentInstance).toBeTruthy();
  });

  it('shows the access-denied message and a link home', async () => {
    const fixture = TestBed.createComponent(DocumentAccessDeniedComponent);
    await fixture.whenStable();
    fixture.detectChanges();
    const el = fixture.nativeElement as HTMLElement;
    expect(el.textContent).toContain('Nie masz uprawnień do podglądu tego dokumentu.');
    expect(el.querySelector('a.ad-link')).not.toBeNull();
  });
});
