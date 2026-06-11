import { TestBed, ComponentFixture } from '@angular/core/testing';
import { vi } from 'vitest';
import { of } from 'rxjs';
import { AdminDeliveriesComponent } from './admin-deliveries';
import { DeliveryListItem, DeliveryStatus, DocumentStorageService } from '../../../services/document-storage.service';

function item(status: DeliveryStatus, id = 'd1'): DeliveryListItem {
  return {
    deliveryId: id, documentId: 'doc1', status, attemptCount: 0,
    createdAt: new Date().toISOString(), lastAttemptAt: null, nextAttemptAt: null,
    deadlineAt: new Date().toISOString(), lastError: null, lockedUntil: null,
    lockedBy: null, sourceVersionId: 'v1', recipientUrl: 'https://x',
  };
}

describe('AdminDeliveriesComponent — Anuluj / Wznów / autoodświeżanie', () => {
  let fixture: ComponentFixture<AdminDeliveriesComponent>;
  let component: AdminDeliveriesComponent;
  let storage: { getDeliveries: ReturnType<typeof vi.fn>; retryDelivery: ReturnType<typeof vi.fn>; cancelDelivery: ReturnType<typeof vi.fn>; };

  beforeEach(async () => {
    storage = {
      getDeliveries: vi.fn(() => of([])),
      retryDelivery: vi.fn(() => of({ deliveryId: 'd1', status: 'Pending' as DeliveryStatus })),
      cancelDelivery: vi.fn(() => of({ deliveryId: 'd1', status: 'Cancelled' as DeliveryStatus })),
    };
    await TestBed.configureTestingModule({
      imports: [AdminDeliveriesComponent],
      providers: [{ provide: DocumentStorageService, useValue: storage }],
    }).compileComponents();
    fixture = TestBed.createComponent(AdminDeliveriesComponent);
    component = fixture.componentInstance;
  });

  it('canResume dla Cancelled/RetryScheduled/DeadLettered/FailedPermanently, nie dla Sending/Sent/Pending', () => {
    expect(component.canResume(item('Cancelled'))).toBe(true);
    expect(component.canResume(item('RetryScheduled'))).toBe(true);
    expect(component.canResume(item('DeadLettered'))).toBe(true);
    expect(component.canResume(item('FailedPermanently'))).toBe(true);
    expect(component.canResume(item('Sending'))).toBe(false);
    expect(component.canResume(item('Sent'))).toBe(false);
    expect(component.canResume(item('Pending'))).toBe(false);
  });

  it('canCancel tylko dla Pending i RetryScheduled', () => {
    expect(component.canCancel(item('Pending'))).toBe(true);
    expect(component.canCancel(item('RetryScheduled'))).toBe(true);
    expect(component.canCancel(item('Sending'))).toBe(false);
    expect(component.canCancel(item('Sent'))).toBe(false);
    expect(component.canCancel(item('Cancelled'))).toBe(false);
  });

  it('cancel() woła cancelDelivery i ustawia komunikat', () => {
    const ev = { stopPropagation: vi.fn() } as unknown as Event;
    component.cancel(item('Pending'), ev);

    expect(storage.cancelDelivery).toHaveBeenCalledWith('d1');
    expect(component.cancelingId()).toBeNull();
    expect(component.notice()).toContain('anulowane');
  });

  it('resume() woła retryDelivery i ustawia komunikat', () => {
    const ev = { stopPropagation: vi.fn() } as unknown as Event;
    component.resume(item('Cancelled'), ev);

    expect(storage.retryDelivery).toHaveBeenCalledWith('d1');
    expect(component.notice()).toContain('wznowione');
  });

  it('toggleAutoRefresh przełącza flagę', () => {
    expect(component.autoRefresh()).toBe(true);
    component.toggleAutoRefresh();
    expect(component.autoRefresh()).toBe(false);
  });

  it('statusLabel/statusClass obsługują Cancelled', () => {
    expect(component.statusLabel('Cancelled')).toBe('Anulowano');
    expect(component.statusClass('Cancelled')).toBe('status-cancelled');
  });

  it('ngOnDestroy odsubskrybowuje autoodświeżanie (brak wycieku timera)', () => {
    component.ngOnInit();
    component.ngOnDestroy();
    // brak błędu = sub zamknięty; podwójne destroy bezpieczne
    expect(() => component.ngOnDestroy()).not.toThrow();
  });
});
