import { Component, computed, inject, OnInit, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import {
  DeliveryListItem,
  DeliveryStatus,
  DocumentStorageService
} from '../../../services/document-storage.service';

@Component({
  selector: 'd2-admin-deliveries',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './admin-deliveries.html',
  styleUrl: './admin-deliveries.scss'
})
export class AdminDeliveriesComponent implements OnInit {
  private storage = inject(DocumentStorageService);

  /** Wszystkie statusy do filtra; DeadLettered najczęściej interesuje admina. */
  readonly statuses: DeliveryStatus[] = [
    'DeadLettered', 'FailedPermanently', 'RetryScheduled', 'Sending', 'Pending', 'Sent'
  ];

  /** Status pobierany z backendu (serwer wymaga konkretnego statusu). */
  selectedStatus = signal<DeliveryStatus>('DeadLettered');

  private allDeliveries = signal<DeliveryListItem[]>([]);
  filterId       = signal('');
  filterDoc      = signal('');
  filterAtt      = signal('');
  filterCreated  = signal('');
  filterLastAtt  = signal('');
  filterNextAtt  = signal('');
  filterDeadline = signal('');
  filterLock     = signal('');
  filterError    = signal('');
  currentPage = signal(0);
  readonly pageSize = 10;

  isLoading = signal(true);
  error = signal<string | null>(null);
  retryingId = signal<string | null>(null);
  notice = signal<string | null>(null);

  private filtered = computed(() => {
    const id       = this.filterId().toLowerCase().trim();
    const doc      = this.filterDoc().toLowerCase().trim();
    const att      = this.filterAtt().toLowerCase().trim();
    const created  = this.filterCreated().toLowerCase().trim();
    const lastAtt  = this.filterLastAtt().toLowerCase().trim();
    const nextAtt  = this.filterNextAtt().toLowerCase().trim();
    const deadline = this.filterDeadline().toLowerCase().trim();
    const lock     = this.filterLock().toLowerCase().trim();
    const err      = this.filterError().toLowerCase().trim();
    return this.allDeliveries().filter(d => {
      if (id       && !d.deliveryId.toLowerCase().includes(id))                          return false;
      if (doc      && !d.documentId.toLowerCase().includes(doc))                         return false;
      if (att      && !String(d.attemptCount).includes(att))                             return false;
      if (created  && !this.formatDate(d.createdAt).toLowerCase().includes(created))     return false;
      if (lastAtt  && !this.formatDate(d.lastAttemptAt).toLowerCase().includes(lastAtt)) return false;
      if (nextAtt  && !this.formatDate(d.nextAttemptAt).toLowerCase().includes(nextAtt)) return false;
      if (deadline && !this.formatDate(d.deadlineAt).toLowerCase().includes(deadline))   return false;
      if (lock     && !(d.lockedBy ?? '').toLowerCase().includes(lock))                  return false;
      if (err      && !(d.lastError ?? '').toLowerCase().includes(err))                  return false;
      return true;
    });
  });

  totalFiltered = computed(() => this.filtered().length);
  totalPages    = computed(() => Math.max(1, Math.ceil(this.totalFiltered() / this.pageSize)));

  deliveries = computed(() => {
    const start = this.currentPage() * this.pageSize;
    return this.filtered().slice(start, start + this.pageSize);
  });

  ngOnInit(): void {
    this.load();
  }

  load(): void {
    this.isLoading.set(true);
    this.error.set(null);
    this.notice.set(null);
    this.currentPage.set(0);
    this.storage.getDeliveriesByStatus(this.selectedStatus()).subscribe({
      next: (items) => {
        this.allDeliveries.set(items);
        this.isLoading.set(false);
      },
      error: () => {
        this.error.set('Nie udało się załadować listy wysyłek.');
        this.isLoading.set(false);
      }
    });
  }

  setStatus(status: string): void {
    this.selectedStatus.set(status as DeliveryStatus);
    this.load();
  }

  setFilter(
    field: 'id' | 'doc' | 'att' | 'created' | 'lastAtt' | 'nextAtt' | 'deadline' | 'lock' | 'error',
    value: string
  ): void {
    if (field === 'id')       this.filterId.set(value);
    if (field === 'doc')      this.filterDoc.set(value);
    if (field === 'att')      this.filterAtt.set(value);
    if (field === 'created')  this.filterCreated.set(value);
    if (field === 'lastAtt')  this.filterLastAtt.set(value);
    if (field === 'nextAtt')  this.filterNextAtt.set(value);
    if (field === 'deadline') this.filterDeadline.set(value);
    if (field === 'lock')     this.filterLock.set(value);
    if (field === 'error')    this.filterError.set(value);
    this.currentPage.set(0);
  }

  /** Retry ma sens tylko dla zadań w stanie terminalnie nieudanym. */
  canRetry(item: DeliveryListItem): boolean {
    return item.status === 'DeadLettered' || item.status === 'FailedPermanently';
  }

  retry(item: DeliveryListItem, event: Event): void {
    event.stopPropagation();
    this.retryingId.set(item.deliveryId);
    this.notice.set(null);
    this.storage.retryDelivery(item.deliveryId).subscribe({
      next: () => {
        this.retryingId.set(null);
        this.notice.set(`Zadanie ${this.shortId(item.deliveryId)} ponowione.`);
        this.load();
      },
      error: () => {
        this.retryingId.set(null);
        this.error.set(`Nie udało się ponowić zadania ${this.shortId(item.deliveryId)}.`);
      }
    });
  }

  /** Etykieta statusu po polsku (wartość enuma zostaje dla backendu). */
  statusLabel(status: DeliveryStatus): string {
    const map: Record<DeliveryStatus, string> = {
      Pending: 'Oczekuje',
      Sending: 'Wysyłanie',
      RetryScheduled: 'Zaplanowano ponowienie',
      Sent: 'Wysłano',
      FailedPermanently: 'Błąd trwały',
      DeadLettered: 'Porzucone'
    };
    return map[status] ?? status;
  }

  /** Status chip CSS class (color) — paleta spójna z listą plików. */
  statusClass(status: DeliveryStatus): string {
    const map: Record<DeliveryStatus, string> = {
      Pending: 'status-saved',
      Sending: 'status-sending',
      RetryScheduled: 'status-editing',
      Sent: 'status-sent',
      FailedPermanently: 'status-failed',
      DeadLettered: 'status-failed'
    };
    return map[status] ?? 'status-saved';
  }

  shortId(id: string): string {
    return id ? id.slice(0, 8) : '—';
  }

  prevPage(): void { if (this.currentPage() > 0) this.currentPage.update(p => p - 1); }
  nextPage(): void { if (this.currentPage() < this.totalPages() - 1) this.currentPage.update(p => p + 1); }
  pageEnd(): number { return Math.min((this.currentPage() + 1) * this.pageSize, this.totalFiltered()); }
  goToPage(value: string): void {
    const n = parseInt(value, 10);
    if (!isNaN(n)) this.currentPage.set(Math.max(0, Math.min(n - 1, this.totalPages() - 1)));
  }

  formatDate(dateStr: string | null): string {
    if (!dateStr) return '—';
    return new Date(dateStr).toLocaleDateString('pl-PL', {
      day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit'
    });
  }
}
