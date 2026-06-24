import { Component, computed, inject, OnDestroy, OnInit, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { Subscription, interval } from 'rxjs';
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
export class AdminDeliveriesComponent implements OnInit, OnDestroy {
  private storage = inject(DocumentStorageService);
  private router = inject(Router);

  readonly statuses: DeliveryStatus[] = [
    'DeadLettered', 'FailedPermanently', 'RetryScheduled', 'Sending', 'Pending', 'Sent', 'Cancelled'
  ];

  /** Interwał autoodświeżania listy (ms). */
  private static readonly RefreshIntervalMs = 3000;
  private refreshSub?: Subscription;
  autoRefresh = signal(true);

  selectedStatus = signal<DeliveryStatus | 'all'>('all');

  private allDeliveries = signal<DeliveryListItem[]>([]);
  filterId       = signal('');
  filterDoc      = signal('');
  filterAtt      = signal('');
  filterCreated  = signal('');
  filterLock     = signal('');
  filterModifiedBy = signal('');
  currentPage = signal(0);
  readonly pageSize = 10;

  isLoading = signal(true);
  /** Błąd ŁADOWANIA listy — zastępuje tabelę (gdy nie da się pobrać danych). */
  error = signal<string | null>(null);
  /** Błąd AKCJI (Wznów/Anuluj) — pokazywany jako baner NAD tabelą; NIE chowa tabeli. */
  actionError = signal<string | null>(null);
  retryingId = signal<string | null>(null);
  cancelingId = signal<string | null>(null);
  notice = signal<string | null>(null);
  expandedId = signal<string | null>(null);

  // Edycja adresu odbiorcy (returnUrl) — modal.
  editingId = signal<string | null>(null);
  editUrlValue = signal('');
  editError = signal<string | null>(null);
  savingEdit = signal(false);

  private filtered = computed(() => {
    const id       = this.filterId().toLowerCase().trim();
    const doc      = this.filterDoc().toLowerCase().trim();
    const att      = this.filterAtt().toLowerCase().trim();
    const created  = this.filterCreated().toLowerCase().trim();
    const lock     = this.filterLock().toLowerCase().trim();
    const modifiedBy = this.filterModifiedBy().toLowerCase().trim();
    return this.allDeliveries().filter(d => {
      if (id       && !d.deliveryId.toLowerCase().includes(id))                          return false;
      if (doc      && !d.documentId.toLowerCase().includes(doc))                         return false;
      if (att      && !String(d.attemptCount).includes(att))                             return false;
      if (created  && !this.formatDate(d.createdAt).toLowerCase().includes(created))     return false;
      if (lock     && !(d.lockedBy ?? '').toLowerCase().includes(lock))                  return false;
      if (modifiedBy && !(d.corporateKey ?? '').toLowerCase().includes(modifiedBy))      return false;
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
    this.startAutoRefresh();
  }

  ngOnDestroy(): void {
    this.refreshSub?.unsubscribe();
  }

  /** Pełne załadowanie (ze spinnerem, resetem strony) — przy wejściu i zmianie filtra statusu. */
  load(): void {
    this.isLoading.set(true);
    this.error.set(null);
    this.actionError.set(null);
    this.notice.set(null);
    this.currentPage.set(0);
    this.fetch(/* silent */ false);
  }

  /** Autoodświeżanie co 3 s: ciche pobranie (bez spinnera, bez resetu strony/filtrów/rozwinięcia). */
  private startAutoRefresh(): void {
    this.refreshSub?.unsubscribe();
    this.refreshSub = interval(AdminDeliveriesComponent.RefreshIntervalMs).subscribe(() => {
      // Nie odświeżamy w trakcie ładowania ani trwającej akcji (Anuluj/Wznów) — uniknięcie migotania
      // i nadpisania stanu tuż przed reloadem akcji.
      if (this.autoRefresh() && !this.isLoading() && !this.retryingId()
          && !this.cancelingId() && !this.editingId()) {
        this.fetch(/* silent */ true);
      }
    });
  }

  toggleAutoRefresh(): void {
    this.autoRefresh.update(v => !v);
  }

  private fetch(silent: boolean): void {
    const status = this.selectedStatus();
    this.storage.getDeliveries(status === 'all' ? null : status).subscribe({
      next: (items) => {
        this.allDeliveries.set(items);
        if (!silent) this.isLoading.set(false);
        // Strona mogła się skurczyć (np. zadanie zniknęło z filtra) — przytnij do zakresu.
        if (this.currentPage() > this.totalPages() - 1) {
          this.currentPage.set(Math.max(0, this.totalPages() - 1));
        }
      },
      error: () => {
        if (!silent) {
          this.error.set('Nie udało się załadować listy plików do wysłania.');
          this.isLoading.set(false);
        }
        // Ciche odświeżenie nie pokazuje błędu — następny tick spróbuje ponownie.
      }
    });
  }

  setStatus(status: string): void {
    this.selectedStatus.set(status as DeliveryStatus | 'all');
    this.load();
  }

  setFilter(
    field: 'id' | 'doc' | 'att' | 'created' | 'lock' | 'modifiedBy',
    value: string
  ): void {
    if (field === 'id')       this.filterId.set(value);
    if (field === 'doc')      this.filterDoc.set(value);
    if (field === 'att')      this.filterAtt.set(value);
    if (field === 'created')  this.filterCreated.set(value);
    if (field === 'lock')     this.filterLock.set(value);
    if (field === 'modifiedBy') this.filterModifiedBy.set(value);
    this.currentPage.set(0);
  }

  toggleExpand(deliveryId: string): void {
    this.expandedId.update(curr => (curr === deliveryId ? null : deliveryId));
  }

  isExpanded(deliveryId: string): boolean {
    return this.expandedId() === deliveryId;
  }

  /** „Wznów" — zadania nieudane, zaplanowane na później lub anulowane wracają do kolejki (wysyłka teraz). */
  canResume(item: DeliveryListItem): boolean {
    return item.status === 'DeadLettered'
        || item.status === 'FailedPermanently'
        || item.status === 'RetryScheduled'
        || item.status === 'Cancelled';
  }

  /** „Anuluj" — tylko zadania jeszcze nieprzetworzone i niezablokowane (Pending / RetryScheduled). */
  canCancel(item: DeliveryListItem): boolean {
    return item.status === 'Pending' || item.status === 'RetryScheduled';
  }

  /** Czy dla zadania jest jakakolwiek akcja w toku (blokuje przyciski w wierszu). */
  isBusy(item: DeliveryListItem): boolean {
    return this.retryingId() === item.deliveryId || this.cancelingId() === item.deliveryId;
  }

  resume(item: DeliveryListItem, event: Event): void {
    event.stopPropagation();
    this.retryingId.set(item.deliveryId);
    this.notice.set(null);
    this.actionError.set(null);
    this.storage.retryDelivery(item.deliveryId).subscribe({
      next: () => {
        this.retryingId.set(null);
        this.notice.set(`Zadanie ${this.shortId(item.deliveryId)} wznowione.`);
        this.refreshAfterAction();
      },
      error: (err) => {
        this.retryingId.set(null);
        // Błąd akcji NIE chowa tabeli — baner nad listą (osobny od błędu ładowania).
        this.actionError.set(err?.error?.error
          || `Nie udało się wznowić zadania ${this.shortId(item.deliveryId)}.`);
      }
    });
  }

  cancel(item: DeliveryListItem, event: Event): void {
    event.stopPropagation();
    this.cancelingId.set(item.deliveryId);
    this.notice.set(null);
    this.actionError.set(null);
    this.storage.cancelDelivery(item.deliveryId).subscribe({
      next: () => {
        this.cancelingId.set(null);
        this.notice.set(`Zadanie ${this.shortId(item.deliveryId)} anulowane.`);
        this.refreshAfterAction();
      },
      error: (err) => {
        this.cancelingId.set(null);
        this.actionError.set(err?.error?.error
          || `Nie udało się anulować zadania ${this.shortId(item.deliveryId)}.`);
      }
    });
  }

  dismissActionError(): void {
    this.actionError.set(null);
  }

  /** Po akcji odświeżamy cicho (bez spinnera/resetu strony) — lista i tak auto-odświeża się co 3 s. */
  private refreshAfterAction(): void {
    this.fetch(/* silent */ true);
  }

  /**
   * „Kopiuj link" — buduje pełny, sformatowany link do edycji na podstawie danych wysyłki
   * (`documentId`=masterId, `sourceVersionId`=versionId) i kopiuje go do schowka. Z versionId edytor
   * otworzy wskazaną wersję edytowalną (tryb edycji), bez niego — podgląd wersji bazowej.
   */
  copyEditLink(item: DeliveryListItem, event: Event): void {
    event.stopPropagation();
    const masterId = item.documentId;
    const versionId = item.sourceVersionId;
    const tree = this.router.createUrlTree(['/editor'], {
      queryParams: versionId ? { masterId, versionId } : { masterId }
    });
    const url = `${window.location.origin}${this.router.serializeUrl(tree)}`;

    this.notice.set(null);
    this.actionError.set(null);

    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(url).then(
        () => this.notice.set('Skopiowano link do edycji do schowka.'),
        () => this.actionError.set(`Nie udało się skopiować linku. Skopiuj ręcznie: ${url}`)
      );
    } else {
      // Brak Clipboard API (np. kontekst nie-HTTPS) — pokaż link do ręcznego skopiowania.
      this.actionError.set(`Kopiowanie niedostępne w tej przeglądarce. Link do edycji: ${url}`);
    }
  }

  /** „Edytuj" — można zmienić adres odbiorcy, dopóki zadanie nie zostało wysłane / nie jest w toku. */
  canEdit(item: DeliveryListItem): boolean {
    return item.status !== 'Sent' && item.status !== 'Sending';
  }

  startEdit(item: DeliveryListItem, event: Event): void {
    event.stopPropagation();
    this.editError.set(null);
    this.editUrlValue.set(item.recipientUrl ?? '');
    this.editingId.set(item.deliveryId);
  }

  cancelEdit(): void {
    this.editingId.set(null);
    this.editUrlValue.set('');
    this.editError.set(null);
    this.savingEdit.set(false);
  }

  saveEdit(): void {
    const deliveryId = this.editingId();
    if (!deliveryId) return;
    const url = this.editUrlValue().trim();
    if (!/^https?:\/\/.+/i.test(url)) {
      this.editError.set('Podaj poprawny adres http(s).');
      return;
    }

    this.savingEdit.set(true);
    this.editError.set(null);
    this.storage.updateDeliveryRecipientUrl(deliveryId, url).subscribe({
      next: () => {
        this.savingEdit.set(false);
        const id = this.shortId(deliveryId);
        this.cancelEdit();
        this.notice.set(`Zmieniono adres odbiorcy zadania ${id}.`);
        this.refreshAfterAction();
      },
      error: (err) => {
        this.savingEdit.set(false);
        this.editError.set(err?.error?.error || 'Nie udało się zmienić adresu odbiorcy.');
      }
    });
  }

  /** Etykieta statusu po polsku (wartość enuma zostaje dla backendu). */
  statusLabel(status: DeliveryStatus): string {
    const map: Record<DeliveryStatus, string> = {
      Pending: 'Oczekuje',
      Sending: 'Wysyłanie',
      RetryScheduled: 'Zaplanowano',
      Sent: 'Wysłano',
      FailedPermanently: 'Błąd',
      DeadLettered: 'Porzucone',
      Cancelled: 'Anulowano'
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
      DeadLettered: 'status-failed',
      Cancelled: 'status-cancelled'
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
