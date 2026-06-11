import { Component, computed, inject, signal } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';

/**
 * Panel administracji — „Otwórz plik w edytorze". Formularz na ręczne wpisanie `masterId`
 * (wymagane) i `versionId` (opcjonalne) + przycisk „Otwórz". Nawiguje do edytora z tymi samymi
 * query params, których edytor już używa (`/editor?masterId=&versionId=`): z `versionId` → tryb
 * edycji wskazanej wersji, bez → tryb podglądu wersji bazowej.
 */
@Component({
  selector: 'd2-admin-open-document',
  standalone: true,
  imports: [FormsModule],
  templateUrl: './admin-open-document.html',
  styleUrl: './admin-open-document.scss'
})
export class AdminOpenDocumentComponent {
  private router = inject(Router);

  masterId = signal('');
  versionId = signal('');

  /** Otwarcie możliwe tylko gdy podano masterId (versionId jest opcjonalne). */
  readonly canOpen = computed(() => this.masterId().trim().length > 0);

  open(): void {
    const masterId = this.masterId().trim();
    if (!masterId) return;
    const versionId = this.versionId().trim();

    this.router.navigate(['/editor'], {
      queryParams: versionId ? { masterId, versionId } : { masterId }
    });
  }
}
