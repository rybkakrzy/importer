import { Component, inject, signal, computed } from '@angular/core';
import { CommonModule } from '@angular/common';
import { switchMap, map, from, Subscription } from 'rxjs';
import { MsalService } from '@azure/msal-angular';
import type { AccountInfo } from '@azure/msal-browser';
import { DocumentService } from '../../services/document.service';
import { DocumentStorageService } from '../../services/document-storage.service';
import { DocumentNavigationService } from '../../core/services/document-navigation.service';

@Component({
  selector: 'd2-dashboard',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class DashboardComponent {
  private documentService = inject(DocumentService);
  private documentStorageService = inject(DocumentStorageService);
  private documentNavigation = inject(DocumentNavigationService);
  private readonly msal = inject(MsalService);

  isLoading = signal(false);
  errorMessage = signal<string | null>(null);
  readonly currentYear = new Date().getFullYear();

  // Imię zalogowanego użytkownika (Entra ID) — powitanie na pulpicie.
  readonly userFirstName = signal<string>(this.resolveFirstName());
  readonly greeting = computed(() => {
    const name = this.userFirstName();
    return name ? `Witaj ${name} w Doc2` : 'Witaj w Doc2';
  });

  private activeSubscription: Subscription | null = null;


  newDocument(): void {
    this.isLoading.set(true);
    this.errorMessage.set(null);

    this.activeSubscription = this.documentService.newDocument().pipe(
      switchMap(content =>
        this.documentService.saveDocument({ html: content.html, metadata: content.metadata }).pipe(
          switchMap(blob =>
            from(this.blobToBase64(blob)).pipe(
              switchMap(base64 =>
                this.documentStorageService.uploadDocument({
                  name: 'Nowy dokument.docx',
                  mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                  content: base64
                }).pipe(
                  // Nowy dokument = zamiar edycji → twórz wersję edytowalną (v2).
                  switchMap(result =>
                    this.documentStorageService.saveDocumentVersion(result.masterId, { content: base64 }).pipe(
                      map(saved => ({ masterId: result.masterId, versionId: saved.versionId }))
                    )
                  )
                )
              )
            )
          )
        )
      )
    ).subscribe({
      next: ({ masterId, versionId }) => {
        this.documentNavigation.navigateToEditableDocument(masterId, versionId);
      },
      error: () => {
        this.isLoading.set(false);
        this.errorMessage.set('Błąd podczas tworzenia dokumentu. Spróbuj ponownie.');
      }
    });
  }

  cancelLoading(): void {
    this.activeSubscription?.unsubscribe();
    this.activeSubscription = null;
    this.isLoading.set(false);
  }

  openFile(): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.docx,.doc,.pdf';

    input.onchange = async (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;

      const lowerName = file.name.toLowerCase();

      // Ta ścieżka (utwórz nowy z dysku) zapisuje plik bez normalizacji, więc binarny .doc /
      // DOCX z hasłem nie zostaną tu przetworzone. Pliki .doc / zabezpieczone hasłem otwiera się
      // w edytorze przez „Plik → Otwórz" (ścieżka /open z dekrypcją i detekcją .doc).
      if (lowerName.endsWith('.doc')) {
        this.errorMessage.set('Plik .doc otwórz w edytorze przez „Plik → Otwórz" (obsługuje .doc oraz DOCX zabezpieczone hasłem). Tutaj wczytasz pliki .docx i .pdf.');
        return;
      }
      if (!lowerName.endsWith('.docx') && !lowerName.endsWith('.pdf')) {
        this.errorMessage.set('Obsługiwane są pliki DOCX i PDF.');
        return;
      }

      if (lowerName.endsWith('.pdf')) {
        this.isLoading.set(true);
        this.errorMessage.set(null);
        try {
          const base64 = await this.documentStorageService.fileToBase64(file);
      this.activeSubscription = this.documentStorageService.uploadDocument({
            name: file.name,
            mimeType: 'application/pdf',
            content: base64,
          }).subscribe({
            next: (result) => {
              this.documentNavigation.navigateToDocument(result.masterId, 'application/pdf');
            },
            error: () => {
              this.isLoading.set(false);
              this.errorMessage.set('Błąd podczas wczytywania pliku PDF. Spróbuj ponownie.');
            },
          });
        } catch {
          this.isLoading.set(false);
          this.errorMessage.set('Błąd podczas odczytu pliku PDF.');
        }
        return;
      }

      this.isLoading.set(true);
      this.errorMessage.set(null);

      try {
        const base64 = await this.documentStorageService.fileToBase64(file);
        this.activeSubscription = this.documentStorageService.uploadDocument({
          name: file.name,
          mimeType: file.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          content: base64
        }).pipe(
          // Ręczne wczytanie z dysku = zamiar edycji → twórz wersję edytowalną (v2)
          // i otwórz w trybie edycji. Oryginał (v1) pozostaje niezmienny.
          switchMap(result =>
            this.documentStorageService.saveDocumentVersion(result.masterId, { content: base64 }).pipe(
              map(saved => ({ masterId: result.masterId, versionId: saved.versionId }))
            )
          )
        ).subscribe({
          next: ({ masterId, versionId }) => {
            this.documentNavigation.navigateToEditableDocument(masterId, versionId);
          },
          error: () => {
            this.isLoading.set(false);
            this.errorMessage.set('Błąd podczas wczytywania dokumentu. Spróbuj ponownie.');
          }
        });
      } catch {
        this.isLoading.set(false);
        this.errorMessage.set('Błąd podczas odczytu pliku.');
      }
    };

    input.click();
  }

  private blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(blob);
      reader.onload = () => {
        const base64 = reader.result as string;
        resolve(base64.split(',')[1]);
      };
      reader.onerror = reject;
    });
  }

  // Imię pobieramy z konta MSAL (Entra ID). `name` ma format "Nazwisko, X. (Imię)" — bierzemy
  // tekst z nawiasów; gdy go brak, korzystamy z claimu `given_name`, a w ostateczności z pierwszego
  // członu nazwy. Aktywne konto ustawia App po zalogowaniu.
  private resolveFirstName(): string {
    const account: AccountInfo | null =
      this.msal.instance.getActiveAccount() ?? this.msal.instance.getAllAccounts()[0] ?? null;
    if (!account) return '';

    const fullName = account.name?.trim() ?? '';
    const parenthesized = fullName.match(/\(([^)]+)\)/)?.[1]?.trim();
    if (parenthesized) return parenthesized;

    const givenName = (account.idTokenClaims as { given_name?: unknown } | undefined)?.given_name;
    if (typeof givenName === 'string' && givenName.trim()) return givenName.trim();

    return fullName ? fullName.split(' ')[0] : '';
  }
}

