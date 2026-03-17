import { Component, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Router } from '@angular/router';
import { switchMap, from } from 'rxjs';
import { DocumentService } from '../../services/document.service';
import { DocumentStorageService } from '../../services/document-storage.service';

@Component({
  selector: 'd2-dashboard',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class DashboardComponent {
  private router = inject(Router);
  private documentService = inject(DocumentService);
  private documentStorageService = inject(DocumentStorageService);

  isLoading = signal(false);
  errorMessage = signal<string | null>(null);
  readonly currentYear = new Date().getFullYear();

  newDocument(): void {
    this.isLoading.set(true);
    this.errorMessage.set(null);

    this.documentService.newDocument().pipe(
      switchMap(content =>
        this.documentService.saveDocument({ html: content.html, metadata: content.metadata }).pipe(
          switchMap(blob =>
            from(this.blobToBase64(blob)).pipe(
              switchMap(base64 =>
                this.documentStorageService.uploadDocument({
                  name: 'Nowy dokument.docx',
                  mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                  content: base64
                })
              )
            )
          )
        )
      )
    ).subscribe({
      next: (result) => {
        this.router.navigate(['/editor'], { queryParams: { masterId: result.masterId } });
      },
      error: () => {
        this.isLoading.set(false);
        this.errorMessage.set('Błąd podczas tworzenia dokumentu. Spróbuj ponownie.');
      }
    });
  }

  openFile(): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.docx,.pdf';

    input.onchange = async (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;

      if (file.name.toLowerCase().endsWith('.pdf')) {
        this.router.navigate(['/pdf-maintenance']);
        return;
      }

      this.isLoading.set(true);
      this.errorMessage.set(null);

      try {
        const base64 = await this.documentStorageService.fileToBase64(file);
        this.documentStorageService.uploadDocument({
          name: file.name,
          mimeType: file.type || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          content: base64
        }).subscribe({
          next: (result) => {
            this.router.navigate(['/editor'], { queryParams: { masterId: result.masterId } });
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
}
