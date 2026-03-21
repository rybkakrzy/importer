import { Injectable, inject } from '@angular/core';
import { Router } from '@angular/router';

export const MIME_PDF = 'application/pdf';
export const MIME_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
export const MIME_DOC = 'application/msword';

/**
 * Centralny serwis nawigacji dokumentów.
 * Na podstawie mimeType decyduje czy otworzyć edytor Word czy podgląd PDF.
 */
@Injectable({ providedIn: 'root' })
export class DocumentNavigationService {
  private router = inject(Router);

  navigateToDocument(masterId: string, mimeType: string): void {
    if (this.isPdf(mimeType)) {
      this.router.navigate(['/viewer'], { queryParams: { masterId } });
    } else {
      this.router.navigate(['/editor'], { queryParams: { masterId } });
    }
  }

  isPdf(mimeType: string): boolean {
    return mimeType === MIME_PDF || mimeType.toLowerCase().includes('pdf');
  }

  isWordDocument(mimeType: string): boolean {
    return mimeType === MIME_DOCX || mimeType === MIME_DOC;
  }
}
