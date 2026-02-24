import { Injectable, inject } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Observable, throwError } from 'rxjs';
import { catchError } from 'rxjs/operators';
import { 
  DocumentContent, 
  SaveDocumentRequest, 
  SignDocumentRequest,
  DocumentTemplate, 
  ImageUploadResponse,
  DigitalSignatureInfo
} from '../models/document.model';
import { ApiConfigService } from '../core/services/api-config.service';

/**
 * Serwis do komunikacji z API dokumentów
 */
@Injectable({
  providedIn: 'root'
})
export class DocumentService {
  private http = inject(HttpClient);
  private apiConfig = inject(ApiConfigService);

  private get apiUrl(): string {
    return this.apiConfig.documentUrl;
  }

  /**
   * Otwiera dokument DOCX
   */
  openDocument(file: File): Observable<DocumentContent> {
    const formData = new FormData();
    formData.append('file', file);

    return this.http.post<DocumentContent>(`${this.apiUrl}/open`, formData)
      .pipe(catchError(this.handleError));
  }

  /**
   * Zapisuje dokument jako DOCX
   */
  saveDocument(request: SaveDocumentRequest): Observable<Blob> {
    return this.http.post(`${this.apiUrl}/save`, request, {
      responseType: 'blob'
    }).pipe(catchError(this.handleError));
  }

  /**
   * Tworzy nowy pusty dokument
   */
  newDocument(): Observable<DocumentContent> {
    return this.http.get<DocumentContent>(`${this.apiUrl}/new`)
      .pipe(catchError(this.handleError));
  }

  /**
   * Pobiera listę szablonów
   */
  getTemplates(): Observable<DocumentTemplate[]> {
    return this.http.get<DocumentTemplate[]>(`${this.apiUrl}/templates`)
      .pipe(catchError(this.handleError));
  }

  /**
   * Pobiera szablon dokumentu
   */
  getTemplate(templateId: string): Observable<DocumentContent> {
    return this.http.get<DocumentContent>(`${this.apiUrl}/templates/${templateId}`)
      .pipe(catchError(this.handleError));
  }

  /**
   * Wgrywa obraz
   */
  uploadImage(file: File): Observable<ImageUploadResponse> {
    const formData = new FormData();
    formData.append('file', file);

    return this.http.post<ImageUploadResponse>(`${this.apiUrl}/upload-image`, formData)
      .pipe(catchError(this.handleError));
  }

  /**
   * Pobiera plik jako blob i uruchamia pobieranie
   */
  downloadDocument(request: SaveDocumentRequest, filename: string): void {
    this.saveDocument(request).subscribe({
      next: (blob) => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename.endsWith('.docx') ? filename : `${filename}.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
      },
      error: (error) => {
        console.error('Błąd podczas pobierania dokumentu:', error);
      }
    });
  }

  /**
   * Podpisuje dokument certyfikatem cyfrowym i pobiera podpisany plik
   */
  signDocument(request: SignDocumentRequest): Observable<Blob> {
    return this.http.post(`${this.apiUrl}/sign`, request, {
      responseType: 'blob'
    }).pipe(catchError(this.handleError));
  }

  /**
   * Podpisuje i pobiera dokument
   */
  downloadSignedDocument(request: SignDocumentRequest, filename: string): void {
    this.signDocument(request).subscribe({
      next: (blob) => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename.endsWith('.docx') ? filename : `${filename}.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
      },
      error: (error) => {
        console.error('Błąd podczas podpisywania dokumentu:', error);
      }
    });
  }

  /**
   * Weryfikuje podpisy cyfrowe w pliku DOCX
   */
  verifySignatures(file: File): Observable<DigitalSignatureInfo[]> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<DigitalSignatureInfo[]>(`${this.apiUrl}/verify-signatures`, formData)
      .pipe(catchError(this.handleError));
  }

  /**
   * Obsługa błędów HTTP
   */
  private handleError(error: HttpErrorResponse): Observable<never> {
    let errorMessage = 'Wystąpił nieznany błąd';

    if (error.error instanceof ErrorEvent) {
      // Błąd po stronie klienta
      errorMessage = `Błąd: ${error.error.message}`;
    } else {
      // Błąd po stronie serwera
      if (error.error?.error) {
        errorMessage = error.error.error;
      } else {
        errorMessage = `Błąd serwera: ${error.status} - ${error.statusText}`;
      }
    }

    console.error('DocumentService Error:', errorMessage);
    return throwError(() => new Error(errorMessage));
  }
}
