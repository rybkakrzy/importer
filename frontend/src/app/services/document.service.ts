import { Injectable, inject } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Observable, throwError } from 'rxjs';
import { catchError } from 'rxjs/operators';
import { 
  DocumentContent, 
  SaveDocumentRequest, 
  DocumentTemplate, 
  ImageUploadResponse 
} from '../models/document.model';

/**
 * Serwis do komunikacji z API dokumentów
 */
@Injectable({
  providedIn: 'root'
})
export class DocumentService {
  private http = inject(HttpClient);
  private apiUrl = 'http://localhost:5190/api/document';

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
