import { Injectable, inject } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable } from 'rxjs';
import { environment } from '../../environments/environment';

export interface UploadDocumentRequest {
  name: string;
  mimeType: string;
  content: string; // Base64 encoded
  createdBy?: string;
}

export interface UploadDocumentResult {
  masterId: string;
  versionId: string;
  fileName: string;
  createdAt: string;
}

export interface SaveDocumentVersionRequest {
  content: string; // Base64 encoded
  createdBy?: string;
}

export interface SaveDocumentVersionResult {
  versionId: string;
  versionNumber: number;
  createdAt: string;
}

export interface UpdateDocumentVersionResult {
  versionId: string;
  versionNumber: number;
  sizeInBytes: number;
  modifiedAt: string;
}

export interface DocumentDto {
  masterId: string;
  name: string;
  mimeType: string;
  createdAt: string;
  activeVersionId: string;
  content: string; // Base64 encoded
  versionNumber: number;
}

export interface DocumentMetadataDto {
  masterId: string;
  mimeType: string;
  returnUrl: string | null;
  classification: string | null;
}

export interface DocumentVersionDto {
  versionId: string;
  versionNumber: number;
  createdAt: string;
  createdBy: string;
  isActive: boolean;
  sizeInBytes: number;
}

export interface RestoreVersionResponse {
  message: string;
  versionId: string;
}

export type DeliveryStatus =
  | 'Pending'
  | 'Sending'
  | 'RetryScheduled'
  | 'Sent'
  | 'FailedPermanently'
  | 'DeadLettered';

export interface FinishAndSendResult {
  deliveryId: string;
  status: DeliveryStatus;
  statusUrl: string;
}

export interface DeliveryStatusDto {
  deliveryId: string;
  documentId: string;
  status: DeliveryStatus;
  attemptCount: number;
  lastAttemptAt: string | null;
  nextAttemptAt: string | null;
  lastError: string | null;
  updatedAt: string;
}

/**
 * Serwis do zarządzania dokumentami z wersjonowaniem
 */
@Injectable({
  providedIn: 'root'
})
export class DocumentStorageService {
  private readonly http = inject(HttpClient);
  private readonly apiUrl = `${environment.apiUrl}/documentstorage`;

  /**
   * Upload nowego dokumentu do systemu
   * @param request - Dane dokumentu (nazwa, typ MIME, zawartość w Base64)
   * @returns Observable z GUID mastera i GUID pierwszej wersji
   */
  uploadDocument(request: UploadDocumentRequest): Observable<UploadDocumentResult> {
    return this.http.post<UploadDocumentResult>(`${this.apiUrl}/upload`, request);
  }

  /**
   * Zapisanie nowej wersji dokumentu (każde zapisanie z GUI)
   * @param masterId - GUID mastera dokumentu
   * @param request - Nowa zawartość dokumentu w Base64
   * @returns Observable z GUID nowej wersji
   */
  saveDocumentVersion(
    masterId: string,
    request: SaveDocumentVersionRequest
  ): Observable<SaveDocumentVersionResult> {
    return this.http.post<SaveDocumentVersionResult>(
      `${this.apiUrl}/${masterId}/save`,
      request
    );
  }

  /**
   * Nadpisanie istniejącej wersji w miejscu (auto-save edytora).
   * Podmienia plik w GCS pod tym samym versionId — nie tworzy nowych wersji.
   * @param masterId - GUID mastera dokumentu
   * @param versionId - GUID wersji edytowalnej do nadpisania
   * @param request - Nowa zawartość dokumentu w Base64
   */
  updateDocumentVersion(
    masterId: string,
    versionId: string,
    request: SaveDocumentVersionRequest
  ): Observable<UpdateDocumentVersionResult> {
    return this.http.put<UpdateDocumentVersionResult>(
      `${this.apiUrl}/${masterId}/versions/${versionId}`,
      request
    );
  }

  /**
   * Pobranie aktywnej wersji dokumentu
   * @param masterId - GUID mastera dokumentu
   * @returns Observable z pełnymi danymi dokumentu (włącznie z contentem)
   */
  getDocument(masterId: string): Observable<DocumentDto> {
    return this.http.get<DocumentDto>(`${this.apiUrl}/${masterId}`);
  }

  /**
   * Pobranie metadanych dokumentu (mimeType, returnUrl, classification) — lekkie, bez contentu.
   * @param masterId - GUID mastera dokumentu
   */
  getDocumentMetadata(masterId: string): Observable<DocumentMetadataDto> {
    return this.http.get<DocumentMetadataDto>(`${this.apiUrl}/${masterId}/metadata`);
  }

  /**
   * Pobranie zawartości wersji BAZOWEJ (v1, oryginał) jako Blob — tryb read-only (Krok 2).
   * @param masterId - GUID mastera dokumentu
   */
  downloadBaseVersion(masterId: string): Observable<Blob> {
    return this.http.get(`${this.apiUrl}/${masterId}/download`, { responseType: 'blob' });
  }

  /**
   * Pobranie zawartości KONKRETNEJ wersji jako Blob — tryb edycji (Krok 3, wersja edytowalna).
   * @param masterId - GUID mastera dokumentu
   * @param versionId - GUID wersji
   */
  downloadVersion(masterId: string, versionId: string): Observable<Blob> {
    return this.http.get(`${this.apiUrl}/${masterId}/versions/${versionId}/download`, { responseType: 'blob' });
  }

  /**
   * Pobranie listy wszystkich wersji dokumentu (historia)
   * @param masterId - GUID mastera dokumentu
   * @returns Observable z listą wersji (bez contentu, tylko metadane)
   */
  getDocumentVersions(masterId: string): Observable<DocumentVersionDto[]> {
    return this.http.get<DocumentVersionDto[]>(`${this.apiUrl}/${masterId}/versions`);
  }

  /**
   * Przywrócenie wybranej wersji dokumentu (cofnięcie edycji)
   * @param masterId - GUID mastera dokumentu
   * @param versionId - GUID wersji do przywrócenia
   * @returns Observable z potwierdzeniem operacji
   */
  restoreDocumentVersion(
    masterId: string,
    versionId: string
  ): Observable<RestoreVersionResponse> {
    return this.http.post<RestoreVersionResponse>(
      `${this.apiUrl}/${masterId}/restore/${versionId}`,
      {}
    );
  }

  /**
   * "Zakończ i wyślij": utrwala stan edytora i tworzy zadanie asynchronicznej wysyłki na returnUrl.
   * Idempotentne — ponowne kliknięcie zwraca to samo zadanie.
   * @param masterId - GUID mastera dokumentu
   * @param versionId - GUID wersji edytowalnej
   * @param request - Aktualna zawartość edytora w Base64
   */
  finishAndSend(
    masterId: string,
    versionId: string,
    request: SaveDocumentVersionRequest
  ): Observable<FinishAndSendResult> {
    return this.http.post<FinishAndSendResult>(
      `${this.apiUrl}/${masterId}/versions/${versionId}/finish`,
      request
    );
  }

  /**
   * Status zadania wysyłki (polling).
   * @param deliveryId - GUID zadania wysyłki
   */
  getDeliveryStatus(deliveryId: string): Observable<DeliveryStatusDto> {
    return this.http.get<DeliveryStatusDto>(`${this.apiUrl}/deliveries/${deliveryId}`);
  }

  /**
   * Konwersja pliku na Base64 (helper dla uploadu)
   * @param file - Plik z input[type=file]
   * @returns Promise z zawartością w Base64
   */
  fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        const base64 = reader.result as string;
        // Usuń prefix "data:*/*;base64,"
        const base64Content = base64.split(',')[1];
        resolve(base64Content);
      };
      reader.onerror = error => reject(error);
    });
  }

  /**
   * Konwersja Base64 na Blob (helper dla pobierania)
   * @param base64 - Zawartość w Base64
   * @param mimeType - Typ MIME
   * @returns Blob
   */
  base64ToBlob(base64: string, mimeType: string): Blob {
    const byteCharacters = atob(base64);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: mimeType });
  }

  /**
   * Download dokumentu jako plik
   * @param doc - Dokument z API
   */
  downloadDocument(doc: DocumentDto): void {
    const blob = this.base64ToBlob(doc.content, doc.mimeType);
    const url = window.URL.createObjectURL(blob);
    const link = window.document.createElement('a');
    link.href = url;
    link.download = doc.name;
    link.click();
    window.URL.revokeObjectURL(url);
  }

  /**
   * Formatowanie rozmiaru pliku (helper dla UI)
   * @param bytes - Rozmiar w bajtach
   * @returns Sformatowany string (np. "1.5 MB")
   */
  formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
  }
}
