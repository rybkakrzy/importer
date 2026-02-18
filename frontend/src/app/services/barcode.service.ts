import { Injectable, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { catchError } from 'rxjs/operators';

/** Request do generowania kodu */
export interface BarcodeRequest {
  content: string;
  barcodeType: string;
  width?: number;
  height?: number;
}

/** Odpowiedź z wygenerowanym kodem */
export interface BarcodeResponse {
  base64Image: string;
  contentType: string;
  barcodeType: string;
}

/**
 * Serwis do generowania kodów kreskowych i QR
 */
@Injectable({
  providedIn: 'root'
})
export class BarcodeService {
  private http = inject(HttpClient);
  private apiUrl = 'http://localhost:5190/api/barcode';

  /** Lista obsługiwanych typów kodów */
  readonly barcodeTypes: { value: string; label: string; is2D: boolean }[] = [
    { value: 'QRCode', label: 'QR Code', is2D: true },
    { value: 'Code128', label: 'Code 128', is2D: false },
    { value: 'EAN13', label: 'EAN-13', is2D: false },
    { value: 'UPCA', label: 'UPC-A', is2D: false },
    { value: 'EAN8', label: 'EAN-8', is2D: false },
    { value: 'Interleaved2of5', label: 'Interleaved 2 of 5', is2D: false },
    { value: 'Code39', label: 'Code 39', is2D: false },
    { value: 'AztecCode', label: 'Aztec Code', is2D: true },
    { value: 'Datamatrix', label: 'Data Matrix', is2D: true },
    { value: 'PDF417', label: 'PDF417', is2D: true },
    { value: 'MicroPDF', label: 'Micro PDF', is2D: true },
    { value: 'Codabar', label: 'Codabar', is2D: false },
    { value: 'Code93', label: 'Code 93', is2D: false },
    { value: 'Maxicode', label: 'Maxicode', is2D: true }
  ];

  /**
   * Generuje kod kreskowy / QR i zwraca jako base64
   */
  generateBarcode(request: BarcodeRequest): Observable<BarcodeResponse> {
    return this.http.post<BarcodeResponse>(`${this.apiUrl}/generate`, {
      content: request.content,
      barcodeType: request.barcodeType,
      width: request.width || 300,
      height: request.height || 300
    }).pipe(
      catchError(error => {
        console.error('Błąd generowania kodu:', error);
        throw error;
      })
    );
  }
}
