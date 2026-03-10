import { Injectable } from '@angular/core';
import { environment } from '../../../environments/environment';

/**
 * Centralna konfiguracja API — single source of truth dla URL-i backendu
 */
@Injectable({
  providedIn: 'root'
})
export class ApiConfigService {
  readonly baseUrl = environment.apiUrl;

  /** API dokumentów */
  get documentUrl(): string {
    return `${this.baseUrl}/document`;
  }

  /** API kodów kreskowych */
  get barcodeUrl(): string {
    return `${this.baseUrl}/barcode`;
  }

  /** API uploadu plików */
  get fileUploadUrl(): string {
    return `${this.baseUrl}/FileUpload`;
  }
}
