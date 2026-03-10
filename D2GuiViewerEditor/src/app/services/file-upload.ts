import { Injectable, inject } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable } from 'rxjs';
import { ApiConfigService } from '../core/services/api-config.service';

export interface ZipFileEntry {
  fileName: string;
  fullPath: string;
  fileSize: number;
  compressedSize: number;
  fileType: string;
  lastModified: string;
  content?: string;
}

export interface ZipUploadResponse {
  success: boolean;
  message: string;
  fileName?: string;
  fileSize?: number;
  filesCount?: number;
  files?: ZipFileEntry[];
  processedAt?: string;
  errorCode?: string;
  errorDetails?: string;
}

@Injectable({
  providedIn: 'root',
})
export class FileUpload {
  private http = inject(HttpClient);
  private apiConfig = inject(ApiConfigService);

  private get apiUrl(): string {
    return this.apiConfig.fileUploadUrl;
  }

  uploadZipFile(file: File): Observable<ZipUploadResponse> {
    const headers = new HttpHeaders({
      'Content-Type': 'application/octet-stream',
      'X-File-Name': encodeURIComponent(file.name)
    });

    return this.http.post<ZipUploadResponse>(`${this.apiUrl}/upload`, file, { headers });
  }
}

