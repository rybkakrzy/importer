import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable } from 'rxjs';

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

export interface DocxOpenResponse {
  success: boolean;
  message: string;
  fileName?: string;
  html?: string;
}

export interface DocxSaveRequest {
  fileName?: string;
  html: string;
}

@Injectable({
  providedIn: 'root',
})
export class FileUpload {
  private zipApiUrl = 'http://localhost:5190/api/FileUpload';
  private docxApiUrl = 'http://localhost:5190/api/DocxEditor';

  constructor(private http: HttpClient) {}

  uploadZipFile(file: File): Observable<ZipUploadResponse> {
    const headers = new HttpHeaders({
      'Content-Type': 'application/octet-stream',
      'X-File-Name': encodeURIComponent(file.name)
    });

    return this.http.post<ZipUploadResponse>(`${this.zipApiUrl}/upload`, file, { headers });
  }

  openDocx(file: File): Observable<DocxOpenResponse> {
    const headers = new HttpHeaders({
      'Content-Type': 'application/octet-stream',
      'X-File-Name': encodeURIComponent(file.name)
    });

    return this.http.post<DocxOpenResponse>(`${this.docxApiUrl}/open`, file, { headers });
  }

  saveDocx(request: DocxSaveRequest): Observable<Blob> {
    return this.http.post(`${this.docxApiUrl}/save`, request, { responseType: 'blob' });
  }
}
