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

@Injectable({
  providedIn: 'root',
})
export class FileUpload {
  private apiUrl = 'http://localhost:5190/api/FileUpload';

  constructor(private http: HttpClient) {}

  uploadZipFile(file: File): Observable<ZipUploadResponse> {
    const headers = new HttpHeaders({
      'Content-Type': 'application/octet-stream',
      'X-File-Name': encodeURIComponent(file.name)
    });

    return this.http.post<ZipUploadResponse>(`${this.apiUrl}/upload`, file, { headers });
  }
}

