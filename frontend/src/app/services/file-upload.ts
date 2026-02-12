import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable } from 'rxjs';
import { environment } from '../../environments/environment';

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
  private apiUrl = `${environment.apiUrl}/api/FileUpload`;

  constructor(private http: HttpClient) {}

  uploadZipFile(file: File): Observable<ZipUploadResponse> {
    const headers = new HttpHeaders({
      'Content-Type': 'application/octet-stream',
      'X-File-Name': encodeURIComponent(file.name)
    });

    return this.http.post<ZipUploadResponse>(`${this.apiUrl}/upload`, file, { headers });
  }
}

