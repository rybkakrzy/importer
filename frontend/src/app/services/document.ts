import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable } from 'rxjs';

export interface DocumentRun {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
}

export interface DocumentParagraph {
  text: string;
  style: string;
  runs: DocumentRun[];
}

export interface DocumentUploadResponse {
  success: boolean;
  message: string;
  fileName?: string;
  content?: string;
  errorDetails?: string;
}

@Injectable({
  providedIn: 'root',
})
export class DocumentService {
  private apiUrl = 'http://localhost:5190/api/Document';

  constructor(private http: HttpClient) {}

  uploadDocument(file: File): Observable<DocumentUploadResponse> {
    const headers = new HttpHeaders({
      'Content-Type': 'application/octet-stream',
      'X-File-Name': encodeURIComponent(file.name)
    });

    return this.http.post<DocumentUploadResponse>(`${this.apiUrl}/upload`, file, { headers });
  }

  saveDocument(content: string, fileName: string): Observable<Blob> {
    return this.http.post(`${this.apiUrl}/save`, 
      { content, fileName }, 
      { responseType: 'blob' }
    );
  }
}
