import { Injectable, inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { ApiConfigService } from '../core/services/api-config.service';
import { UploadSession } from '../models/upload-session.model';

@Injectable({
  providedIn: 'root',
})
export class UploadSessionsService {
  private http = inject(HttpClient);
  private apiConfig = inject(ApiConfigService);

  getSessions(): Observable<UploadSession[]> {
    return this.http.get<UploadSession[]>(this.apiConfig.uploadSessionsUrl);
  }
}
