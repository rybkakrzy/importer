import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { UploadSessionsService } from '../../services/upload-sessions.service';
import { UploadSession } from '../../models/upload-session.model';

@Component({
  selector: 'app-upload-sessions',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './upload-sessions.html',
  styleUrl: './upload-sessions.scss',
})
export class UploadSessionsComponent implements OnInit, OnDestroy {
  sessions: UploadSession[] = [];
  isLoading = false;
  errorMessage: string | null = null;
  private refreshInterval?: ReturnType<typeof setInterval>;

  constructor(private uploadSessionsService: UploadSessionsService) {}

  ngOnInit(): void {
    this.loadSessions();
    this.refreshInterval = setInterval(() => this.loadSessions(), 5000);
  }

  ngOnDestroy(): void {
    if (this.refreshInterval) {
      clearInterval(this.refreshInterval);
    }
  }

  loadSessions(): void {
    this.isLoading = true;
    this.uploadSessionsService.getSessions().subscribe({
      next: (sessions) => {
        this.sessions = sessions;
        this.isLoading = false;
        this.errorMessage = null;
      },
      error: () => {
        this.errorMessage = 'Nie można załadować listy sesji';
        this.isLoading = false;
      }
    });
  }

  formatFileSize(bytes?: number): string {
    if (bytes == null) return '—';
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
  }

  getStatusClass(status: string): string {
    switch (status) {
      case 'Completed': return 'status-completed';
      case 'Processing': return 'status-processing';
      case 'Failed': return 'status-failed';
      default: return '';
    }
  }

  getStatusLabel(status: string): string {
    switch (status) {
      case 'Completed': return '✅ Zakończono';
      case 'Processing': return '⏳ Przetwarzanie';
      case 'Failed': return '❌ Błąd';
      default: return status;
    }
  }
}
