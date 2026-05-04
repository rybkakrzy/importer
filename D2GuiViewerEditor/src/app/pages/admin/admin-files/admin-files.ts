import { Component, computed, inject, OnInit, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { AdminService, DocumentWithVersions } from '../../../services/admin.service';

@Component({
  selector: 'd2-admin-files',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './admin-files.html',
  styleUrl: './admin-files.scss'
})
export class AdminFilesComponent implements OnInit {
  private adminService = inject(AdminService);

  private allDocuments = signal<DocumentWithVersions[]>([]);
  filterId   = signal('');
  filterType = signal('');
  filterDate = signal('');
  currentPage = signal(0);
  readonly pageSize = 10;
  isLoading = signal(true);
  error = signal<string | null>(null);
  loadingVersionsFor = signal<string | null>(null);
  downloadingId = signal<string | null>(null);

  private filteredDocuments = computed(() => {
    const id   = this.filterId().toLowerCase().trim();
    const type = this.filterType().toLowerCase().trim();
    const date = this.filterDate().toLowerCase().trim();
    return this.allDocuments().filter(d => {
      if (id   && !d.masterId.toLowerCase().includes(id))                       return false;
      if (type && !d.mimeType.toLowerCase().includes(type))                     return false;
      if (date && !this.formatDate(d.createdAt).toLowerCase().includes(date))   return false;
      return true;
    });
  });

  totalFiltered = computed(() => this.filteredDocuments().length);
  totalPages    = computed(() => Math.max(1, Math.ceil(this.totalFiltered() / this.pageSize)));

  documents = computed(() => {
    const start = this.currentPage() * this.pageSize;
    return this.filteredDocuments().slice(start, start + this.pageSize);
  });

  setFilter(field: 'id' | 'type' | 'date', value: string): void {
    if (field === 'id')   this.filterId.set(value);
    if (field === 'type') this.filterType.set(value);
    if (field === 'date') this.filterDate.set(value);
    this.currentPage.set(0);
  }

  prevPage(): void { if (this.currentPage() > 0) this.currentPage.update(p => p - 1); }
  nextPage(): void { if (this.currentPage() < this.totalPages() - 1) this.currentPage.update(p => p + 1); }
  pageEnd(): number { return Math.min((this.currentPage() + 1) * this.pageSize, this.totalFiltered()); }
  goToPage(value: string): void {
    const n = parseInt(value, 10);
    if (!isNaN(n)) this.currentPage.set(Math.max(0, Math.min(n - 1, this.totalPages() - 1)));
  }

  ngOnInit(): void {
    this.adminService.getAllDocuments().subscribe({
      next: (docs) => {
        this.allDocuments.set(docs.map(d => ({ ...d, versions: [], expanded: false })));
        this.isLoading.set(false);
      },
      error: () => {
        this.error.set('Nie udało się załadować listy plików.');
        this.isLoading.set(false);
      }
    });
  }

  toggleExpand(doc: DocumentWithVersions): void {
    const docs = this.allDocuments();
    const idx = docs.findIndex(d => d.masterId === doc.masterId);
    if (idx === -1) return;

    const current = docs[idx];
    if (!current.expanded && current.versions.length === 0) {
      this.loadingVersionsFor.set(doc.masterId);
      this.adminService.getDocumentVersions(doc.masterId).subscribe({
        next: (versions) => {
          this.updateDoc(doc.masterId, { versions, expanded: true });
          this.loadingVersionsFor.set(null);
        },
        error: () => this.loadingVersionsFor.set(null)
      });
    } else {
      this.updateDoc(doc.masterId, { expanded: !current.expanded });
    }
  }

  downloadActive(doc: DocumentWithVersions, event: Event): void {
    event.stopPropagation();
    this.downloadingId.set(doc.masterId);
    this.adminService.downloadDocumentVersion(doc.masterId, doc.activeVersionId).subscribe({
      next: (blob) => {
        this.triggerDownload(blob, doc.name, doc.mimeType);
        this.downloadingId.set(null);
      },
      error: () => this.downloadingId.set(null)
    });
  }

  downloadVersion(masterId: string, versionId: string, fileName: string, mimeType: string, event: Event): void {
    event.stopPropagation();
    const key = `${masterId}_${versionId}`;
    this.downloadingId.set(key);
    this.adminService.downloadDocumentVersion(masterId, versionId).subscribe({
      next: (blob) => {
        this.triggerDownload(blob, fileName, mimeType);
        this.downloadingId.set(null);
      },
      error: () => this.downloadingId.set(null)
    });
  }

  private triggerDownload(blob: Blob, fileName: string, mimeType: string): void {
    const url = URL.createObjectURL(new Blob([blob], { type: mimeType }));
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 200);
  }

  private updateDoc(masterId: string, patch: Partial<DocumentWithVersions>): void {
    const docs = this.allDocuments();
    const idx = docs.findIndex(d => d.masterId === masterId);
    if (idx === -1) return;
    const updated = [...docs];
    updated[idx] = { ...updated[idx], ...patch };
    this.allDocuments.set(updated);
  }

  friendlyType(mimeType: string): string {
    const map: Record<string, string> = {
      'application/pdf': 'PDF',
      'application/msword': 'Word',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word',
      'application/vnd.ms-excel': 'Excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel',
      'application/vnd.ms-powerpoint': 'PowerPoint',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint',
      'application/zip': 'ZIP',
      'application/json': 'JSON',
      'application/xml': 'XML',
      'text/xml': 'XML',
      'text/plain': 'TXT',
      'text/csv': 'CSV',
      'text/html': 'HTML',
      'image/png': 'PNG',
      'image/jpeg': 'JPG',
      'image/gif': 'GIF',
      'image/svg+xml': 'SVG',
    };
    return map[mimeType] ?? mimeType.split('/')[1]?.toUpperCase() ?? mimeType;
  }

  formatBytes(bytes: number): string {
    if (!bytes || bytes === 0) return '—';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
  }

  formatDate(dateStr: string): string {
    return new Date(dateStr).toLocaleDateString('pl-PL', {
      day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit'
    });
  }
}
