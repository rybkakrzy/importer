import { CommonModule } from '@angular/common';
import { Component, SecurityContext } from '@angular/core';
import { DomSanitizer } from '@angular/platform-browser';
import { FileUpload as FileUploadService } from '../../services/file-upload';

@Component({
  selector: 'app-docx-editor',
  imports: [CommonModule],
  templateUrl: './docx-editor.html',
  styleUrl: './docx-editor.css',
})
export class DocxEditorComponent {
  selectedDocx: File | null = null;
  openedFileName: string | null = null;
  editorContent = '<p><br></p>';
  isOpening = false;
  isSaving = false;
  errorMessage: string | null = null;
  infoMessage: string | null = null;

  constructor(
    private fileUploadService: FileUploadService,
    private sanitizer: DomSanitizer
  ) {}

  onDocxSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0] ?? null;
    if (file && !file.name.toLowerCase().endsWith('.docx')) {
      this.selectedDocx = null;
      this.errorMessage = 'Można wybrać tylko plik DOCX.';
      this.infoMessage = null;
      return;
    }

    this.selectedDocx = file;
    this.errorMessage = null;
    this.infoMessage = null;
  }

  onEditorInput(event: Event): void {
    this.editorContent = this.sanitizer.sanitize(
      SecurityContext.HTML,
      (event.target as HTMLElement).innerHTML
    ) ?? '<p><br></p>';
  }

  openDocx(): void {
    if (!this.selectedDocx) {
      this.errorMessage = 'Wybierz plik DOCX.';
      return;
    }

    if (!this.selectedDocx.name.toLowerCase().endsWith('.docx')) {
      this.errorMessage = 'Można otworzyć tylko plik DOCX.';
      return;
    }

    this.isOpening = true;
    this.errorMessage = null;
    this.infoMessage = null;

    this.fileUploadService.openDocx(this.selectedDocx).subscribe({
      next: (response) => {
        if (response.success && response.html) {
          this.editorContent = this.sanitizer.sanitize(SecurityContext.HTML, response.html) ?? '<p><br></p>';
          this.openedFileName = response.fileName ?? this.selectedDocx?.name ?? null;
          this.infoMessage = response.message;
        } else {
          this.errorMessage = response.message || 'Nie udało się odczytać pliku.';
        }
        this.isOpening = false;
      },
      error: (error) => {
        this.errorMessage = error.error?.message || 'Błąd podczas odczytu DOCX.';
        this.isOpening = false;
      }
    });
  }

  saveDocx(): void {
    if (!this.editorContent.trim()) {
      this.errorMessage = 'Brak treści do zapisu.';
      return;
    }

    this.isSaving = true;
    this.errorMessage = null;
    this.infoMessage = null;

    this.fileUploadService.saveDocx({
      fileName: this.openedFileName ?? this.selectedDocx?.name ?? 'dokument.docx',
      html: this.editorContent
    }).subscribe({
      next: (blob) => {
        const rawName = this.openedFileName ?? this.selectedDocx?.name ?? 'dokument.docx';
        const suggestedName = rawName.toLowerCase().endsWith('.docx') ? rawName : `${rawName}.docx`;
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = suggestedName;
        link.click();
        URL.revokeObjectURL(url);
        this.infoMessage = 'Plik DOCX został zapisany.';
        this.isSaving = false;
      },
      error: (error) => {
        this.errorMessage = error.error?.message || 'Błąd podczas zapisu DOCX.';
        this.isSaving = false;
      }
    });
  }
}
