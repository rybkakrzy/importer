import { Component, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DocumentService, DocumentParagraph, DocumentRun } from '../../services/document';

@Component({
  selector: 'app-word-editor',
  imports: [CommonModule],
  templateUrl: './word-editor.html',
  styleUrl: './word-editor.css',
})
export class WordEditorComponent {
  selectedFile = signal<File | null>(null);
  isUploading = signal(false);
  errorMessage = signal<string | null>(null);
  documentContent = signal<DocumentParagraph[]>([]);
  fileName = signal<string>('');
  isSaving = signal(false);

  constructor(private documentService: DocumentService) {}

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.selectedFile.set(input.files[0]);
      this.errorMessage.set(null);
    }
  }

  uploadFile(): void {
    const file = this.selectedFile();
    if (!file) {
      this.errorMessage.set('Please select a DOCX file');
      return;
    }

    if (!file.name.toLowerCase().endsWith('.docx')) {
      this.errorMessage.set('Selected file must be a DOCX document');
      return;
    }

    this.isUploading.set(true);
    this.errorMessage.set(null);

    this.documentService.uploadDocument(file).subscribe({
      next: (response) => {
        if (response.success && response.content) {
          try {
            const content = JSON.parse(response.content);
            // Convert property names from PascalCase to camelCase
            const normalizedContent = content.map((para: any) => ({
              text: para.Text || para.text || '',
              style: para.Style || para.style || 'normal',
              runs: (para.Runs || para.runs || []).map((run: any) => ({
                text: run.Text || run.text || '',
                bold: run.Bold ?? run.bold ?? false,
                italic: run.Italic ?? run.italic ?? false,
                underline: run.Underline ?? run.underline ?? false
              }))
            }));
            this.documentContent.set(normalizedContent);
            this.fileName.set(response.fileName || 'document.docx');
          } catch (e) {
            this.errorMessage.set('Error parsing document content');
          }
        } else {
          this.errorMessage.set(response.message || 'Error processing document');
        }
        this.isUploading.set(false);
      },
      error: (error) => {
        this.errorMessage.set(error.message || 'Error uploading document');
        this.isUploading.set(false);
        console.error('Upload error:', error);
      }
    });
  }

  onContentChange(event: Event, paragraphIndex: number, runIndex: number): void {
    const element = event.target as HTMLElement;
    const newText = element.textContent || '';
    
    const content = [...this.documentContent()];
    if (content[paragraphIndex] && content[paragraphIndex].runs[runIndex]) {
      content[paragraphIndex].runs[runIndex].text = newText;
      content[paragraphIndex].text = content[paragraphIndex].runs.map(r => r.text).join('');
      this.documentContent.set(content);
    }
  }

  toggleFormatting(paragraphIndex: number, runIndex: number, format: 'bold' | 'italic' | 'underline'): void {
    const content = [...this.documentContent()];
    if (content[paragraphIndex] && content[paragraphIndex].runs[runIndex]) {
      content[paragraphIndex].runs[runIndex][format] = !content[paragraphIndex].runs[runIndex][format];
      this.documentContent.set(content);
    }
  }

  changeStyle(paragraphIndex: number, style: string): void {
    const content = [...this.documentContent()];
    if (content[paragraphIndex]) {
      content[paragraphIndex].style = style;
      this.documentContent.set(content);
    }
  }

  addParagraph(): void {
    const content = [...this.documentContent()];
    content.push({
      text: '',
      style: 'normal',
      runs: [{ text: '', bold: false, italic: false, underline: false }]
    });
    this.documentContent.set(content);
  }

  deleteParagraph(index: number): void {
    const content = [...this.documentContent()];
    content.splice(index, 1);
    this.documentContent.set(content);
  }

  saveDocument(): void {
    const content = this.documentContent();
    if (content.length === 0) {
      this.errorMessage.set('No content to save');
      return;
    }

    this.isSaving.set(true);
    const contentJson = JSON.stringify(content);
    const fileName = this.fileName();

    this.documentService.saveDocument(contentJson, fileName).subscribe({
      next: (blob) => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        this.isSaving.set(false);
      },
      error: (error) => {
        this.errorMessage.set('Error saving document');
        this.isSaving.set(false);
        console.error('Save error:', error);
      }
    });
  }

  clearDocument(): void {
    this.documentContent.set([]);
    this.fileName.set('');
    this.selectedFile.set(null);
    this.errorMessage.set(null);
  }
}
