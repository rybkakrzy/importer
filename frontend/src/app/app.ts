import { Component, signal } from '@angular/core';
import { FileUploadComponent } from './components/file-upload/file-upload';
import { DocxEditorComponent } from './components/docx-editor/docx-editor';

@Component({
  selector: 'app-root',
  imports: [FileUploadComponent, DocxEditorComponent],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App {
  protected readonly title = signal('Importer Parametryzacji');
}
