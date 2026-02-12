import { Component, signal } from '@angular/core';
import { FileUploadComponent } from './components/file-upload/file-upload';
import { WordEditorComponent } from './components/word-editor/word-editor';

@Component({
  selector: 'app-root',
  imports: [FileUploadComponent, WordEditorComponent],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App {
  protected readonly title = signal('Importer Parametryzacji');
}
