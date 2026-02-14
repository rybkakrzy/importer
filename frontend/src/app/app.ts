import { Component } from '@angular/core';
import { DocumentEditorComponent } from './components/document-editor/document-editor';

@Component({
  selector: 'app-root',
  imports: [DocumentEditorComponent],
  template: '<app-document-editor />',
  styles: [':host { display: block; height: 100vh; }']
})
export class App {}
