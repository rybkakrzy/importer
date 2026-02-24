import { Component } from '@angular/core';
import { DocumentEditorComponent } from './components/document-editor/document-editor';

@Component({
  selector: 'd2-root',
  imports: [DocumentEditorComponent],
  template: '<d2-document-editor />',
  styles: [':host { display: block; height: 100vh; }']
})
export class App {}
