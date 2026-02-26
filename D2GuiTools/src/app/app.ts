import { Component } from '@angular/core';
import { DocumentEditorComponent } from './components/document-editor/document-editor';
import { OfflineBannerComponent } from './components/offline-banner/offline-banner';

@Component({
  selector: 'd2-root',
  imports: [DocumentEditorComponent, OfflineBannerComponent],
  template: `
    <d2-offline-banner />
    <d2-document-editor />
  `,
  styles: [`
    :host {
      display: flex;
      flex-direction: column;
      height: 100vh;
    }
    d2-document-editor {
      flex: 1;
      min-height: 0;
    }
  `]
})
export class App {}
