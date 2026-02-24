import { Component } from '@angular/core';
import { DocumentEditorComponent } from './components/document-editor/document-editor';
import { UploadSessionsComponent } from './components/upload-sessions/upload-sessions';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [DocumentEditorComponent, UploadSessionsComponent],
  template: `
    <div class="app-layout">
      <app-document-editor class="editor-area" />
      <aside class="sessions-panel">
        <app-upload-sessions />
      </aside>
    </div>
  `,
  styles: [`
    :host { display: block; height: 100vh; }
    .app-layout { display: flex; height: 100%; }
    .editor-area { flex: 1; overflow: auto; }
    .sessions-panel { width: 420px; border-left: 1px solid #e0e0e0; overflow-y: auto; background: #fafafa; }
  `]
})
export class App {}
