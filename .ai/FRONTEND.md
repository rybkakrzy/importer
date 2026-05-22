# Frontend — D2GuiViewerEditor

## Tech

Angular 20, standalone components, **Signals** for state, RxJS only for HTTP. No NgModule, no NgRx. Vitest 3.1.1 for tests.

---

## App structure

```
src/app/
├── app.ts              — root component, router-outlet + offline banner
├── app.routes.ts       — route definitions
├── app.config.ts       — provideRouter, provideHttpClient, error handler
│
├── components/
│   ├── document-editor/        — main DOCX editor (smart, 3326 lines)
│   ├── wysiwyg-editor/         — contenteditable WYSIWYG
│   ├── editor-toolbar/         — formatting toolbar
│   ├── barcode-dialog/         — barcode/QR dialog
│   ├── file-upload/            — file picker
│   ├── offline-banner/         — connection status banner
│   └── ruler/                  — page ruler (margins indicator)
│
├── pages/
│   ├── dashboard/              — home / file browser
│   ├── document-editor/        — page wrapping document-editor component
│   ├── pdf-viewer/             — lazy, pdfjs-dist
│   ├── pdf-maintenance/        — PDF → DOCX tools
│   └── admin/
│       ├── admin-shell/        — admin layout (lazy)
│       └── admin-files/        — file management (lazy)
│
├── services/
│   ├── document.service.ts         — stateless file API (open/save/sign)
│   ├── document-storage.service.ts — storage CRUD (upload/download/versions)
│   ├── barcode.service.ts          — barcode generation
│   ├── file-upload.ts              — file helper
│   └── admin.service.ts            — admin operations
│
├── core/
│   ├── services/
│   │   ├── api-config.service.ts           — reads apiUrl from environment
│   │   ├── build-info.service.ts           — version/env/date
│   │   ├── connection-status.service.ts    — online/offline detection
│   │   ├── document-navigation.service.ts  — routing helpers
│   │   └── notification.service.ts         — toast/alert
│   ├── interceptors/http-error.interceptor.ts
│   └── error-handling/global-error-handler.ts
│
└── models/document.model.ts
```

---

## Routes

| Path | Component | Lazy? | Purpose |
|---|---|---|---|
| `/` | DashboardComponent | no | Home / file browser |
| `/editor` | DocumentEditorComponent (page) | no | DOCX editor |
| `/viewer` | PdfViewerComponent | yes | PDF viewer |
| `/pdf-maintenance` | PdfMaintenanceComponent | no | PDF conversion |
| `/admin` | AdminShellComponent | yes | Admin |
| `/admin/files` | AdminFilesComponent | yes | File management |
| `**` | redirect `/` | — | — |

---

## DocumentEditorComponent — key signals

File: `components/document-editor/document-editor.ts`

```typescript
documentContent     = signal<string>('')              // HTML shown in WYSIWYG
documentMasterId    = signal<string | null>(null)      // guid_master (null if unsaved)
documentMetadata    = signal<DocumentMetadata>(...)    // title, author, created, etc.
editorState         = signal<EditorState | null>(null) // current text selection/format state
pageSettings        = signal<PageSettings>(...)        // margins, orientation
documentSignatures  = signal<DigitalSignatureInfo[]>([])
zoomLevel           = signal(100)                      // 50–200 %
```

Key methods:
- `newDocument()` — calls `DocumentService.newDocument()`, resets signals
- `openDocument()` — file picker → `DocumentService.openDocument(file)`
- `saveDocument()` — calls `DocumentService.downloadDocument(request, filename)`
- `loadFromStorage(masterId)` — `DocumentStorageService.getDocument(masterId)` + download binary + `DocumentService.openDocument`
- `uploadDocument()` — sends current HTML as DOCX bytes to `DocumentStorageService.uploadDocument`

---

## Services — method signatures

### DocumentService

```typescript
// Stateless — no DB interaction, pure file conversion
openDocument(file: File): Observable<DocumentContent>          // POST /api/document/open
saveDocument(req: SaveDocumentRequest): Observable<Blob>       // POST /api/document/save
newDocument(): Observable<DocumentContent>                     // GET  /api/document/new
getTemplates(): Observable<DocumentTemplate[]>                 // GET  /api/document/templates
getTemplate(id: string): Observable<DocumentContent>           // GET  /api/document/templates/{id}
uploadImage(file: File): Observable<ImageUploadResponse>       // POST /api/document/upload-image
downloadDocument(req: SaveDocumentRequest, filename: string)   // helper: save + FileSaver
signDocument(req: SignDocumentRequest): Observable<Blob>       // POST /api/document/sign
downloadSignedDocument(req, filename)                          // helper: sign + FileSaver
verifySignatures(file: File): Observable<DigitalSignatureInfo[]> // POST /api/document/verify-signatures
```

### DocumentStorageService

```typescript
// Persistent — uses DocumentStorage controller (CRUD with DB + GCS)
uploadDocument(name, mimeType, content: Uint8Array, createdBy): Observable<UploadDocumentResult>
saveDocumentVersion(masterId, content: Uint8Array, createdBy): Observable<SaveDocumentVersionResult>
getDocument(masterId: string): Observable<DocumentDto>
getDocuments(skip?, take?): Observable<DocumentListItemDto[]>
getDocumentVersions(masterId): Observable<DocumentVersionDto[]>
downloadBaseDocument(masterId): Observable<Blob>
downloadDocumentVersion(masterId, versionId): Observable<Blob>
restoreDocumentVersion(masterId, versionId): Observable<void>
```

---

## TypeScript models — `models/document.model.ts`

```typescript
interface DocumentContent {
  html: string
  metadata: DocumentMetadata
  styles: DocumentStyle[]
  images: DocumentImage[]
  header?: HeaderFooterContent
  footer?: HeaderFooterContent
  margins?: PageMargins
}

interface DocumentMetadata { title, author, created, modified, subject, keywords, description, lastModifiedBy }
interface DigitalSignatureInfo { signerName, signerTitle, signerEmail, reason, certificateSubject, certificateIssuer, certificateSerialNumber, signedAt, certificateValidFrom, certificateValidTo, isValid, validationMessage }
interface SignDocumentRequest { html, originalFileName, metadata, header, footer, certificateBase64, certificatePassword, signerName, signerTitle, signerEmail, signatureReason, margins }
interface SaveDocumentRequest  { html, originalFileName, metadata, header, footer, margins }
interface PageMargins          { top, bottom, left, right }  // cm
interface PageSettings         { margins: PageMargins, orientation: 'portrait'|'landscape' }
interface DocumentTemplate     { id, name, description }
interface ImageUploadResponse  { base64, mimeType, fileName }
type EditorCommand = 'bold'|'italic'|'underline'|'strikethrough'|...
```

---

## Environment

```typescript
// environment.development.ts
{ production: false, apiUrl: 'http://localhost:5190/api', environmentName: 'DEV', buildVersion: '1.0.0-dev' }

// environment.ts (production)
{ production: true,  apiUrl: 'http://localhost:5190/api', environmentName: 'PRD', buildVersion: '1.0.0' }
```

Note: prod `apiUrl` still points to localhost — likely replaced at deploy time (nginx env-var injection or CI/CD).

---

## Build

```bash
npm start          # ng serve (dev, port 4200)
npm run build      # production build → dist/d2-gui-viewereditor/browser/
npm test           # vitest
```

Served in prod via nginx (`nginx.conf` in repo root for this project). SPA routing — all paths rewrite to `index.html`.
