# REST API — D2ApiViewerEditor

Base: `http://localhost:5190/api` (dev)

---

## DocumentController — `/api/document`

Stateless file operations. No database, no GCS persistence.

### `POST /api/document/open`

Upload DOCX file, receive HTML.

- Content-Type: `multipart/form-data`
- Field: `file` (.docx, max 50 MB)
- Response 200: `DocumentContent` JSON
- Response 400: `{ error: string }`

### `POST /api/document/save`

Convert HTML back to DOCX, receive file download.

- Body JSON: `SaveDocumentRequest`
  ```json
  { "html": "...", "originalFileName": "doc.docx",
    "metadata": { "title":"", "author":"", ... },
    "header": { "html":"", ... }, "footer": { "html":"", ... },
    "margins": { "top":2.5, "bottom":2.5, "left":3.0, "right":2.0 } }
  ```
- Response 200: `application/vnd.openxmlformats-officedocument.wordprocessingml.document` (file download)
- Response 400: `{ error: string }`

### `GET /api/document/new`

Blank document template.

- Response 200: `DocumentContent` JSON

### `POST /api/document/upload-image`

Upload image, get Base64 back.

- Content-Type: `multipart/form-data`
- Field: `file` (max 10 MB)
- Response 200: `ImageUploadResponse` `{ base64, mimeType, fileName }`

### `GET /api/document/templates`

- Response 200: `DocumentTemplate[]` `{ id, name, description }[]`

### `GET /api/document/templates/{templateId}`

- Response 200: `DocumentContent`

### `POST /api/document/sign`

HTML → DOCX → digitally signed DOCX.

- Body JSON: `SignDocumentRequest`
  ```json
  { "html": "...", "originalFileName": "doc.docx",
    "metadata": {...}, "header": {...}, "footer": {...}, "margins": {...},
    "certificateBase64": "...", "certificatePassword": "...",
    "signerName": "Jan Kowalski", "signerTitle": "Dyrektor",
    "signerEmail": "jan@example.com", "signatureReason": "Zatwierdzam" }
  ```
- Response 200: signed `.docx` file download
- Response 400: `{ error: string }`

### `POST /api/document/verify-signatures`

Upload DOCX, receive list of signature verification results.

- Content-Type: `multipart/form-data`, field: `file` (max 50 MB)
- Response 200: `DigitalSignatureInfo[]`
  ```json
  [{ "signerName":"", "signerTitle":"", "signerEmail":"", "reason":"",
     "certificateSubject":"", "certificateIssuer":"", "certificateSerialNumber":"",
     "signedAt":"ISO8601", "certificateValidFrom":"ISO8601", "certificateValidTo":"ISO8601",
     "isValid": true|false, "validationMessage":"" }]
  ```

### `POST /api/document/export-pdf`

**501 Not Implemented** — intentional placeholder.

---

## DocumentStorageController — `/api/documentstorage`

Persistent operations — DB metadata + GCS binary storage.

### `GET /api/documentstorage?skip=0&take=200`

List all documents (admin use).

- Response 200: `DocumentListItemDto[]`

### `POST /api/documentstorage/upload`

Upload new document (first-time save).

- Body JSON:
  ```json
  { "name": "document.docx", "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "content": [/* byte array */], "createdBy": "user123" }
  ```
- Response 200: `UploadDocumentResult` `{ masterId: guid, versionId: guid }`
- Response 400: `{ error: string }`

### `POST /api/documentstorage/{masterId}/save`

Save new version of existing document.

- Body JSON:
  ```json
  { "content": [/* byte array */], "createdBy": "user123" }
  ```
- Response 200: `SaveDocumentVersionResult` `{ versionId: guid }`
- Response 404: `{ error: string }`

### `GET /api/documentstorage/{masterId}`

Get document metadata with active version.

- Response 200: `DocumentDto`
- Response 404

### `GET /api/documentstorage/{masterId}/versions`

Get version history (metadata only, no content).

- Response 200: `DocumentVersionDto[]`

### `GET /api/documentstorage/{masterId}/download`

Download first (base) version binary.

- Response 200: file with correct Content-Type

### `GET /api/documentstorage/{masterId}/versions/{versionId}/download`

Download specific version binary.

- Response 200: file

### `POST /api/documentstorage/{masterId}/restore/{versionId}`

Restore a previous version as active.

- Response 200: `{ message: "...", versionId: guid }`
- Response 404 / 400

---

## BarcodeController — `/api/barcode`

### `POST /api/barcode/generate`

- Body JSON:
  ```json
  { "content": "https://example.com", "barcodeType": "QR_CODE",
    "width": 200, "height": 200, "showText": false }
  ```
- Response 200: `{ imageBase64: "...", mimeType: "image/png" }`

### `GET /api/barcode/types`

- Response 200: `string[]` — e.g. `["QR_CODE", "CODE_128", "EAN_13", ...]`

---

## HealthController — `/api/health`

Standard health check endpoints.

---

## Error response shape

All error responses follow: `{ "error": "message string" }` or Problem Details (check BaseApiController).
