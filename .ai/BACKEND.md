# Backend — D2ApiViewerEditor

## Architecture

```
Api → Application → Domain ← Infrastructure
```

- **Domain** has zero external dependencies (only BCL).
- **Application** depends on Domain abstractions only.
- **Infrastructure** implements Domain interfaces.
- **Api** wires everything via DI (Program.cs + DependencyInjection.cs in each layer).

---

## Domain

### Entities

**`Document`** — `Domain/Entities/Document.cs`
- Aggregate root. Constructor validates name/mimeType (throws ArgumentException on empty).
- `AddVersion(storagePath, sizeInBytes, createdBy)` → creates new DocumentVersion, deactivates all existing versions, returns new version.
- `RestoreVersion(versionId)` → deactivates all, activates target (throws if versionId unknown).
- `GetActiveVersion()` → `_versions.FirstOrDefault(v => v.IsActive)` — can return null.
- `Delete()` → sets `IsDeleted = true` (soft delete).
- Private parameterless ctor for EF Core.

**`DocumentVersion`** — `Domain/Entities/DocumentVersion.cs`
- Not an aggregate root. Has `Activate()` / `Deactivate()` (likely internal/package-private).

### Interfaces — `Domain/Interfaces/`

```csharp
IDocumentRepository
  Task<Document?> GetByIdAsync(Guid id, CancellationToken)
  Task<Document?> GetByIdWithVersionsAsync(Guid id, CancellationToken)
  Task<IReadOnlyList<Document>> GetAllAsync(int skip, int take, CancellationToken)
  Task AddAsync(Document document, CancellationToken)
  Task UpdateAsync(Document document, CancellationToken)
  Task<int> SaveChangesAsync(CancellationToken)
  Task<bool> ExistsAsync(Guid id, CancellationToken)

IDocumentStorageService   (GCS abstraction)
  Task<string> UploadAsync(Guid versionId, byte[] content, string mimeType, CancellationToken)
  Task<byte[]> DownloadAsync(string storagePath, CancellationToken)
  Task DeleteAsync(string storagePath, CancellationToken)
  Task<bool> ExistsAsync(string storagePath, CancellationToken)

IDocxToHtmlConverter
  DocumentContent Convert(Stream docxStream)

IHtmlToDocxConverter
  byte[] Convert(string html, DocumentMetadata? metadata,
                 HeaderFooterContent? header, HeaderFooterContent? footer,
                 PageMargins? margins)

IDigitalSignatureService
  byte[] SignDocument(byte[] docxBytes, byte[] certificateBytes, string password,
                      string signerName, string? signerTitle, string? signerEmail, string? reason)
  List<DigitalSignatureInfo> VerifySignatures(Stream docxStream)

IBarcodeGenerator
  byte[] Generate(string content, string barcodeType, int width, int height, bool showText)
  IReadOnlyList<string> GetSupportedTypes()
```

### Domain Models — `Domain/Models/DocumentModels.cs`

`DocumentContent`, `DocumentMetadata`, `DigitalSignatureInfo`, `SignDocumentRequest`, `SaveDocumentRequest`, `DocumentImage`, `DocumentStyle`, `ParagraphStyle`, `PageMargins`, `HeaderFooterContent`

---

## Application Layer

### Pipeline Behaviours

1. `LoggingBehaviour<TRequest, TResponse>` — logs all MediatR requests
2. `ValidationBehaviour<TRequest, TResponse>` — runs all FluentValidation validators before handler

### Commands (state-changing)

| Command | Handler | What it does |
|---|---|---|
| `UploadDocumentCommand` | `UploadDocumentCommandHandler` | Creates Document aggregate, uploads bytes to GCS, saves to DB |
| `SaveDocumentVersionCommand(masterId, content, createdBy)` | `SaveDocumentVersionCommandHandler` | Loads doc, calls AddVersion, uploads to GCS, saves DB |
| `RestoreDocumentVersionCommand(masterId, versionId)` | `RestoreDocumentVersionCommandHandler` | Loads doc with versions, calls RestoreVersion, saves DB |
| `SaveDocumentCommand(html, filename, metadata, header, footer, margins)` | `SaveDocumentCommandHandler` | HTML→DOCX conversion (stateless, no DB) |
| `SignDocumentCommand` | `SignDocumentCommandHandler` | HTML→DOCX, then sign via IDigitalSignatureService (stateless) |
| `UploadImageCommand(stream, fileName, contentType, length)` | `UploadImageCommandHandler` | Reads image bytes, returns Base64 |
| `GenerateBarcodeCommand` | `GenerateBarcodeCommandHandler` | IBarcodeGenerator.Generate, returns Base64 PNG |

### Queries (read-only)

| Query | What it returns |
|---|---|
| `GetDocumentsQuery(skip, take)` | Paginated list of `DocumentListItemDto` |
| `GetDocumentQuery(masterId)` | `DocumentDto` with active version |
| `GetDocumentVersionsQuery(masterId)` | `List<DocumentVersionDto>` |
| `GetDocumentVersionContentQuery(masterId, versionId)` | File bytes + mime + name |
| `GetDocumentBaseContentQuery(masterId)` | First version file bytes |
| `OpenDocumentQuery(stream, fileName)` | `DocumentContent` (DOCX → HTML) |
| `GetNewDocumentQuery()` | Blank `DocumentContent` |
| `GetTemplatesQuery()` | `List<DocumentTemplate>` |
| `GetTemplateQuery(templateId)` | `DocumentContent` |
| `GetSupportedTypesQuery()` | `IReadOnlyList<string>` barcode types |

---

## Infrastructure

### Database — `Infrastructure/Persistence/`

- `DocumentDbContext` : DbContext
  - `DbSet<Document> Documents`
  - `DbSet<DocumentVersion> DocumentVersions`
  - Config: `DocumentConfiguration`, `DocumentVersionConfiguration`
- `DocumentRepository` : IDocumentRepository
  - Read queries use `AsNoTracking()`
  - Soft-delete filter applied on queries (IsDeleted == false)
- Provider: Npgsql, PostgreSQL

### GCS Storage — `Infrastructure/Services/GcsDocumentStorageService.cs`

- Config: `GcsStorageOptions` (bucket name, optional custom endpoint for DEV)
- DEV: points to `fake-gcs-server` (runs locally), unauthenticated
- PROD: service account file OR Workload Identity
- Object naming: `{versionId}` (Guid as storage key)

### DOCX Processing

- `DocxToHtmlConverter` : IDocxToHtmlConverter — uses `DocumentFormat.OpenXml` + `HtmlAgilityPack`
- `HtmlToDocxConverter` : IHtmlToDocxConverter — builds DOCX from HTML, preserves metadata/margins/header/footer

### Digital Signature — `Infrastructure/Services/DigitalSignatureService.cs`

- **NOT** standard OOXML/XAdES signatures.
- Custom XML Part in `MainDocumentPart.CustomXmlParts`.
- Namespace: `http://schemas.D2ViewerEditor.app/digitalsignatures`
- Root element: `<DigitalSignatures>`, children: `<Signature>`
- Each `<Signature>` stores: SignerName, SignerTitle, SignerEmail, Reason, SignedAt, CertificateSubject/Issuer/Serial, CertificateValidFrom/To, DocumentHash (Base64 SHA-256), SignatureValue (Base64 RSA-SHA256), CertificateData (Base64 raw cert bytes)
- Hash is computed from `MainDocumentPart.GetStream()` bytes only.
- `VerifySignatures` checks: (a) RSA signature valid, (b) stored hash == current hash. Reports each independently.

### Barcode — `Infrastructure/Services/BarcodeGeneratorService.cs`

- ZXing.Net encodes, SkiaSharp renders to PNG
- Returns `byte[]` PNG

---

## Configuration

| Env | appsettings file | Swagger | CORS allowed origin |
|---|---|---|---|
| DEV | `appsettings.DEV.json` | enabled | `http://localhost:4200` |
| UAT | `appsettings.UAT.json` | (check file) | (check file) |
| PRD | `appsettings.PRD.json` | disabled | (check file) |

Secrets in `appsettings.{env}.secrets.json` — **not committed to repo**.

---

## Result pattern

`Result<T>` in `Domain/Common/Result.cs`. Methods: `.IsSuccess`, `.IsNotFound`, `.Error`, `.Value`, `.Match<T>(onSuccess, onFailure)`.

Controllers use `HandleResult(result)` (base class helper) or explicit `result.Match(...)`.
