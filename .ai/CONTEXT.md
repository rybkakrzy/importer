# D2 ViewerEditor — Master Context

> **Start here.** Load this file first in every new conversation.

## What this system is

A DOCX web editor with versioned storage. User uploads a .docx file, edits it in-browser (WYSIWYG), saves back to server. Supports digital signatures (custom X.509 via Custom XML Part), barcode/QR generation, PDF viewing, template library, and multi-version history.

## Projects

| Project | Type | Port (dev) | Root path |
|---|---|---|---|
| `D2ApiViewerEditor` | ASP.NET Core 8 Web API | 5190 | `D2ApiViewerEditor/` |
| `D2GuiViewerEditor` | Angular 20 SPA | 4200 | `D2GuiViewerEditor/` |

## Tech stack — key facts

**Backend**
- .NET 8, Clean Architecture (Domain → Application → Infrastructure → Api)
- CQRS via MediatR 12.4.1 (commands change state, queries read)
- FluentValidation 11.11 — all commands have validators
- PostgreSQL via EF Core 8 + Npgsql 8.0.11
- Google Cloud Storage for binary files (fake-gcs-server in DEV)
- DocumentFormat.OpenXml 3.0.2 for DOCX ↔ HTML conversion
- ZXing.Net 0.16.9 + SkiaSharp 3.116.1 for barcodes
- Digital signatures: RSA-SHA256, stored as Custom XML Part in DOCX, **NOT** standard OOXML signatures

**Frontend**
- Angular 20, standalone components, Signals (no NgRx)
- RxJS 7.8 (HTTP only, not used for state)
- Vitest 3.1.1 for tests
- pdfjs-dist 5.5.207 for PDF viewer page

## Solution layout (backend)

```
D2ViewerEditor.sln
├── D2ViewerEditor.Domain          — entities, interfaces, models
├── D2ViewerEditor.Application     — use cases, CQRS handlers, validators, behaviours
├── D2ViewerEditor.Infrastructure  — EF Core, GCS, converters, signature service
├── D2ViewerEditor.Api             — controllers, middleware, Program.cs
└── *.UnitTests / Benchmarks
```

## Key data model

```
Document (aggregate root)
  Id (Guid)          — guid_master (used in all API routes)
  Name, MimeType, CreatedAt, CreatedBy, IsDeleted (soft delete)
  Versions: List<DocumentVersion>
    Id (Guid)        — guid_wersji
    StoragePath      — GCS object path
    SizeInBytes, VersionNumber, CreatedAt, CreatedBy
    IsActive         — only ONE version active per document at a time
```

## Critical behaviour to know

1. `Document.AddVersion()` automatically deactivates all previous versions.
2. Digital signatures are stored as Custom XML Part (namespace `http://schemas.D2ViewerEditor.app/digitalsignatures`), not standard OOXML. Signing after saving regenerates DOCX and adds the XML part.
3. Document hash is SHA-256 of `MainDocumentPart` bytes only (not the whole package). Modifying the DOCX after signing invalidates the hash but the RSA signature itself stays valid — `VerifySignatures` reports both states separately.
4. PDF export → `501 Not Implemented` (intentional placeholder).
5. `DocumentController` handles stateless file ops (open/save/sign without persistence). `DocumentStorageController` handles CRUD with DB+GCS.

## Sub-documents

- [BACKEND.md](BACKEND.md) — full backend details (all interfaces, handlers, DB, infra services)
- [FRONTEND.md](FRONTEND.md) — Angular architecture, components, signals, services
- [API.md](API.md) — all REST endpoints with request/response shapes
