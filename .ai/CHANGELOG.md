# AI-Assisted Changelog

Istotne zmiany dla kontynuacji pracy (nie zastępuje changeloga produktu).

## Format

```md
## YYYY-MM-DD — <tytuł>
### Changed
### Verified
### Notes
```

## Entries

## 2026-05-23 — Dostosowanie `.ai/` do realnego projektu

### Changed
- Przyjęto strukturę `.ai/` wg `template/` (INDEX, PROJECT_CONTEXT, TECH_STACK, ARCHITECTURE, DOMAIN, FEATURES, DATABASE, API_CONTRACTS, SECURITY, TESTING_QUALITY, DEVOPS_DEPLOYMENT, DECISIONS, RISKS_ASSUMPTIONS, GLOSSARY, CURRENT_STATE, TASK_HANDOFF, CHANGELOG, PRODUCT_GOALS + pliki-wytyczne i `templates/`).
- Wypełniono faktami z repo; oznaczono niepewności (port External API, brak compose/CI).
- Dodano `CLAUDE.md` i `AGENTS.md` w root.
- Usunięto stare `.ai/{CONTEXT,API,BACKEND,FRONTEND}.md` (treść przeniesiona do nowej struktury).

### Verified
- Pliki zapisane; fakty zebrane z `*.csproj`, `package.json`, `docker/*`, `launchSettings`, `appsettings`, `infra/sql/`.

### Notes
- Pliki ogólne (BACKEND_DOTNET, FRONTEND_ANGULAR, CODING_STANDARDS, AI_AGENT_WORKFLOW, PROMPTS, README, templates) pozostawiono zgodne ze wzorcem.

## 2026-05-23 — Płaski zapis wersji + auto-save + ujednolicony „Zapisz"

### Changed
- Backend: `Document.UpdateVersion`, `DocumentVersion.UpdateContent`/`ModifiedAt`, `UpdateDocumentVersionCommand` + `PUT /api/documentstorage/{masterId}/versions/{versionId}`, `GetDocumentMetadataQuery` + `GET .../{masterId}/metadata`, skrypt `infra/sql/005_add_version_modified_at.sql` + mapowanie EF.
- Frontend: mechanizm auto-save (timer + `environment.autoSave`), switch „AutoSave", ujednolicony `saveDocument()` (API), `downloadDocument()` (pobranie lokalne), usunięto `saveDocumentAs()`.

### Verified
- `dotnet build D2ViewerEditor.sln` — OK (0 błędów).
- `npm run build` (GUI) — OK.

### Notes
- Tryb podglądu (Krok 2) wciąż ładuje aktywną wersję zamiast v1 — do dokończenia.
- `finishDocument()` (zwrot na returnUrl) — TODO.
