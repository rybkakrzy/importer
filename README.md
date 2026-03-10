# D2 ViewerEditor

System do przeglądania, edycji i zarządzania dokumentami DOCX z obsługą kodów kreskowych, podpisów cyfrowych i konwersji HTML↔DOCX.

## Struktura projektu

```
importer/
├── D2ApiViewerEditor/          # Backend - ASP.NET 8 Web API
│   ├── D2ViewerEditor.Api/
│   ├── D2ViewerEditor.Application/
│   ├── D2ViewerEditor.Domain/
│   └── D2ViewerEditor.Infrastructure/
│
├── D2GuiViewerEditor/          # Frontend - Angular 20
│   └── src/
│
└── D2TestViewerEditor/         # Testy automatyczne
    ├── config/                 # Testy E2E (Playwright + pytest-bdd)
    ├── features/
    ├── pages/
    ├── tests/
    └── performance/            # Testy wydajnościowe (Locust)
```

## Funkcjonalność

### Backend (ASP.NET 8 Web API)
- **Dokumenty**: Tworzenie, edycja, zapisywanie dokumentów DOCX
- **Konwersja**: DOCX ↔ HTML (z zachowaniem formatowania)
- **Kody kreskowe**: Generowanie QR, Code128, EAN-13, etc. (ZXing)
- **Podpisy cyfrowe**: Podpisywanie dokumentów certyfikatem X.509
- **Upload plików**: Obsługa ZIP z wieloma dokumentami
- **Architektura**: Clean Architecture + CQRS (MediatR)

### Frontend (Angular 20)
- Edytor WYSIWYG dokumentów z toolbar
- Dialog generowania kodów kreskowych
- Upload plików ZIP
- Podgląd i zapis dokumentów
- Standalone components, Signals

### Testy
- **E2E**: Playwright + pytest-bdd (scenariusze BDD po polsku)
- **Wydajnościowe**: Locust (load testing wszystkich endpointów)

## Uruchomienie

### Backend

```powershell
cd D2ApiViewerEditor\D2ViewerEditor.Api
dotnet run --environment DEV
```

API: `http://localhost:5190`  
Swagger: `http://localhost:5190/swagger`

### Frontend

```powershell
cd D2GuiViewerEditor
npm start
```

Aplikacja: `http://localhost:4200`

### Testy E2E

```powershell
cd D2TestViewerEditor
pip install -r requirements.txt
playwright install chromium
pytest -m "ui and smoke"
```

### Testy wydajnościowe

```powershell
cd D2TestViewerEditor\performance
pip install -r requirements.txt
locust -f locustfiles/all_endpoints_locust.py
```

Locust UI: `http://localhost:8089`

## Technologie

### Backend
- ASP.NET 8 Web API
- MediatR (CQRS), FluentValidation
- DocumentFormat.OpenXml (DOCX)
- HtmlAgilityPack (HTML parsing)
- ZXing.Net (kody kreskowe)
- SkiaSharp (rendering grafiki)

### Frontend
- Angular 20 (standalone, signals)
- TypeScript
- SCSS

### Testy
- Python 3.11+
- Playwright (E2E)
- pytest-bdd (BDD)
- Locust (load testing)


