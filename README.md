# Doc2 — Edytor dokumentów Word online

Autorski edytor DOCX w przeglądarce oparty na **Angular 20** i **ASP.NET Core 8**.  
Umożliwia otwieranie, edycję i zapisywanie plików DOCX bez instalacji Microsoft Office — w pełni w przeglądarce.

## Struktura projektu

```
importer/
├── D2ApiTools/        # Backend — ASP.NET Core 8 Web API (Clean Architecture)
├── D2GuiTools/        # Frontend — Angular 20 (standalone components)
└── D2E2ETools/        # Testy automatyczne E2E — Python + Playwright + pytest-bdd
```

## Funkcjonalność

### Backend — D2ApiTools (ASP.NET Core 8 Web API)

Architektura czysta (Clean Architecture) z warstwami: Api, Application, Domain, Infrastructure.

#### Endpointy API

**Dokumenty** (`/api/Document`)

| Metoda | Ścieżka | Opis |
|--------|---------|------|
| `POST` | `/api/Document/open` | Wczytuje plik DOCX i zwraca HTML |
| `POST` | `/api/Document/save` | Zapisuje HTML jako plik DOCX |
| `GET`  | `/api/Document/new` | Tworzy nowy pusty dokument |
| `POST` | `/api/Document/upload-image` | Wgrywa obraz i zwraca Base64 |
| `GET`  | `/api/Document/templates` | Pobiera listę szablonów dokumentów |
| `GET`  | `/api/Document/templates/{id}` | Pobiera wybrany szablon |
| `POST` | `/api/Document/sign` | Podpisuje dokument certyfikatem X.509 |
| `POST` | `/api/Document/verify-signatures` | Weryfikuje podpisy cyfrowe w dokumencie |
| `POST` | `/api/Document/export-pdf` | Eksport do PDF *(w przygotowaniu)* |

**Kody kreskowe / QR** (`/api/Barcode`)

| Metoda | Ścieżka | Opis |
|--------|---------|------|
| `GET`  | `/api/Barcode/types` | Pobiera listę obsługiwanych typów kodów |
| `POST` | `/api/Barcode/generate` | Generuje kod kreskowy / QR — zwraca Base64 |
| `POST` | `/api/Barcode/generate-image` | Generuje kod kreskowy / QR — zwraca plik PNG |

**Import plików ZIP** (`/api/FileUpload`)

| Metoda | Ścieżka | Opis |
|--------|---------|------|
| `POST` | `/api/FileUpload/upload` | Przyjmuje plik ZIP i analizuje jego zawartość |

### Frontend — D2GuiTools (Angular 20)

- Edytor WYSIWYG dokumentów DOCX w przeglądarce
- Pasek narzędzi (formatowanie, style, wstawianie obrazów i kodów QR)
- Dialog generowania kodów kreskowych / QR (13 formatów)
- Wyszukiwanie i zamiana tekstu z licznikiem wyników
- Zarządzanie metadanymi dokumentu (właściwości OOXML)
- Podpisywanie dokumentów certyfikatem cyfrowym
- Wielostronicowy podgląd A4 z zoomem (50%–200%)
- Import pliku ZIP z parametryzacją

### Testy E2E — D2E2ETools (Python + Playwright)

- Testy BDD w języku polskim (Gherkin / pytest-bdd)
- Scenariusze UI: edytor dokumentów, pasek narzędzi, kody kreskowe, znajdź i zamień
- Scenariusze API: endpointy dokumentów
- Raporty HTML i Allure

## Uruchomienie

### Backend

```powershell
cd D2ApiTools
dotnet run --project D2Tools.Api
```

API będzie dostępne pod adresem: `http://localhost:5190`  
Swagger UI (środowisko deweloperskie): `http://localhost:5190/swagger`

### Frontend

```powershell
cd D2GuiTools
npm install
npm start
```

Aplikacja będzie dostępna pod adresem: `http://localhost:4200`

### Testy E2E

```bash
cd D2E2ETools

# Utwórz wirtualne środowisko
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Linux/macOS

# Zainstaluj zależności
pip install -r requirements.txt
playwright install --with-deps chromium

# Uruchom testy
pytest                          # Wszystkie testy
pytest -m "ui and smoke"        # Testy UI (smoke)
pytest -m api                   # Testy API
pytest -m regression            # Testy regresji
pytest --html=report.html       # Z raportem HTML
```

> Przed uruchomieniem testów E2E upewnij się, że backend i frontend są uruchomione.

## Technologie

### Backend
- ASP.NET Core 8 Web API
- MediatR (CQRS / Clean Architecture)
- System.IO.Compression (obsługa ZIP)
- Konwersja DOCX ↔ HTML (dwukierunkowa)
- Podpisy cyfrowe X.509 (RSA-SHA256)
- Generowanie kodów kreskowych i QR (13 formatów)
- Swagger / OpenAPI

### Frontend
- Angular 20 (standalone components)
- HttpClient — komunikacja z API
- WYSIWYG edytor dokumentów DOCX
- TypeScript 5.8

### Testy
- Python 3.11+
- Playwright (automatyzacja przeglądarki)
- pytest-bdd (BDD / Gherkin, język polski)
- Allure / pytest-html (raporty)

## Roadmap

- [ ] Eksport do PDF
- [ ] Kolaboracja w czasie rzeczywistym (WebSocket)
- [ ] Śledzenie zmian (Track Changes)
- [ ] Komentarze w dokumentach
- [ ] Tryb offline (PWA)


