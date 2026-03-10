# D2TestViewerEditor — testy automatyczne

## Opis

Projekt testów automatycznych dla aplikacji **D2 ViewerEditor**.  
Zawiera:
- **Testy E2E** — Python + Playwright + pytest-bdd (podejście BDD)
- **Testy wydajnościowe** — Locust (load testing)

## Struktura projektu

```
D2TestViewerEditor/
├── .env                       # Zmienne środowiskowe
├── requirements.txt           # Zależności Python (E2E)
├── pytest.ini                 # Konfiguracja pytest
├── conftest.py                # Globalne fixtures (Playwright browser/page)
│
├── config/                    # Konfiguracja środowisk (E2E)
│   ├── environments.py
│   └── settings.py
│
├── pages/                     # Page Object Model
│   ├── base_page.py
│   ├── document_editor_page.py
│   ├── toolbar_page.py
│   ├── barcode_dialog_page.py
│   └── file_upload_page.py
│
├── features/                  # Scenariusze BDD (Gherkin, język: pl)
│   ├── ui/
│   │   ├── document_editor.feature
│   │   ├── toolbar.feature
│   │   ├── barcode.feature
│   │   └── find_replace.feature
│   └── api/
│       └── documents_api.feature
│
├── tests/                     # Implementacje kroków (step definitions)
│   ├── ui/
│   │   ├── conftest.py
│   │   ├── test_document_editor.py
│   │   ├── test_toolbar.py
│   │   ├── test_barcode.py
│   │   └── test_find_replace.py
│   └── api/
│       ├── conftest.py
│       └── test_documents_api.py
│
├── utils/                     # Narzędzia pomocnicze
│   ├── waits.py
│   ├── artifacts.py
│   └── test_data.py
│
├── test_data/                 # Pliki testowe (DOCX, ZIP, …)
│
└── performance/               # Testy wydajnościowe (Locust)
    ├── .env                   # Zmienne środowiskowe
    ├── locust.conf            # Konfiguracja Locust
    ├── requirements.txt       # Zależności Python (Locust)
    ├── README.md              # Dokumentacja testów wydajnościowych
    ├── config/
    │   └── settings.py
    ├── helpers/
    │   └── payloads.py
    └── locustfiles/
        ├── barcode_locust.py
        ├── document_locust.py
        ├── fileupload_locust.py
        ├── health_locust.py
        └── all_endpoints_locust.py
```

## Wymagania

- Python ≥ 3.11
- Działający backend (D2ApiViewerEditor) na `http://localhost:5190`
- Działający frontend (D2GuiViewerEditor) na `http://localhost:4200`

## Instalacja

### Testy E2E

```bash
cd D2TestViewerEditor

# Utwórz wirtualne środowisko
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Linux/macOS

# Zainstaluj zależności
pip install -r requirements.txt

# Zainstaluj przeglądarki Playwright
playwright install --with-deps chromium
```

### Testy wydajnościowe (Locust)

```bash
cd D2TestViewerEditor/performance

# Użyj tego samego .venv lub stwórz nowe
# Zainstaluj zależności Locust
pip install -r requirements.txt
```

## Uruchamianie testów

### Testy E2E

```bash
# Wszystkie testy
pytest

# Tylko testy UI (smoke)
pytest -m "ui and smoke"

# Tylko testy API
pytest -m api

# Tylko testy regresji
pytest -m regression

# Równoległe (4 procesy)
pytest -n 4

# Z raportem HTML
pytest --html=test-results/report.html

# Z raportem Allure
pytest --alluredir=test-results/allure
allure serve test-results/allure

# Tryb headed (widoczna przeglądarka)
D2_HEADLESS=false pytest -m "ui and smoke"
```

### Testy wydajnościowe

```bash
cd performance

# Web UI (http://localhost:8089)
locust -f locustfiles/all_endpoints_locust.py

# Headless mode
locust -f locustfiles/all_endpoints_locust.py --headless -u 100 -r 10 -t 60s

# Konkretny endpoint
locust -f locustfiles/barcode_locust.py
locust -f locustfiles/document_locust.py
```

## Markery

| Marker       | Opis                              |
|--------------|-----------------------------------|
| `@ui`        | Testy interfejsu użytkownika      |
| `@api`       | Testy REST API                    |
| `@smoke`     | Szybkie testy krytycznej ścieżki |
| `@regression`| Pełne testy regresyjne            |
| `@slow`      | Testy wymagające więcej czasu     |

## Konfiguracja

Zmienne środowiskowe (`.env`):

| Zmienna            | Domyślna                    | Opis                       |
|--------------------|-----------------------------|-----------------------------|
| `D2_BASE_URL`      | `http://localhost:4200`     | URL frontendu               |
| `D2_API_BASE_URL`  | `http://localhost:5190/api` | URL backendu API            |
| `D2_HEADLESS`      | `true`                      | Tryb headless przeglądarki  |
| `D2_SLOW_MO`       | `0`                         | Opóźnienie (ms) w Playwright|
| `D2_TIMEOUT`       | `30000`                     | Timeout nawigacji (ms)      |
| `D2_BROWSER`       | `chromium`                  | Przeglądarka testowa        |
