# D2PerformenceTools — Testy wydajnościowe (Locust)

Testy wydajnościowe dla **D2Tools API** oparte o framework [Locust](https://locust.io/).

## Struktura projektu

```
D2PerformenceTools/
├── .env                          # Zmienne środowiskowe (URL, parametry)
├── locust.conf                   # Domyślna konfiguracja Locusta
├── requirements.txt              # Zależności Python
├── README.md
├── config/
│   └── settings.py               # Centralna konfiguracja
├── helpers/
│   └── payloads.py               # Generatory danych testowych
└── locustfiles/
    ├── health_locust.py          # Smoke test — Health Check
    ├── barcode_locust.py         # Barcode API (types, generate, generate-image)
    ├── document_locust.py        # Document API (new, save, open, templates, upload-image)
    └── all_endpoints_locust.py   # Mix wszystkich endpointów (realistyczny ruch)
```

## Wymagania

- Python 3.10+
- Działające API D2Tools (domyślnie `http://localhost:5190`)

## Instalacja

```bash
cd D2PerformenceTools
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Linux / macOS
pip install -r requirements.txt
```

## Uruchamianie

### Web UI (domyślnie)

```bash
# Wszystkie endpointy (mix ruchu) — otwiera UI na http://localhost:8089
locust

# Lub konkretny plik:
locust -f locustfiles/barcode_locust.py
locust -f locustfiles/document_locust.py
locust -f locustfiles/health_locust.py
```

### Tryb headless (CLI)

```bash
locust --headless -u 20 -r 5 -t 120s
```

| Parametr | Opis                               |
|----------|-------------------------------------|
| `-u`     | Liczba wirtualnych użytkowników     |
| `-r`     | Spawn rate (użytk./s)               |
| `-t`     | Czas trwania testu                  |
| `--host` | Nadpisanie URL hosta API            |
| `--tags` | Filtrowanie tasków po tagach        |

### Filtrowanie po tagach

```bash
# Tylko operacje odczytu
locust --tags read

# Tylko Barcode
locust --tags barcode

# Barcode + Document (write)
locust --tags barcode document --exclude-tags read
```

Dostępne tagi: `health`, `barcode`, `document`, `read`, `write`, `large`.

## Konfiguracja

Edytuj plik `.env` lub nadpisz zmiennymi środowiskowymi:

| Zmienna            | Domyślna                     | Opis                     |
|--------------------|------------------------------|--------------------------|
| `API_BASE_URL`     | `http://localhost:5190/api`  | Bazowy URL API           |
| `LOCUST_USERS`     | `10`                         | Liczba użytkowników      |
| `LOCUST_SPAWN_RATE`| `2`                          | Spawn rate               |
| `LOCUST_RUN_TIME`  | `60s`                        | Czas trwania testu       |

## Pokryte endpointy

| Endpoint                                     | Metoda | Locustfile               |
|----------------------------------------------|--------|--------------------------|
| `/api/Health`                                | GET    | health_locust.py         |
| `/api/Barcode/types`                         | GET    | barcode_locust.py        |
| `/api/Barcode/generate`                      | POST   | barcode_locust.py        |
| `/api/Barcode/generate-image`                | POST   | barcode_locust.py        |
| `/api/Document/new`                          | GET    | document_locust.py       |
| `/api/Document/save`                         | POST   | document_locust.py       |
| `/api/Document/open`                         | POST   | document_locust.py       |
| `/api/Document/templates`                    | GET    | document_locust.py       |
| `/api/Document/upload-image`                 | POST   | document_locust.py       |

Plik `all_endpoints_locust.py` łączy wszystkie powyższe w jeden realistyczny scenariusz.

## Raporty

Locust generuje raporty w Web UI. Aby wyeksportować do pliku:

```bash
locust --headless -u 20 -r 5 -t 120s --csv=reports/perf --html=reports/report.html
```

Pliki wyjściowe:
- `reports/perf_stats.csv` — statystyki per request
- `reports/perf_failures.csv` — błędne requesty
- `reports/perf_stats_history.csv` — historia w czasie
- `reports/report.html` — raport HTML z wykresami
