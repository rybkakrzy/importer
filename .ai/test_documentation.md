# Test Documentation - D2ViewerEditor

## 1. Zakres testów

Projekt wykorzystuje trzy główne warstwy testów:

1. Testy jednostkowe dla warstw backendu .NET
2. Testy behawioralne E2E/API w Pythonie (pytest + Playwright + BDD)
3. Testy performance/load z Locust

## 2. Testy jednostkowe backendu (.NET)

### Test Projects
- `D2ViewerEditor.Domain.UnitTests`
- `D2ViewerEditor.Application.UnitTests`
- `D2ViewerEditor.Infrastructure.UnitTests`
- `D2ViewerEditor.Api.UnitTests`

### Typowe polecenie
```powershell
cd D2ApiViewerEditor
dotnet test
```

### Obszary fokusowe
- Result pattern i niezmienniki domenowe
- Handlery CQRS command/query
- Zachowanie validatorów
- Logika serwisów (barcode, signatures, operacje na dokumentach)
- Zachowanie na poziomie kontrolerów i mapowanie odpowiedzi

## 3. Testy E2E i API (Python)

### Stack
- `pytest`
- `pytest-bdd`
- `playwright`

### Konfiguracja
- `D2TestViewerEditor/pytest.ini`
- `D2TestViewerEditor/conftest.py`
- Feature files in `D2TestViewerEditor/features`

### Markery
- `ui`
- `api`
- `smoke`
- `regression`
- `slow`
- `performance`

### Polecenia
```powershell
cd D2TestViewerEditor
pip install -r requirements.txt
playwright install chromium

# smoke UI
pytest -m "ui and smoke"

# testy API
pytest -m api

# pełna regresja
pytest -m regression
```

## 4. Testy performance (Locust)

### Lokalizacja
- `D2TestViewerEditor/performance`

### Scenariusze
- ruch na endpoint health
- endpointy barcode
- endpointy dokumentowe
- mieszany profil all-endpoints

### Polecenia
```powershell
cd D2TestViewerEditor\performance
pip install -r requirements.txt

# z UI
locust -f locustfiles/all_endpoints_locust.py

# przykładowy headless
locust -f locustfiles/all_endpoints_locust.py --headless -u 20 -r 5 -t 120s
```

## 5. Dane testowe i środowisko

Wymagane usługi runtime dla integracji/E2E/perf:

- działające API (`http://localhost:5190` lub URL kontenera)
- działające GUI (`http://localhost:4200`) dla przepływów UI
- dostępny Postgres
- osiągalny endpoint fake-gcs lub GCS

Wartości environment są podawane w plikach `.env` w katalogach testowych.

## 6. Quality gates (zalecane)

Minimalne checki CI przed merge:

1. `dotnet test` for all backend test projects
2. `npm test` in GUI project
3. `pytest -m "ui and smoke"` for critical path
4. Locust smoke run against `/api/health` + one document scenario

## 7. Checklista regresji

- DOCX open dla poprawnego i niepoprawnego pliku
- DOCX save zwraca plik do pobrania
- Version save zwiększa numer i zmienia aktywną wersję
- Restore version ustawia poprawny aktywny rekord
- Generowanie barcode wspiera wybrane formaty
- Signature signing i verification w scenariuszu happy path
- Endpoint health zwraca status/build metadata

