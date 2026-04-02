# Infrastructure Setup - D2 Viewer Editor

Konfiguracja infrastruktury dla aplikacji D2 Viewer Editor: **PostgreSQL**, **fake-gcs-server** (emulacja Google Cloud Storage) i kontenery aplikacji na **Podman**.

---

## 📦 Wymagania

| Narzędzie | Wersja | Instalacja (Windows) |
|-----------|--------|----------------------|
| **Podman** | 4.x+ | `winget install RedHat.Podman` |
| **podman-compose** | 1.x+ | `pip install podman-compose` |
| **Git** | dowolna | `winget install Git.Git` |
| **.NET SDK** | 8.0+ | `winget install Microsoft.DotNet.SDK.8` (tylko przy pracy z API lokalnie) |

> **Uwaga:** Na Windows Podman wymaga aktywnego WSL 2. Jeśli jeszcze go nie masz:
> ```powershell
> wsl --install
> podman machine init
> podman machine start
> ```

---

## 🚀 Szybki Start — wszystko jednym poleceniem

```bash
cd infra
podman-compose up -d
```

To polecenie uruchomi **5 kontenerów**:

| Kontener | Port | Opis |
|----------|------|------|
| `d2viewereditor_postgres` | **5432** | Baza PostgreSQL 16 |
| `d2viewereditor_pgadmin` | **5050** | pgAdmin (GUI do bazy) |
| `d2viewereditor_fake_gcs` | **4443** | Emulacja Google Cloud Storage |
| `d2viewereditor_api` | **8080** | Backend .NET 8 |
| `d2viewereditor_frontend` | **4200** | Frontend Angular (nginx) |

### Sprawdzenie statusu

```bash
podman-compose ps
```

Oczekiwany wynik — wszystkie kontenery w statusie `Up`:

```
d2viewereditor_postgres    ... Up (healthy)   0.0.0.0:5432->5432/tcp
d2viewereditor_pgadmin     ... Up             0.0.0.0:5050->80/tcp
d2viewereditor_fake_gcs    ... Up             0.0.0.0:4443->4443/tcp
d2viewereditor_api         ... Up             0.0.0.0:8080->8080/tcp
d2viewereditor_frontend    ... Up             0.0.0.0:4200->80/tcp
```

### Zatrzymanie

```bash
# Zatrzymaj kontenery (zachowaj dane)
podman-compose stop

# Zatrzymaj i usuń kontenery (dane w volumes zostają)
podman-compose down

# Zatrzymaj i usuń wszystko włącznie z danymi
podman-compose down -v
```

---

## 🗄️ Krok 1 — Baza danych PostgreSQL

### Co się dzieje automatycznie

Przy pierwszym `podman-compose up` kontener PostgreSQL:
1. Tworzy bazę `d2viewereditor` z użytkownikiem `postgres`/`postgres`
2. Automatycznie wykonuje skrypty SQL z `sql/` (zamontowane do `/docker-entrypoint-initdb.d`)
   - `001_init_schema.sql` — tabele `documents` i `document_versions`
   - `002_init_indexes.sql` — indeksy wydajnościowe
   - `003_migrate_storage_to_gcs.sql` — dodanie kolumn `storage_path` i `size_in_bytes`

### Dane dostępowe (DEV)

| Parametr | Wartość |
|----------|---------|
| Host | `localhost` |
| Port | `5432` |
| Baza | `d2viewereditor` |
| User | `postgres` |
| Password | `postgres` |

### Connection String (.NET)

```
Host=localhost;Port=5432;Database=d2viewereditor;Username=postgres;Password=postgres
```

### Weryfikacja — czy baza działa

```bash
# Healthcheck
podman exec -it d2viewereditor_postgres pg_isready -U postgres -d d2viewereditor

# Lista tabel
podman exec -it d2viewereditor_postgres psql -U postgres -d d2viewereditor -c "\dt"
```

Oczekiwany wynik `\dt`:

```
 Schema |       Name        | Type  |  Owner
--------+-------------------+-------+----------
 public | document_versions | table | postgres
 public | documents         | table | postgres
```

### Ręczne uruchomienie migracji (istniejąca baza)

Jeśli baza już istnieje (volume z danymi) i chcesz dodać nową migrację:

```bash
cd infra
podman exec -i d2viewereditor_postgres psql -U postgres -d d2viewereditor < sql/003_migrate_storage_to_gcs.sql
```

### Bezpośrednie połączenie (psql)

```bash
podman exec -it d2viewereditor_postgres psql -U postgres -d d2viewereditor
```

---

## ☁️ Krok 2 — Google Cloud Storage (fake-gcs-server)

W środowisku DEV zamiast prawdziwego GCS używamy **fake-gcs-server** — lekkiego emulatora HTTP-compatible z API Google Cloud Storage.

### Jak to działa

```
┌─────────────┐       HTTP        ┌──────────────────┐
│  .NET API    │ ──────────────── │  fake-gcs-server  │
│  (port 8080) │  PUT/GET objects │  (port 4443)      │
└─────────────┘                   └──────────────────┘
```

- API wysyła requesty do `http://fake-gcs:4443` (wewnątrz sieci Podman) lub `http://localhost:4443` (z hosta)
- fake-gcs przechowuje dane **w pamięci** (backend `memory`) — po restarcie kontenera pliki znikają
- Bucket `d2viewereditor-documents` jest tworzony automatycznie z pliku `fake-gcs-data/d2viewereditor-documents/.bucket_metadata.json`

### Konfiguracja w podman-compose.yml

```yaml
fake-gcs:
  image: docker.io/fsouza/fake-gcs-server:latest
  command: ["-scheme", "http", "-backend", "memory", "-data", "/data"]
  ports:
    - "4443:4443"
  volumes:
    - ./fake-gcs-data:/data:ro
```

### Konfiguracja w API (.NET)

W `appsettings.DEV.json`:
```json
{
  "GoogleCloudStorage": {
    "BucketName": "d2viewereditor-documents",
    "ApiEndpoint": "http://localhost:4443"
  }
}
```

W `podman-compose.yml` (wewnątrz sieci kontenerów):
```yaml
environment:
  GoogleCloudStorage__ApiEndpoint: http://fake-gcs:4443
  GoogleCloudStorage__BucketName: d2viewereditor-documents
```

### Weryfikacja — czy GCS działa

```bash
# Sprawdź listę bucketów
curl http://localhost:4443/storage/v1/b

# Oczekiwany wynik:
# {"kind":"storage#buckets","items":[{"kind":"storage#bucket","id":"d2viewereditor-documents","name":"d2viewereditor-documents"}]}
```

### Upload testowego pliku do GCS

```bash
# Upload
curl -X POST "http://localhost:4443/upload/storage/v1/b/d2viewereditor-documents/o?uploadType=media&name=test/hello.txt" ^
  -H "Content-Type: text/plain" ^
  -d "Hello from fake GCS!"

# Download
curl "http://localhost:4443/storage/v1/b/d2viewereditor-documents/o/test%%2Fhello.txt?alt=media"

# Lista obiektów
curl "http://localhost:4443/storage/v1/b/d2viewereditor-documents/o"
```

> **Windows cmd:** Użyj `^` do łamania linii. W PowerShell użyj `` ` ``.

### Struktura danych inicjalizacyjnych

```
infra/fake-gcs-data/
└── d2viewereditor-documents/
    └── .bucket_metadata.json    ← definiuje bucket
```

Aby dodać pliki startowe do bucketa, wystarczy umieścić je w katalogu `fake-gcs-data/d2viewereditor-documents/`:

```
infra/fake-gcs-data/
└── d2viewereditor-documents/
    ├── .bucket_metadata.json
    └── documents/
        └── sample.pdf          ← będzie dostępny jako "documents/sample.pdf"
```

---

## 🎨 Krok 3 — pgAdmin (opcjonalny GUI do bazy)

pgAdmin jest dostępny pod: **http://localhost:5050**

| Parametr | Wartość |
|----------|---------|
| Email | `admin@d2viewereditor.local` |
| Password | `admin` |

### Dodanie serwera

1. Otwórz http://localhost:5050
2. **Add New Server**
   - General → Name: `D2 ViewerEditor DEV`
   - Connection:
     - Host: `postgres` (nazwa serwisu w docker-compose, **nie** `localhost`)
     - Port: `5432`
     - Database: `d2viewereditor`
     - Username: `postgres`
     - Password: `postgres`

---

## 🧪 Krok 4 — Uruchomienie API (.NET)

### Opcja A: W kontenerze (przez podman-compose)

API startuje automatycznie razem z `podman-compose up -d`. Dostępne pod `http://localhost:8080`.

### Opcja B: Lokalnie (development)

```bash
cd D2ApiViewerEditor/D2ViewerEditor.Api
dotnet run
```

Upewnij się, że:
- PostgreSQL działa na `localhost:5432`
- fake-gcs działa na `localhost:4443`
- `appsettings.DEV.json` ma poprawne connection stringi

### Weryfikacja

```bash
# Swagger UI
# Otwórz w przeglądarce: http://localhost:8080/swagger

# Upload dokumentu (przez API)
curl -X POST http://localhost:8080/api/documentstorage/upload ^
  -H "Content-Type: application/json" ^
  -d "{\"name\": \"test.pdf\", \"mimeType\": \"application/pdf\", \"content\": \"SGVsbG8gV29ybGQh\", \"createdBy\": \"TestUser\"}"
```

---

## 🔧 Uruchomienie samej bazy i GCS (bez API/frontendu)

Jeśli chcesz pracować tylko z bazą i GCS (np. rozwijasz API lokalnie):

```bash
cd infra
podman-compose up -d postgres fake-gcs
```

Opcjonalnie z pgAdmin:

```bash
podman-compose up -d postgres fake-gcs pgadmin
```

---

## 📊 Struktura Bazy Danych

### Tabela `documents` (Aggregate Root)

```sql
id UUID PRIMARY KEY              -- guid_master
name VARCHAR(500) NOT NULL       -- Nazwa dokumentu
mime_type VARCHAR(255) NOT NULL  -- Typ MIME (application/pdf, image/jpeg, etc.)
created_at TIMESTAMP NOT NULL    -- Data utworzenia
created_by VARCHAR(255) NOT NULL -- Twórca (System/JWT username)
is_deleted BOOLEAN DEFAULT FALSE -- Soft delete flag
```

### Tabela `document_versions`

```sql
id UUID PRIMARY KEY              -- guid_wersji
document_id UUID FK NOT NULL     -- FK do documents.id (CASCADE)
version_number INT NOT NULL      -- Numer wersji (1, 2, 3...)
storage_path VARCHAR(500) NOT NULL -- Ścieżka w GCS (np. "documents/{id}")
size_in_bytes BIGINT NOT NULL    -- Rozmiar pliku w bajtach
created_at TIMESTAMP NOT NULL    -- Data utworzenia wersji
created_by VARCHAR(255) NOT NULL -- Twórca wersji
is_active BOOLEAN DEFAULT FALSE  -- Czy to aktywna wersja
```

> **Uwaga:** Kolumna `content BYTEA` została zastąpiona przez `storage_path` + `size_in_bytes` w migracji `003_migrate_storage_to_gcs.sql`. Zawartość binarna plików jest teraz przechowywana w GCS.

### Migracje SQL

| Plik | Opis |
|------|------|
| `sql/001_init_schema.sql` | Tabele `documents` i `document_versions` |
| `sql/002_init_indexes.sql` | Indeksy wydajnościowe, unique constraint na aktywną wersję |
| `sql/003_migrate_storage_to_gcs.sql` | Dodanie `storage_path`, `size_in_bytes`; usunięcie `content` |

---

## 🧹 Maintenance

### Backup bazy

```bash
podman exec d2viewereditor_postgres pg_dump -U postgres d2viewereditor > backup.sql
```

### Restore

```bash
podman exec -i d2viewereditor_postgres psql -U postgres -d d2viewereditor < backup.sql
```

### Pełny reset (baza + GCS)

```bash
podman-compose down -v
podman-compose up -d
```

> To usunie volume `postgres_data` i **pgadmin_data**. Dane w fake-gcs (backend=memory) i tak nie przeżywają restartu.

### Logi

```bash
# Wszystkie kontenery
podman-compose logs -f

# Tylko postgres
podman-compose logs -f postgres

# Tylko fake-gcs
podman-compose logs -f fake-gcs
```

---

## 📁 Volumes

| Volume | Opis |
|--------|------|
| `postgres_data` | Dane PostgreSQL (persystentne między restartami) |
| `pgadmin_data` | Konfiguracja pgAdmin |

```bash
# Lista volumes
podman volume ls | Select-String d2viewereditor

# Inspekcja
podman volume inspect infra_postgres_data
```

---

## 🌐 Sieć

Kontenery komunikują się wewnątrz sieci `d2viewereditor_network` (bridge). Nazwy hostów odpowiadają nazwom serwisów z `podman-compose.yml`:

| Serwis z compose | Hostname wewnętrzny | Port wewnętrzny |
|-------------------|---------------------|-----------------|
| `postgres` | `postgres` | 5432 |
| `fake-gcs` | `fake-gcs` | 4443 |
| `api` | `api` | 8080 |
| `frontend` | `frontend` | 80 |

```bash
podman network inspect d2viewereditor_network
```

---

## 🔍 Troubleshooting

### Kontener nie startuje

```bash
podman-compose logs postgres
podman-compose logs fake-gcs

# Czy port jest zajęty?
netstat -ano | findstr :5432
netstat -ano | findstr :4443
```

### Podman machine nie działa (Windows)

```powershell
podman machine stop
podman machine rm
podman machine init
podman machine start
```

### Nie można połączyć się z API

- Sprawdź `podman-compose ps` — czy postgres jest `healthy`
- Sprawdź connection string w `appsettings.DEV.json`
- Sprawdź firewall Windows (porty 5432, 4443, 8080)

### fake-gcs nie zwraca bucketów

```bash
# Sprawdź czy kontener działa
podman logs d2viewereditor_fake_gcs

# Sprawdź czy plik metadata jest zamontowany
podman exec d2viewereditor_fake_gcs ls -la /data/d2viewereditor-documents/
```

### Migracja 003 nie zadziałała (istniejąca baza)

Jeśli baza została utworzona wcześniej (przed dodaniem `003_migrate_storage_to_gcs.sql`), skrypty z `docker-entrypoint-initdb.d` **nie wykonają się ponownie** — PostgreSQL inicjalizuje je tylko przy pustym volume.

Rozwiązanie:
```bash
# Opcja 1: Ręcznie uruchom migrację
podman exec -i d2viewereditor_postgres psql -U postgres -d d2viewereditor < sql/003_migrate_storage_to_gcs.sql

# Opcja 2: Pełny reset (usuwa dane!)
podman-compose down -v
podman-compose up -d
```

---

## 📚 Więcej informacji

- [PostgreSQL Documentation](https://www.postgresql.org/docs/)
- [Podman Compose](https://github.com/containers/podman-compose)
- [fake-gcs-server (GitHub)](https://github.com/fsouza/fake-gcs-server)
- [Google Cloud Storage Client Libraries (.NET)](https://cloud.google.com/storage/docs/reference/libraries#client-libraries-install-csharp)
- [EF Core with PostgreSQL](https://www.npgsql.org/efcore/)
