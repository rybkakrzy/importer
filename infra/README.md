# Infrastructure Setup - D2 Viewer Editor

Konfiguracja infrastruktury dla aplikacji D2 Viewer Editor (PostgreSQL + Podman).

## 📦 Wymagania

- **Podman** lub **Docker** zainstalowany
- **podman-compose** lub **docker-compose**

## 🚀 Szybki Start

### 1. Uruchomienie bazy danych

```bash
# Z poziomu głównego katalogu projektu
cd infra

# Uruchomienie z Podman
podman-compose up -d

# LUB z Docker
docker-compose up -d
```

### 2. Sprawdzenie statusu

```bash
# Status kontenerów
podman-compose ps

# Logi PostgreSQL
podman-compose logs -f postgres

# Test połączenia
podman exec -it d2viewereditor_postgres psql -U postgres -d d2viewereditor -c "\dt"
```

### 3. Zatrzymanie

```bash
# Zatrzymaj kontenery (zachowaj dane)
podman-compose stop

# Zatrzymaj i usuń kontenery (zachowaj dane w volume)
podman-compose down

# Zatrzymaj i usuń wszystko włącznie z danymi
podman-compose down -v
```

## 🗄️ PostgreSQL

### Dane dostępowe

| Środowisko | Host | Port | Database | User | Password |
|-----------|------|------|----------|------|----------|
| DEV | localhost | 5432 | d2viewereditor | postgres | postgres |
| UAT | uat-postgres-server | 5432 | d2viewereditor_uat | d2app_user | changeme_uat |
| PRD | prd-postgres-server | 5432 | d2viewereditor_prd | d2app_user | changeme_prd |

### Connection String

```
Host=localhost;Port=5432;Database=d2viewereditor;Username=postgres;Password=postgres
```

### Bezpośrednie połączenie (psql)

```bash
# Psql wewnątrz kontenera
podman exec -it d2viewereditor_postgres psql -U postgres -d d2viewereditor
```

## 🎨 pgAdmin (Opcjonalny GUI)

pgAdmin jest dostępny pod adresem: **http://localhost:5050**

### Dane logowania
- Email: `admin@d2viewereditor.local`
- Password: `admin`

### Dodanie serwera w pgAdmin

1. Otwórz http://localhost:5050
2. Add New Server
   - **General** → Name: `D2 ViewerEditor DEV`
   - **Connection**:
     - Host: `postgres` (nazwa serwisu z docker-compose)
     - Port: `5432`
     - Database: `d2viewereditor`
     - Username: `postgres`
     - Password: `postgres`

## 📊 Struktura Bazy Danych

### Tabele

#### `documents` (Aggregate Root)
```sql
id UUID PRIMARY KEY              -- guid_master
name VARCHAR(500)                -- Nazwa dokumentu
mime_type VARCHAR(255)           -- Typ MIME
created_at TIMESTAMP             -- Data utworzenia
created_by VARCHAR(255)          -- Twórca (System/JWT username)
is_deleted BOOLEAN               -- Soft delete flag
```

#### `document_versions`
```sql
id UUID PRIMARY KEY              -- guid_wersji
document_id UUID FK              -- FK do documents.id
version_number INT               -- Numer wersji (1, 2, 3...)
content BYTEA                    -- Zawartość binarna dokumentu
created_at TIMESTAMP             -- Data utworzenia wersji
created_by VARCHAR(255)          -- Twórca wersji
is_active BOOLEAN                -- Czy to aktywna wersja
```

## 🔧 Migracje (Liquibase)

Skrypty SQL znajdują się w katalogu `sql/`:

- `001_init_schema.sql` - Tworzenie tabel
- `002_init_indexes.sql` - Indeksy wydajnościowe

### Zastosowanie skryptów ręcznie

```bash
# Wykonaj skrypt w kontenerze
podman exec -i d2viewereditor_postgres psql -U postgres -d d2viewereditor < sql/001_init_schema.sql
podman exec -i d2viewereditor_postgres psql -U postgres -d d2viewereditor < sql/002_init_indexes.sql
```

### Auto-inicjalizacja

Skrypty w katalogu `sql/` są automatycznie wykonywane przy pierwszym uruchomieniu kontenera (volume mount do `/docker-entrypoint-initdb.d`).

## 🧪 Testowanie API

Po uruchomieniu bazy danych możesz przetestować API:

```bash
# Z katalogu D2ApiViewerEditor/D2ViewerEditor.Api
dotnet run --launch-profile Development

# Swagger UI dostępny: http://localhost:5001/swagger
```

### Przykładowe zapytania (curl)

```bash
# Health check
curl http://localhost:5001/health

# Upload dokumentu
curl -X POST http://localhost:5001/api/documentstorage/upload \
  -H "Content-Type: application/json" \
  -d '{
    "name": "test.pdf",
    "mimeType": "application/pdf",
    "content": "SGVsbG8gV29ybGQh",
    "createdBy": "TestUser"
  }'
```

## 🧹 Maintenance

### Backup bazy danych

```bash
# Dump całej bazy
podman exec d2viewereditor_postgres pg_dump -U postgres d2viewereditor > backup_$(date +%Y%m%d).sql

# Dump tylko danych
podman exec d2viewereditor_postgres pg_dump -U postgres -a d2viewereditor > data_backup.sql
```

### Restore

```bash
# Restore z pliku
podman exec -i d2viewereditor_postgres psql -U postgres -d d2viewereditor < backup_20241215.sql
```

### Reset bazy

```bash
# Usuń i uruchom ponownie (dane zostaną zresetowane przy auto-init)
podman-compose down -v
podman-compose up -d
```

## 📁 Volumes

Dane PostgreSQL są przechowywane w volume `postgres_data`. Nawet po `podman-compose down` dane są zachowane.

```bash
# Lista volumes
podman volume ls | grep d2viewereditor

# Inspekcja volume
podman volume inspect infra_postgres_data

# Usunięcie volume (UWAGA: usuwa wszystkie dane!)
podman volume rm infra_postgres_data
```

## 🔍 Troubleshooting

### Kontener nie startuje

```bash
# Sprawdź logi
podman-compose logs postgres

# Sprawdź czy port 5432 nie jest zajęty
netstat -ano | findstr :5432
```

### Nie można połączyć się z API

- Upewnij się, że connection string w `appsettings.DEV.json` jest poprawny
- Sprawdź czy kontener PostgreSQL jest `healthy`: `podman-compose ps`
- Sprawdź firewall (Windows może blokować port 5432)

### Wolne zapytania

```sql
-- Sprawdź wykonanie zapytań
SELECT schemaname, tablename, indexname 
FROM pg_indexes 
WHERE schemaname = 'public';

-- Statystyki tabel
SELECT * FROM pg_stat_user_tables;
```

## 🌐 Sieć

Wszystkie kontenery działają w sieci `d2viewereditor_network`:

```bash
# Inspekcja sieci
podman network inspect d2viewereditor_network
```

## 📚 Więcej informacji

- [PostgreSQL Documentation](https://www.postgresql.org/docs/)
- [Podman Compose](https://github.com/containers/podman-compose)
- [EF Core with PostgreSQL](https://www.npgsql.org/efcore/)
