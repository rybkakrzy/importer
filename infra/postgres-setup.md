# Uruchomienie PostgreSQL na Podman (bez podman-compose)

## Wymagania

- Podman zainstalowany i dostępny w PATH

---

## Krok 1 — Pobierz obraz

```bash
podman pull docker.io/library/postgres:16-alpine
```

---

## Krok 2 — Utwórz volume na dane

```bash
podman volume create d2viewereditor_postgres_data
```

---

## Krok 3 — Uruchom kontener

```bash
podman run -d \
  --name d2viewereditor_postgres \
  -e POSTGRES_DB=d2viewereditor \
  -e POSTGRES_USER=postgres \
  -e POSTGRES_PASSWORD=postgres \
  -e PGDATA=/var/lib/postgresql/data/pgdata \
  -p 5432:5432 \
  -v d2viewereditor_postgres_data:/var/lib/postgresql/data \
  docker.io/library/postgres:16-alpine
```

> **Windows (cmd/PowerShell)** — zastąp `\` znakiem `` ` `` (PowerShell) lub wpisz wszystko w jednej linii:
> ```
> podman run -d --name d2viewereditor_postgres -e POSTGRES_DB=d2viewereditor -e POSTGRES_USER=postgres -e POSTGRES_PASSWORD=postgres -e PGDATA=/var/lib/postgresql/data/pgdata -p 5432:5432 -v d2viewereditor_postgres_data:/var/lib/postgresql/data docker.io/library/postgres:16-alpine
> ```

---

## Krok 4 — Sprawdź czy baza jest gotowa

```bash
podman exec -it d2viewereditor_postgres pg_isready -U postgres -d d2viewereditor
```

Oczekiwany wynik: `localhost:5432 - accepting connections`

---

## Krok 5 — Wykonaj skrypty inicjalizacyjne SQL

```bash
podman exec -i d2viewereditor_postgres psql -U postgres -d d2viewereditor < sql/001_init_schema.sql
podman exec -i d2viewereditor_postgres psql -U postgres -d d2viewereditor < sql/002_init_indexes.sql
```

> Komendy wykonuj z poziomu folderu `infra/`.

---

## Weryfikacja — sprawdź tabele

```bash
podman exec -it d2viewereditor_postgres psql -U postgres -d d2viewereditor -c "\dt"
```

Oczekiwany wynik:

```
           List of relations
 Schema |       Name        | Type  |  Owner
--------+-------------------+-------+----------
 public | document_versions | table | postgres
 public | documents         | table | postgres
```

---

## Parametry połączenia

| Parametr | Wartość             |
|----------|---------------------|
| Host     | localhost           |
| Port     | 5432                |
| Baza     | d2viewereditor      |
| User     | postgres            |
| Hasło    | postgres            |

---

## Zarządzanie kontenerem

```bash
# Zatrzymaj kontener (dane pozostają w volume)
podman stop d2viewereditor_postgres

# Uruchom ponownie
podman start d2viewereditor_postgres

# Usuń kontener (dane pozostają w volume)
podman rm d2viewereditor_postgres

# Usuń też dane (volume)
podman volume rm d2viewereditor_postgres_data
```
