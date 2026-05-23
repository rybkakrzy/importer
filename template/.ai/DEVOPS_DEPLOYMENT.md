# DevOps and Deployment

## Cel pliku

Ten plik opisuje sposób uruchamiania, budowania i wdrażania projektu.

## Środowiska

Do uzupełnienia po sprawdzeniu repo.

| Środowisko | Cel | Uwagi |
|---|---|---|
| Local | Praca developerska | Docker Compose + lokalne CLI, do potwierdzenia |
| Dev | Integracja zmian | TODO |
| Staging | Test przed produkcją | TODO |
| Production | Produkcja | TODO |

## Lokalny start

Typowy wariant do dostosowania:

```bash
# Infrastructure
docker compose up -d

# Backend
dotnet restore
dotnet run

# Frontend
npm install
npm start
```

## Docker

Agent powinien sprawdzić:

- `Dockerfile`,
- `.dockerignore`,
- `docker-compose.yml`,
- konfigurację build args,
- healthchecki,
- porty,
- wolumeny,
- zależności usług.

Zasady:

- Nie wrzucaj sekretów do obrazu.
- Nie używaj obrazu `latest` bez powodu.
- Minimalizuj finalny obraz produkcyjny.
- Nie instaluj dev dependencies w finalnym runtime image, jeżeli nie są potrzebne.
- Dbaj o deterministyczne buildy.
- Nie kopiuj całego repo do obrazu, jeżeli `.dockerignore` nie jest poprawny.

## CI/CD

Status: `TODO: opisać realny pipeline`

Typowe etapy:

1. Restore backend dependencies.
2. Build backend.
3. Test backend.
4. Install frontend dependencies.
5. Build frontend.
6. Test frontend.
7. Build Docker image.
8. Security scan, jeżeli dostępny.
9. Deploy.

## GCP

Jeżeli projekt używa Google Cloud Platform, sprawdź:

- Cloud Run / GKE / Compute Engine,
- Artifact Registry,
- Cloud SQL,
- Secret Manager,
- Cloud Build,
- IAM,
- VPC / networking,
- observability.

Zasady:

- Sekrety trzymaj w Secret Manager lub równoważnym mechanizmie.
- Uprawnienia IAM powinny być minimalne.
- Konfiguracja środowisk powinna być jawna i powtarzalna.
- Nie zmieniaj infrastruktury produkcyjnej bez osobnego zadania.
- Deployment powinien być możliwy do odtworzenia.

## Konfiguracja

Rozróżnij:

- konfigurację jawną,
- sekrety,
- konfigurację środowiskową,
- wartości developerskie,
- wartości produkcyjne.

Frontendowe environment variables nie są sekretami, bo trafiają do bundla.

## Monitoring i logi

Status: `TODO`

Minimalnie warto znać:

- gdzie trafiają logi,
- jak sprawdzić błędy produkcyjne,
- jakie są healthchecki,
- jakie metryki są krytyczne,
- jak wygląda rollback,
- jak sprawdzić aktualną wersję wdrożenia.

## Checklist przed zmianą DevOps

- Czy zmiana wpływa na produkcję?
- Czy wymaga sekretów?
- Czy można ją przetestować lokalnie?
- Czy pipeline dalej buduje backend i frontend?
- Czy obraz Docker nadal jest deterministyczny?
- Czy dokumentacja uruchomienia jest aktualna?
