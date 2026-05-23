# DevOps and Deployment

## Cel pliku

Sposób uruchamiania, budowania i wdrażania projektu. Część szczegółów deploymentu jest poza repo — oznaczone jako „do potwierdzenia".

## Obrazy Docker (w repo)

| Obraz | Plik | Bazowy | Port | Uwagi |
|---|---|---|---|---|
| Internal API | `docker/api.dockerfile` | `dotnet/aspnet:8.0` (build `sdk:8.0`) | `ASPNETCORE_URLS=http://+:8080`, EXPOSE 8080 | `ASPNETCORE_ENVIRONMENT=DEV`, ENTRYPOINT `D2ViewerEditor.Api.dll` |
| GUI | `docker/gui.dockerfile` | build `node:22-alpine` → runtime `nginx:1.27-alpine` | EXPOSE 80 | kopiuje `D2GuiViewerEditor/nginx.conf` + `dist/d2-gui-viewereditor/browser` |
| External API | `D2ServicesViewerEditor/Dockerfile` | `dotnet/aspnet:8.0` | EXPOSE 80/443 | ENTRYPOINT `D2ServicesViewerEditor.Api.dll` |

## Porty DEV (launchSettings)

| Usługa | HTTP | HTTPS |
|---|---|---|
| Internal API (`D2ApiViewerEditor`) | 5190 | 7190 |
| External API (`D2ServicesViewerEditor`) | 5000 | 7000 |
| GUI (`ng serve`) | 4200 | — |

> Aplikacja zewnętrzna ma rzekomo wołać External API na porcie **15112** (z rozmowy projektowej), ale repo (`launchSettings`) tego nie potwierdza — patrz `RISKS_ASSUMPTIONS.md` A-01.

## Lokalny start (do potwierdzenia — brak docker-compose w repo)

```bash
# Wymagane zależności zewnętrzne (uruchamiane poza repo):
#  - PostgreSQL  (schemat: uruchom skrypty infra/sql/ w kolejności 001..005)
#  - GCS lub fake-gcs-server (bucket: d2viewereditor-documents)

# Internal API
cd D2ApiViewerEditor && dotnet run --project D2ViewerEditor.Api

# External API
cd D2ServicesViewerEditor && dotnet run --project D2ServicesViewerEditor.Api

# GUI
cd D2GuiViewerEditor && npm install && npm start   # http://localhost:4200
```

## CI/CD

**Brak w repo** — nie znaleziono `.github/workflows`, `.gitlab-ci.yml` ani innego pipeline. Do uzupełnienia po ustaleniu z zespołem. Typowy łańcuch: restore → build → test (backend NUnit, frontend Vitest) → build obrazów Docker → deploy.

## GCP

- GCS używany na pliki binarne (bucket `d2viewereditor-documents`); w DEV fake-gcs-server.
- Pozostała infrastruktura (Cloud Run/GKE, Cloud SQL, Secret Manager, IAM) — do potwierdzenia, brak IaC w repo.

## Konfiguracja vs sekrety

- Jawna konfiguracja: `appsettings.json`, `appsettings.{ENV}.json` (Internal: DEV/UAT/PRD; External: dev/tst/acc/prd/local).
- Sekrety: `appsettings.{ENV}.secrets.json` — poza repo.
- Frontend `environment.*` nie jest sekretem (trafia do bundla); prod `apiUrl` wskazuje localhost → prawdopodobnie podmieniany przy deployu.

## Checklist przed zmianą DevOps

- Czy zmiana wpływa na produkcję? Czy wymaga sekretów?
- Czy da się przetestować lokalnie?
- Czy obrazy Docker nadal budują się deterministycznie?
- Czy dokumentacja uruchomienia jest aktualna?
