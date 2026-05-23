# Technology Stack

## Zasada podstawowa

Agent nie może zgadywać wersji technologii. Musi sprawdzić fakty w repozytorium.

Źródła prawdy:

- `.csproj`,
- `global.json`,
- `Directory.Build.props`,
- `Directory.Packages.props`,
- `NuGet.config`,
- `package.json`,
- `package-lock.json` / `yarn.lock` / `pnpm-lock.yaml`,
- `angular.json`,
- `tsconfig.json`,
- `Dockerfile`,
- `docker-compose.yml`,
- pliki CI/CD,
- konfiguracja środowisk.

## Backend — domyślny profil

| Obszar | Domyślna wartość | Źródło prawdy |
|---|---|---|
| Runtime | .NET / ASP.NET Core | `.csproj` / `global.json` |
| Język | C# | `.csproj` |
| API | REST API | endpointy / kontrolery |
| ORM | EF Core albo inny użyty w repo | `.csproj` |
| Walidacja | FluentValidation / DataAnnotations / custom | `.csproj` / kod |
| Auth | JWT / cookies / OAuth / OIDC | konfiguracja auth |
| Logowanie | `ILogger`, opcjonalnie Serilog/OpenTelemetry | `.csproj` / konfiguracja |
| Testy | xUnit / NUnit / MSTest | projekty testowe |

## Frontend — domyślny profil

| Obszar | Domyślna wartość | Źródło prawdy |
|---|---|---|
| Framework | Angular | `package.json` |
| Język | TypeScript | `tsconfig.json` |
| UI | Angular Material / PrimeNG / Tailwind / custom | `package.json` / kod |
| State management | Services / Signals / RxJS / NgRx | kod |
| Formularze | Reactive Forms lub obecny styl projektu | kod |
| Testy | Karma / Jest / Vitest / Cypress / Playwright | `package.json` |

## DevOps — domyślny profil

| Obszar | Domyślna wartość | Źródło prawdy |
|---|---|---|
| Konteneryzacja | Docker | `Dockerfile` / `docker-compose.yml` |
| Local env | Docker Compose + lokalne CLI | `docker-compose.yml` / README |
| CI/CD | GitHub Actions / GitLab CI / Azure DevOps | katalog `.github`, `.gitlab-ci.yml`, itp. |
| Chmura | GCP lub inne środowisko kontenerowe | IaC / dokumentacja |
| Sekrety | Secret Manager / CI secrets / env vars | konfiguracja środowisk |
| Monitoring | Cloud logs / OpenTelemetry / provider-specific | konfiguracja |

## Reguły dla zależności

- Nie dodawaj biblioteki tylko dlatego, że skraca pojedynczy fragment kodu.
- Preferuj rozwiązania już obecne w projekcie.
- Nowa zależność musi mieć uzasadnienie techniczne.
- Nie aktualizuj dużych zależności przy okazji małych tasków.
- Nie mieszaj modernizacji frameworka z implementacją funkcji.
- Nie zmieniaj lockfile bez realnej zmiany zależności.
- Jeżeli zmieniasz zależności, opisz to w `CHANGELOG.md`.

## Komendy do uzupełnienia po wrzuceniu do repo

```bash
# Backend
dotnet --info
dotnet restore
dotnet build
dotnet test

# Frontend
npm install
npm run build
npm test
npm run lint

# Docker
docker compose up -d
docker compose down
```

Usuń lub zmień komendy, które nie pasują do projektu.
