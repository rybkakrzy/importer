# Technology Stack

## Zasada podstawowa

Agent nie może zgadywać wersji technologii. Musi sprawdzić fakty w repozytorium. Poniższe wartości zweryfikowano w `.csproj` / `package.json` na dzień ostatniej aktualizacji — przy wątpliwości sprawdź źródło ponownie.

Źródła prawdy: `*.csproj`, `package.json`, `angular.json`, `tsconfig.json`, `docker/*.dockerfile`, `D2ServicesViewerEditor/Dockerfile`, `appsettings*.json`, `infra/sql/`.

## Backend — `D2ApiViewerEditor` (internal API) + `D2ServicesViewerEditor` (external API)

| Obszar | Wartość | Źródło |
|---|---|---|
| Runtime | .NET 8 (`net8.0`) | `*.csproj` |
| Język | C# | `*.csproj` |
| Architektura | Clean Architecture (Domain/Application/Infrastructure/Api) | layout solucji |
| CQRS / mediator | MediatR | `Application` + behaviours |
| Walidacja | FluentValidation (wszystkie komendy mają walidatory) | `Application/Validators` |
| Auth (Internal API) | Microsoft Entra ID — **`Microsoft.Identity.Web` 3.12.0** (`AddMicrosoftIdentityWebApi`; JwtBearer dostarczany tranzytywnie, nie pinowany) + `Microsoft.Graph` 5.103.0 (app-only lookup userów) + `Azure.Identity` 1.13.1 + `Google.Cloud.SecretManager.V1` 2.6.0 (ClientSecret z GCP) | `Api.csproj`, `Program.cs` |
| ORM | EF Core 8 (`Microsoft.EntityFrameworkCore` 8.0.12) | `Infrastructure.csproj` |
| Baza | PostgreSQL via Npgsql (`Npgsql.EntityFrameworkCore.PostgreSQL` 8.0.11) | `Infrastructure.csproj` |
| Storage | Google Cloud Storage (`Google.Cloud.Storage.V1` 4.10.0); fake-gcs-server w DEV | `Infrastructure` |
| Worker w tle | `BackgroundService` (`DocumentDeliveryWorker`) + typed `HttpClient` — pakiety `Microsoft.Extensions.Hosting.Abstractions` 8.0.1, `Microsoft.Extensions.Http` 8.0.1 | `Infrastructure.csproj` |
| DOCX | `DocumentFormat.OpenXml` 3.0.2 + `HtmlAgilityPack` 1.11.61; `NPOI` 2.8.0 (dekrypcja DOCX z hasłem) | `Infrastructure.csproj` |
| Kody kreskowe / grafika | `ZXing.Net` 0.16.9 + `SkiaSharp` 3.119.2 (+ `SkiaSharp.NativeAssets.Linux`) | `Infrastructure.csproj` |
| Podpisy | RSA-SHA256, Custom XML Part w DOCX (**nie** standardowe OOXML); `System.Security.Cryptography.Xml` 10.0.6 | `DigitalSignatureService`, `Infrastructure.csproj` |
| Swagger | `Swashbuckle.AspNetCore` 6.5.0 (Internal) / 6.6.2 (External) | `*.csproj` |
| Logowanie (External API) | Serilog (`Serilog.AspNetCore` 8.0.3, sinks Console/File) + własny `GcpJsonSerilogFormatter`; Internal API: MS logging + własny `GcpJsonConsoleFormatter` | `D2ServicesViewerEditor.Api.csproj`, `Api/Logging` |
| Testy | **NUnit 4.2.2** + FluentAssertions 8.8.0 + Moq 4.20.72 / NSubstitute 5.3.0. Unit: `D2ViewerEditor.{Api,Application,Domain,Infrastructure}.UnitTests` (target `net9.0`) + `D2ServicesViewerEditor.Api.UnitTests`. Integracyjne: `D2ViewerEditor.Api.IntegrationTests` (`net8.0`, `Microsoft.AspNetCore.Mvc.Testing` — testy autoryzacji 401/403) + `D2ViewerEditor.Infrastructure.IntegrationTests`. Kod aplikacji `net8.0` | `*.Tests.csproj` |
| Benchmarki | BenchmarkDotNet (`D2ViewerEditor.Benchmarks`) | projekt |
| Przykładowy integrator | `D2ExampleExternalApp` (net8.0, port 15120) — symulator aplikacji źródłowej: upload do External API + odbiór zwrotki na własnym callbacku | `D2ExampleExternalApp/README.md` |

## Frontend — `D2GuiViewerEditor`

| Obszar | Wartość | Źródło |
|---|---|---|
| Framework | Angular 20 (`@angular/core` ^20.0.0), standalone components | `package.json` |
| Język | TypeScript ~5.8.2 | `package.json` |
| Stan | Signals (brak NgRx); RxJS ~7.8 tylko do HTTP | kod |
| Auth | MSAL — `@azure/msal-angular` ^4 + `@azure/msal-browser` ^4 (Entra ID; wymaga `npm install`) | `package.json`, `core/auth/msal.config.ts` |
| PDF | `pdfjs-dist` ^5.5.207 | `package.json` |
| Testy | **Vitest** ^3.1.1 (`ng test`) | `package.json` |
| Formatowanie | Prettier (printWidth 100, singleQuote) | `package.json` |

## DevOps

| Obszar | Wartość | Źródło |
|---|---|---|
| Konteneryzacja | Docker — `docker/api.dockerfile`, `docker/gui.dockerfile`, `D2ServicesViewerEditor/Dockerfile` | repo |
| GUI runtime | nginx 1.27-alpine, build na node:22-alpine | `docker/gui.dockerfile` + `D2GuiViewerEditor/nginx.conf` |
| API runtime | `dotnet/aspnet:8.0`, `ASPNETCORE_URLS=http://+:8080` | `docker/api.dockerfile` |
| docker-compose | **brak w repo** (nie znaleziono `docker-compose*.yml`) | recon |
| CI/CD | **brak w repo** (brak `.github/workflows`, `.gitlab-ci.yml`) | recon |
| Chmura | GCP (GCS na pliki); szczegóły deploymentu poza repo | `appsettings` + `Infrastructure` |
| Sekrety | `appsettings.{ENV}.secrets.json` — poza repo | konwencja |

## Komendy (zweryfikowane lub do potwierdzenia)

```bash
# Backend (internal API) — z katalogu D2ApiViewerEditor
dotnet build D2ViewerEditor.sln
dotnet test  D2ViewerEditor.sln

# External API
dotnet build D2ServicesViewerEditor/D2ServicesViewerEditor.Api/D2ServicesViewerEditor.Api.csproj

# Frontend — z katalogu D2GuiViewerEditor
npm install
npm start                # ng serve, port 4200
npm run build            # ng build
npm test                 # vitest
```

> Uwaga: `npm run lint` NIE jest zdefiniowany w `package.json` (są tylko: ng, start, build, watch, test). Nie zakładaj lintu.

## Reguły dla zależności

- Preferuj rozwiązania już obecne w projekcie.
- Nowa zależność wymaga uzasadnienia; opisz ją w `CHANGELOG.md`.
- Nie aktualizuj major wersji .NET ani Angulara przy okazji małych tasków.
- Nie zmieniaj lockfile bez realnej zmiany zależności.
