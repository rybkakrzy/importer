# D2 ViewerEditor API

Backend ASP.NET 8 Web API dla systemu do przeglądania i edycji dokumentów DOCX.

## Architektura

Projekt wykorzystuje **Clean Architecture** z podziałem na warstwy:

```
D2ApiViewerEditor/
├── D2ViewerEditor.Api/              # Warstwa prezentacji (Controllers, Middleware)
├── D2ViewerEditor.Application/      # Logika biznesowa (CQRS, Validators)
├── D2ViewerEditor.Domain/           # Encje domenowe, interfejsy
└── D2ViewerEditor.Infrastructure/   # Implementacje serwisów
```

## Funkcjonalność

- **Dokumenty DOCX**: Tworzenie, edycja, zapisywanie
- **Konwersja**: DOCX ↔ HTML (z zachowaniem formatowania)
- **Kody kreskowe**: Generowanie QR, Code128, EAN-13 (ZXing)
- **Podpisy cyfrowe**: Podpisywanie certyfikatem X.509
- **Health Check**: Endpoint monitoringu `/health`

## Technologie

- **.NET 8** + ASP.NET Core Web API
- **MediatR** - CQRS pattern
- **FluentValidation** - Walidacja żądań
- **DocumentFormat.OpenXml** - Manipulacja DOCX
- **HtmlAgilityPack** - Parsing HTML
- **ZXing.Net** - Generowanie kodów kreskowych
- **SkiaSharp** - Rendering grafiki

## Uruchomienie

### Wymagania

- .NET 8 SDK
- Visual Studio 2022 / Rider / VS Code

### Instalacja zależności

```powershell
cd D2ViewerEditor.Api
dotnet restore
```

### Uruchomienie - tryb deweloperski (DEV)

```powershell
cd D2ViewerEditor.Api
dotnet run
```

Domyślnie używa profilu `http` z pliku [Properties/launchSettings.json](D2ViewerEditor.Api/Properties/launchSettings.json).

### Uruchomienie - tryb UAT

```powershell
cd D2ViewerEditor.Api
dotnet run --launch-profile uat
```

### Uruchomienie - tryb produkcyjny (PRD)

```powershell
cd D2ViewerEditor.Api
dotnet run --environment PRD
```

### Build

```powershell
cd D2ViewerEditor.Api
dotnet build --configuration Release
```

## Konfiguracja

Pliki konfiguracyjne znajdują się w `D2ViewerEditor.Api/`:

| Plik | Środowisko | Opis |
|------|------------|------|
| `appsettings.json` | Bazowy | Wspólna konfiguracja |
| `appsettings.DEV.json` | Development | Konfiguracja deweloperska |
| `appsettings.DEV.secrets.json` | Development | Sekrety (nie commitowane) |
| `appsettings.UAT.json` | UAT | Konfiguracja testowa |
| `appsettings.UAT.secrets.json` | UAT | Sekrety testowe |
| `appsettings.PRD.json` | Production | Konfiguracja produkcyjna |
| `appsettings.PRD.secrets.json` | Production | Sekrety produkcyjne |

⚠️ **Pliki `.secrets.json` zawierają wrażliwe dane i są ignorowane przez git.**

## Punkty końcowe API

Po uruchomieniu aplikacja dostępna jest pod adresem: **http://localhost:5190**

### Swagger UI

Dokumentacja interaktywna: **http://localhost:5190/swagger**

### Główne endpointy

| Endpoint | Metoda | Opis |
|----------|--------|------|
| `/api/health` | GET | Status aplikacji |
| `/api/document/create` | POST | Tworzenie dokumentu DOCX |
| `/api/document/save` | POST | Zapisywanie dokumentu |
| `/api/document/convert/html-to-docx` | POST | Konwersja HTML → DOCX |
| `/api/document/convert/docx-to-html` | POST | Konwersja DOCX → HTML |
| `/api/barcode/generate` | POST | Generowanie kodu kreskowego |

## Tryby uruchomieniowe

### Profile w launchSettings.json

**http** - Domyślny profil deweloperski
```json
{
  "applicationUrl": "http://localhost:5190",
  "environmentVariables": {
    "ASPNETCORE_ENVIRONMENT": "DEV"
  }
}
```

**https** - Profil z HTTPS
```json
{
  "applicationUrl": "https://localhost:7190;http://localhost:5190",
  "environmentVariables": {
    "ASPNETCORE_ENVIRONMENT": "DEV"
  }
}
```

**uat** - Profil UAT
```json
{
  "applicationUrl": "http://localhost:5190",
  "environmentVariables": {
    "ASPNETCORE_ENVIRONMENT": "UAT"
  }
}
```

## Struktura projektu

### D2ViewerEditor.Api (Warstwa prezentacji)

```
Controllers/
├── BaseApiController.cs       # Bazowy kontroler
├── DocumentController.cs      # Zarządzanie dokumentami
├── BarcodeController.cs       # Generowanie kodów kreskowych
└── HealthController.cs        # Health check

Middleware/
└── ExceptionHandlingMiddleware.cs  # Obsługa błędów
```

### D2ViewerEditor.Application (Logika biznesowa)

```
Features/
├── Documents/
│   ├── Commands/
│   └── Queries/
└── Barcodes/
    └── Commands/

Validators/
└── FluentValidation validators
```

### D2ViewerEditor.Domain (Domena)

```
Models/
└── Modele domenowe

Interfaces/
└── Interfejsy serwisów
```

### D2ViewerEditor.Infrastructure (Infrastruktura)

```
Services/
├── DocumentService.cs         # Manipulacja DOCX
├── BarcodeService.cs          # Generowanie kodów
└── DigitalSignatureService.cs # Podpisy cyfrowe
```

## Testowanie

### Health Check

```powershell
curl http://localhost:5190/api/health
```

### Przykładowe żądanie (PowerShell)

```powershell
$body = @{
    content = "Test document"
    fileName = "test.docx"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:5190/api/document/create" `
    -Method POST `
    -Body $body `
    -ContentType "application/json"
```

## Logowanie

Aplikacja wykorzystuje wbudowany system logowania ASP.NET Core. Logi są wyświetlane w konsoli podczas uruchomienia.

## Troubleshooting

### Port zajęty

Jeśli port 5190 jest zajęty, zmień go w pliku `Properties/launchSettings.json`.

### Brak .NET SDK

Sprawdź wersję:
```powershell
dotnet --version
```

Zainstaluj .NET 8: https://dotnet.microsoft.com/download

### Błędy kompilacji

Wyczyść cache i przywróć zależności:
```powershell
dotnet clean
dotnet restore
dotnet build
```

## Deweloperzy

### Hot Reload

Wykorzystaj `dotnet watch` dla automatycznego hot reload:
```powershell
cd D2ViewerEditor.Api
dotnet watch run
```

### Debugowanie w VS Code

Skonfiguruj `.vscode/launch.json`:
```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": ".NET Core Launch (web)",
      "type": "coreclr",
      "request": "launch",
      "preLaunchTask": "build",
      "program": "${workspaceFolder}/D2ApiViewerEditor/D2ViewerEditor.Api/bin/Debug/net8.0/D2ViewerEditor.Api.dll",
      "args": [],
      "cwd": "${workspaceFolder}/D2ApiViewerEditor/D2ViewerEditor.Api",
      "stopAtEntry": false,
      "env": {
        "ASPNETCORE_ENVIRONMENT": "DEV"
      }
    }
  ]
}
```

## Deployment

### Build produkcyjny

```powershell
dotnet publish -c Release -o ./publish
```

### Docker (opcjonalnie)

Przykładowy Dockerfile znajduje się w katalogu `infra/`.
