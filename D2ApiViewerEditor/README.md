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

## Testy jednostkowe

Projekt zawiera kompleksowe testy jednostkowe dla wszystkich warstw architektury.

### Struktura projektów testowych

```
D2ApiViewerEditor/
├── D2ViewerEditor.Domain.UnitTests/         # Testy modeli i Result pattern
├── D2ViewerEditor.Application.UnitTests/    # Testy handlerów CQRS i walidatorów
├── D2ViewerEditor.Infrastructure.UnitTests/ # Testy serwisów (BarcodeGenerator, etc.)
└── D2ViewerEditor.Api.UnitTests/            # Testy kontrolerów API
```

### Stack technologiczny testów

- **NUnit 3** - Framework testowy
- **FluentAssertions** - Czytelne asercje
- **NSubstitute** - Mocking framework
- **Microsoft.AspNetCore.Mvc.Testing** - Testy kontrolerów

### Uruchamianie testów

**Wszystkie testy:**
```powershell
dotnet test
```

**Konkretny projekt:**
```powershell
dotnet test D2ViewerEditor.Application.UnitTests
```

**Z code coverage:**
```powershell
dotnet test --collect:"XPlat Code Coverage"
```

**W trybie watch (ciągłe uruchamianie):**
```powershell
dotnet watch test
```

**Filtrowanie testów:**
```powershell
# Tylko testy zawierające "Barcode" w nazwie
dotnet test --filter "FullyQualifiedName~Barcode"

# Tylko testy z kategorii (jeśli są zdefiniowane)
dotnet test --filter "TestCategory=Unit"
```

### Przykłady testów

**Domain - Test wzorca Result:**
```csharp
[Test]
public void Success_ShouldCreateSuccessResult()
{
    var result = Result<string>.Success("test");
    
    result.IsSuccess.Should().BeTrue();
    result.Value.Should().Be("test");
}
```

**Application - Test handler z mockowaniem:**
```csharp
[Test]
public async Task Handle_WithValidCommand_ShouldReturnSuccessResult()
{
    var barcodeGenerator = Substitute.For<IBarcodeGenerator>();
    barcodeGenerator.Generate(Arg.Any<string>(), Arg.Any<string>(), 
        Arg.Any<int>(), Arg.Any<int>(), Arg.Any<bool>())
        .Returns(new byte[] { 0x89, 0x50, 0x4E, 0x47 });
    
    var handler = new GenerateBarcodeCommandHandler(barcodeGenerator);
    var command = new GenerateBarcodeCommand("TEST", "Code128", 300, 300, false);
    
    var result = await handler.Handle(command, CancellationToken.None);
    
    result.IsSuccess.Should().BeTrue();
}
```

**Infrastructure - Test integracyjny serwisu:**
```csharp
[Test]
public void Generate_WithValidQRCode_ShouldReturnPngBytes()
{
    var service = new BarcodeGeneratorService();
    
    var result = service.Generate("https://example.com", "QRCode", 300, 300, false);
    
    result.Should().NotBeEmpty();
    result.Take(4).Should().Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }); // PNG header
}
```

**API - Test kontrolera:**
```csharp
[Test]
public async Task GenerateBarcode_WithValidRequest_ShouldReturnOk()
{
    var mediator = Substitute.For<IMediator>();
    mediator.Send(Arg.Any<GenerateBarcodeCommand>(), Arg.Any<CancellationToken>())
        .Returns(Result<BarcodeResponse>.Success(new BarcodeResponse()));
    
    var controller = CreateController(mediator);
    
    var result = await controller.GenerateBarcode(new BarcodeRequest());
    
    result.Should().BeOfType<OkObjectResult>();
}
```

### Metryki pokrycia

Aktualny stan testów:
- **52 testy jednostkowe**
- Pokrycie wszystkich 4 warstw architektury
- Testy walidatorów (FluentValidation)
- Testy handlerów (MediatR)
- Testy serwisów domenowych
- Testy kontrolerów API

### Dodawanie nowych testów

1. Wybierz odpowiedni projekt testowy (Domain/Application/Infrastructure/Api)
2. Utwórz plik testowy w strukturze odzwierciedlającej testowany kod
3. Użyj konwencji nazewnictwa: `[Klasa]Tests.cs`
4. Implementuj testy zgodnie z wzorcem AAA (Arrange-Act-Assert)

Przykład:
```csharp
namespace D2ViewerEditor.Application.UnitTests.Features.Documents;

[TestFixture]
public class SaveDocumentCommandHandlerTests
{
    [Test]
    public async Task Handle_WithValidCommand_ShouldSaveDocument()
    {
        // Arrange
        var converter = Substitute.For<IHtmlToDocxConverter>();
        var handler = new SaveDocumentCommandHandler(converter);
        
        // Act
        var result = await handler.Handle(command, CancellationToken.None);
        
        // Assert
        result.IsSuccess.Should().BeTrue();
    }
}
```

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
