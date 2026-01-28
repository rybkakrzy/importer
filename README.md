# Importer Parametryzacji

System do importu i analizy plików ZIP zawierających pliki parametryzacji.

## Struktura projektu

- **backend** - ASP.NET 8 Web API
- **frontend** - Angular (najnowsza wersja)

## Funkcjonalność

### Backend (ASP.NET 8 Web API)
- Endpoint: `POST /api/FileUpload/upload`
- Przyjmuje plik ZIP
- Analizuje zawartość archiwum
- Zwraca informacje o plikach znajdujących się w archiwum (docx, json, certyfikaty)

### Frontend (Angular)
- Komponent do wyboru pliku ZIP
- Wysyłanie pliku do API
- Wyświetlanie szczegółowej informacji o zawartości archiwum

## Uruchomienie

### Backend

```powershell
cd backend
dotnet run
```

API będzie dostępne pod adresem: `http://localhost:5190`

### Frontend

```powershell
cd frontend
npm start
```

Aplikacja będzie dostępna pod adresem: `http://localhost:4200`

## Oczekiwana struktura pliku ZIP

W archiwum ZIP powinny znajdować się:
- Pliki DOCX (dokumenty Word)
- Plik JSON z konfiguracją (struktura zostanie zdefiniowana później)
- Certyfikat (.cer, .crt, .pem, .pfx)

## Technologie

### Backend
- ASP.NET 8 Web API
- System.IO.Compression dla obsługi ZIP

### Frontend
- Angular (standalone components)
- HttpClient dla komunikacji z API
- Reactive Forms

## TODO
- Zdefiniować dokładną strukturę pliku JSON
- Dodać walidację zawartości JSON
- Dodać obsługę certyfikatów


