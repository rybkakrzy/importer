# Importer Parametryzacji

System do importu i analizy plików ZIP zawierających pliki parametryzacji oraz edytor dokumentów Word online.

## Struktura projektu

- **backend** - ASP.NET 9 Web API
- **frontend** - Angular (najnowsza wersja)

## Funkcjonalność

### 1. Edytor Word Online

Autorski edytor dokumentów Word zbudowany w oparciu o bezpłatne rozwiązania:

#### Backend
- Endpoint: `POST /api/Document/upload` - Upload pliku DOCX i ekstrakcja zawartości
- Endpoint: `POST /api/Document/save` - Zapis edytowanej zawartości do pliku DOCX
- Parsowanie DOCX do JSON z zachowaniem formatowania (pogrubienie, kursywa, podkreślenie)
- Konwersja z JSON z powrotem do DOCX
- Wykorzystuje bibliotekę **DocumentFormat.OpenXml** (MIT license - darmowa)

#### Frontend
- Upload plików DOCX
- Wyświetlanie zawartości dokumentu z formatowaniem
- Edycja tekstu w czasie rzeczywistym (contenteditable)
- Toolbar z formatowaniem:
  - Pogrubienie (Bold)
  - Kursywa (Italic)
  - Podkreślenie (Underline)
  - Style nagłówków (Heading 1, 2, 3)
- Dodawanie/usuwanie paragrafów
- Zapis edytowanego dokumentu jako DOCX
- W pełni autorska implementacja bez zewnętrznych bibliotek WYSIWYG

### 2. Importer ZIP

#### Backend
- Endpoint: `POST /api/FileUpload/upload`
- Przyjmuje plik ZIP
- Analizuje zawartość archiwum
- Zwraca informacje o plikach znajdujących się w archiwum (docx, json, certyfikaty)

#### Frontend
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
npm install
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
- ASP.NET 9 Web API
- System.IO.Compression dla obsługi ZIP
- DocumentFormat.OpenXml dla DOCX (MIT license)

### Frontend
- Angular (standalone components)
- HttpClient dla komunikacji z API
- Signals dla zarządzania stanem
- Własna implementacja edytora (contenteditable)

## Bezpłatne rozwiązania

Wszystkie wykorzystane biblioteki są bezpłatne:
- **DocumentFormat.OpenXml** - MIT License
- **ASP.NET Core** - MIT License
- **Angular** - MIT License

## TODO
- Zdefiniować dokładną strukturę pliku JSON
- Dodać walidację zawartości JSON
- Dodać obsługę certyfikatów


