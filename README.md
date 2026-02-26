# Importer Parametryzacji

System do importu i analizy plików ZIP zawierających pliki parametryzacji.

## Struktura projektu

- **backend** - ASP.NET 8 Web API
- **frontend** - Angular (najnowsza wersja)
  - **projects/shared** - Współdzielona biblioteka Angular (`@importer/shared`)

## Współdzielona biblioteka Angular (`@importer/shared`)

Biblioteka `projects/shared` rozwiązuje problem duplikowania kodu między wieloma projektami
Angular. Zamiast pisać ten sam kod (style, serwisy, komponenty) trzy razy, każdy projekt
importuje gotowe elementy z tej biblioteki.

### Zawartość biblioteki

| Ścieżka | Opis |
|---------|------|
| `src/lib/styles/_variables.css` | Zmienne CSS (kolory, odstępy, typografia) — design tokens marki |
| `src/lib/button/` | `ButtonComponent` (`<lib-button>`) – przycisk z wariantami `primary / success / danger` |
| `src/lib/error-message/` | `ErrorMessageComponent` (`<lib-error-message>`) – wyświetlanie błędów |

### Jak używać biblioteki w innych projektach Angular

**Opcja A – monorepozytorium (zalecane)**

Dodaj projekt do tego samego workspace Angular (patrz `angular.json`). W `tsconfig.json`
projektu dodaj alias ścieżki:

```jsonc
// tsconfig.json
{
  "compilerOptions": {
    "baseUrl": ".",
    "paths": {
      "@importer/shared": ["projects/shared/src/public-api.ts"]
    }
  }
}
```

Następnie importuj komponenty/serwisy bezpośrednio:

```typescript
import { ButtonComponent, ErrorMessageComponent } from '@importer/shared';
```

**Opcja B – publikacja jako paczka npm (prywatny rejestr / npm link)**

Zbuduj bibliotekę:

```powershell
cd frontend
ng build shared
```

Opublikuj `dist/shared` do prywatnego rejestru npm lub użyj `npm link`, a następnie dodaj
`@importer/shared` do `dependencies` w docelowym projekcie.

> **Dlaczego nie subrepo?**
> Git submodules / subrepos skomplikują pipeline CI/CD i wymuszają ręczne aktualizacje
> wskaźnika commita. Natywna biblioteka Angular w jednym workspace to prostsze i lepiej
> wspierane przez narzędzia rozwiązanie.

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


