# App Documentation - D2ViewerEditor

## 1. Co robi aplikacja

D2ViewerEditor to webowy edytor dokumentów dla workflow DOCX.
Pozwala użytkownikom:

- otwierać i edytować DOCX w przeglądarce
- zapisywać edytowany dokument z powrotem do DOCX
- zarządzać historią wersji w storage
- wstawiać kody barcode i QR
- podpisywać i weryfikować podpisy cyfrowe

## 2. Architektura wysokiego poziomu

```text
Angular GUI (D2GuiViewerEditor)
	|
	| REST
	v
ASP.NET API (D2ViewerEditor.Api)
	|
	+--> Application (CQRS handlers)
	+--> Domain (entities + interfaces)
	+--> Infrastructure (OpenXml, Barcode, Signatures, Storage)
		|
		+--> PostgreSQL (metadata + versions)
		+--> GCS/fake-gcs (binary payload)
```

## 3. Główne moduły biznesowe

### Document Module
- Otwiera uploadowany DOCX i mapuje go na model edytora HTML
- Zapisuje treść HTML do odpowiedzi z plikiem DOCX
- Obsługuje obrazy, metadata, header/footer i page margins

### Versioning Module
- Utrzymuje `masterId` jako stabilną tożsamość dokumentu
- Przechowuje niemutowalne wersje z rosnącym `version_number`
- Wymusza jedną aktywną wersję na dokument
- Przywraca historyczną wersję jako aktywną

### Barcode Module
- Zwraca wspierane typy kodów
- Generuje barcode jako base64 PNG lub obraz do pobrania

### Signature Module
- Podpisuje wynikowy DOCX danymi certyfikatu X.509
- Weryfikuje podpisy w uploadowanym dokumencie

## 4. User flow (typowy)

1. Użytkownik otwiera DOCX w GUI.
2. GUI wywołuje `POST /api/document/open`.
3. Użytkownik edytuje treść w edytorze WYSIWYG.
4. Użytkownik zapisuje dokument przez `POST /api/document/save`.
5. GUI opcjonalnie utrwala binarkę w storage przez API wersjonowania.
6. Użytkownik może przejrzeć historię wersji i przywrócić dowolną wersję.

## 5. Konfiguracja

### Backend
- `appsettings.DEV.json`
- `appsettings.UAT.json`
- `appsettings.PRD.json`

### Frontend
- `src/environments/environment.development.ts`
- `src/environments/environment.ts`

Kluczowe wartości:
- API base URL
- Build info (environment, build number, build date)
- Ustawienia storage endpoint i bucket

## 6. Runbook (local development)

### Uruchom infra
```powershell
cd infra
podman-compose up -d
```

### Uruchom API
```powershell
cd D2ApiViewerEditor\D2ViewerEditor.Api
dotnet run
```

### Uruchom GUI
```powershell
cd D2GuiViewerEditor
npm install
npm start
```

Punkty dostępu:
- GUI: `http://localhost:4200`
- API: `http://localhost:5190`
- Swagger: `http://localhost:5190/swagger`

## 7. Znane luki / planowane prace

- Endpoint PDF export jest placeholderem (`501 Not Implemented`)
- Real-time collaborative editing nie jest zaimplementowany
- Zaawansowany enterprise IAM/SSO nie jest udokumentowany w aktualnym codebase

