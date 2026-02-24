# Doc2 — Edytor dokumentów Word online
## Materiały do prezentacji PPTX

> Każda sekcja `## Slajd X` odpowiada jednemu slajdowi w prezentacji.
> Podsekcje `###` to elementy do umieszczenia na slajdzie (nagłówki, wypunktowania, opisy).

---

## Slajd 1 — Strona tytułowa

### Tytuł
**Doc2 — Edytor dokumentów Word online**

### Podtytuł
Autorski edytor DOCX w przeglądarce oparty na Angular 19 i ASP.NET Core 8

### Elementy wizualne
- Logo ING (branding aplikacji)
- Data: Luty 2026
- Autor / zespół: *[wpisz]*

---

## Slajd 2 — Problem i motywacja

### Nagłówek
Dlaczego Doc2?

### Treść
- Potrzeba edycji dokumentów DOCX **bez instalacji** Microsoft Office
- Praca z dokumentami bezpośrednio **w przeglądarce** — zero pluginów
- Pełna **dwukierunkowa** konwersja DOCX ↔ HTML — otwierasz prawdziwy DOCX, edytujesz, zapisujesz z powrotem jako DOCX
- Dokument po zapisie otwiera się **poprawnie w Microsoft Word** — zachowane: style, tabele, nagłówki, stopki, metadane, obrazy
- Wbudowane **podpisy cyfrowe** i zarządzanie **metadanymi** dokumentu
- Branding ING Bank Śląski — gotowy do użytku korporacyjnego

---

## Slajd 3 — Architektura systemu (diagram)

### Nagłówek
Architektura — Clean Architecture + CQRS

### Diagram (do narysowania / wklejenia)
```
┌─────────────────────────────────────────────────┐
│  FRONTEND  (Angular 19, Signals, Standalone)    │
│  localhost:4200                                  │
└──────────────────┬──────────────────────────────┘
                   │ REST API (JSON + FormData)
┌──────────────────▼──────────────────────────────┐
│  API Layer  (ASP.NET Core 8, Controllers)       │
│  Swagger · CORS · Middleware · ProblemDetails    │
├─────────────────────────────────────────────────┤
│  APPLICATION Layer  (MediatR, CQRS, Validation) │
│  Commands / Queries / Handlers / Behaviours     │
├─────────────────────────────────────────────────┤
│  DOMAIN Layer  (zero zależności)                │
│  Modele · Interfejsy · Result Pattern           │
├─────────────────────────────────────────────────┤
│  INFRASTRUCTURE Layer  (implementacje)          │
│  OpenXML · HtmlAgilityPack · ZXing · SkiaSharp  │
└─────────────────────────────────────────────────┘
```

### Kluczowe zasady
- **Dependency Rule** — warstwy wewnętrzne nie znają zewnętrznych
- **Domain** nie ma żadnych paczek NuGet — czysty C#
- Cały przepływ: Controller → MediatR → Handler → Service → Result

---

## Slajd 4 — Stos technologiczny

### Nagłówek
Technologie i biblioteki

### Frontend
| Technologia | Zastosowanie |
|---|---|
| **Angular 19** | Framework UI (standalone components, signals) |
| **TypeScript** | Typowany język frontend |
| **SCSS** | Stylowanie komponentów |
| **ContentEditable** | Silnik WYSIWYG edycji |

### Backend
| Technologia | Zastosowanie |
|---|---|
| **ASP.NET Core 8** | Framework API |
| **MediatR 12** | Wzorzec Mediator / CQRS |
| **FluentValidation 11** | Walidacja zapytań |
| **DocumentFormat.OpenXml 3.0** | Odczyt / zapis DOCX |
| **HtmlAgilityPack** | Parsowanie HTML |
| **ZXing.Net + SkiaSharp** | Generowanie kodów kreskowych i QR |
| **X509Certificate2** | Podpisy cyfrowe RSA + SHA-256 |

---

## Slajd 5 — Wzorce projektowe

### Nagłówek
Wzorce projektowe w Doc2

### Lista
- **CQRS** (Command Query Responsibility Segregation) — osobne komendy (zapis) i zapytania (odczyt)
- **Mediator** (MediatR) — kontrolery nie wywołują serwisów bezpośrednio; wszystko przechodzi przez `IMediator`
- **Pipeline Behaviours** — automatyczny logging i walidacja każdego żądania
- **Result Pattern** — `Result<T>` / `Result` zamiast wyjątków do obsługi błędów biznesowych
- **Dependency Inversion** — Domain definiuje interfejsy, Infrastructure je implementuje
- **Clean Architecture** — 4 warstwy z jednokierunkowym przepływem zależności
- **Standalone Components** (Angular) — brak NgModules, każdy komponent jest samowystarczalny
- **Signals** (Angular) — reaktywny stan bez RxJS tam, gdzie to możliwe

---

## Slajd 6 — Endpointy API

### Nagłówek
REST API — pełna mapa endpointów

### Dokument
| Metoda | Endpoint | Opis |
|---|---|---|
| `POST` | `/api/Document/open` | Upload DOCX → zwraca HTML + metadane + style |
| `POST` | `/api/Document/save` | HTML → pobieranie DOCX |
| `GET` | `/api/Document/new` | Nowy pusty dokument |
| `POST` | `/api/Document/sign` | Podpisanie dokumentu certyfikatem X.509 |
| `POST` | `/api/Document/verify-signatures` | Weryfikacja podpisów w DOCX |
| `POST` | `/api/Document/upload-image` | Upload obrazu → Base64 |
| `GET` | `/api/Document/templates` | Lista szablonów dokumentów |

### Kody kreskowe
| Metoda | Endpoint | Opis |
|---|---|---|
| `POST` | `/api/Barcode/generate` | Generowanie kodu → Base64 PNG |
| `GET` | `/api/Barcode/types` | Lista obsługiwanych typów kodów |

### Inne
| Metoda | Endpoint | Opis |
|---|---|---|
| `POST` | `/api/FileUpload/upload` | Upload ZIP → ekstrakcja plików |

---

## Slajd 7 — Konwersja DOCX ↔ HTML

### Nagłówek
Dwukierunkowa konwersja DOCX ↔ HTML

### Otwieranie dokumentu (DOCX → HTML)
1. Użytkownik uploaduje plik `.docx`
2. Backend otwiera plik przez **OpenXml SDK**
3. Parsuje: akapity, tabelki, style, obrazy, nagłówki/stopki, metadane, podpisy
4. Konwertuje na **HTML** z osadzonymi obrazami (Base64 data URI)
5. Zwraca pełny `DocumentContent` { HTML, Metadata, Images, Styles, Header, Footer }

### Zapisywanie dokumentu (HTML → DOCX)
1. Frontend wysyła HTML + metadane + nagłówek + stopkę
2. Backend parsuje HTML przez **HtmlAgilityPack**
3. Tworzy nowy plik DOCX: akapity, tabele, formatowanie, obrazy
4. Ustawia metadane: core properties (PackageProperties) + extended properties (app.xml)
5. Zwraca plik `.docx` gotowy do otwarcia w **Microsoft Word**

### Klucz
> Round-trip: DOCX → HTML (edycja) → DOCX — dokument zachowuje wierność oryginałowi

---

## Slajd 8 — Interfejs użytkownika (screenshot)

### Nagłówek
Interfejs edytora — przegląd

### Elementy do zaznaczenia na screenshocie
1. **Pasek menu** — Plik, Edytuj, Wstaw, Formatuj, Narzędzia
2. **Toolbar** — formatowanie tekstu, czcionki, kolory, wyrównanie, wstawianie
3. **Obszar edycji** — wielostronicowy podgląd A4 z marginesami
4. **Pasek statusu** — numer strony, liczba słów/znaków, czas modyfikacji, suwak zoom
5. **Pasek tabeli** (kontekstowy) — pojawia się gdy kursor jest w tabeli
6. **Nagłówek/Stopka** — edytowalne regiony z konfigurowalnymi marginesami

---

## Slajd 9 — Menu i nawigacja

### Nagłówek
System menu — pełna funkcjonalność

### Plik
Nowy · Otwórz (.docx) · Zapisz · Zapisz jako... · Ustawienia strony

### Edytuj
Cofnij / Ponów · Wytnij / Kopiuj / Wklej · Wklej bez formatowania · Zaznacz wszystko · Znajdź i zamień · **Właściwości dokumentu** · **Podpisy cyfrowe**

### Wstaw
Obraz · Tabela (szybkie: 2×2–5×5 + niestandardowa) · Kod QR / kreskowy · Linia pozioma · Podział strony · Nagłówek · Stopka

### Formatuj
Tekst (B/I/U/przekreślenie/indeks) · Wyrównanie · Wcięcia · Interlinia · Listy · Orientacja · Prowadnice marginesów · Wyczyść formatowanie

### Narzędzia
Interlinia (predefiniowana + niestandardowa) · Dialog akapitu (pełny Word-style)

---

## Slajd 10 — Formatowanie tekstu

### Nagłówek
Pełne formatowanie jak w Microsoft Word

### Styl tekstu
- **Pogrubienie**, *Kursywa*, <u>Podkreślenie</u>, ~~Przekreślenie~~
- Indeks górny (x²), Indeks dolny (H₂O)
- Kolor tekstu · Kolor tła / podświetlenie

### Czcionka
- Wybór rodziny czcionki z dropdown
- Wybór rozmiaru + zwiększanie / zmniejszanie z menu

### Zmiana wielkości liter
- WIELKIE LITERY · małe litery · Jak W Tytule

### Inne
- **Format Painter** — kopiuj i wklej formatowanie
- **Wyczyść formatowanie** (`Ctrl+\`)
- Style dokumentu — Normalny, Nagłówek 1–6, Tytuł, Podtytuł, Cytat, Akapit listy

---

## Slajd 11 — Akapity i układ

### Nagłówek
Zaawansowane ustawienia akapitów

### Wyrównanie
Do lewej · Do środka · Do prawej · Wyjustuj

### Wcięcia
- Z lewej / z prawej (w cm)
- Specjalne: Pierwszy wiersz / Wiszący
- Wcięcia lustrzane (do druku dwustronnego)

### Interlinia
Pojedyncza · 1,15 · 1,5 · Podwójna · Co najmniej · Dokładnie · Wielokrotność

### Odstępy
Przed akapitem / Po akapicie (w pt)

### Dialog akapitu (styl Word)
- Zakładka 1: Wcięcia i odstępy — z podglądem na żywo (3 akapity próbne!)
- Zakładka 2: Podziały wiersza i strony — kontrola wdów/sierot, „trzymaj z następnym", łamanie strony

---

## Slajd 12 — Tabele

### Nagłówek
Zaawansowana obsługa tabel

### Wstawianie
- Szybkie wstawienie: 2×2, 3×3, 4×4, 5×5
- Dialog: do 63 kolumn × 500 wierszy
- Autodopasowanie: do zawartości / do okna / stała szerokość

### Operacje na wierszach i kolumnach
- Wstaw wiersz powyżej / poniżej
- Wstaw kolumnę z lewej / z prawej
- Usuń wiersz / kolumnę / całą tabelę

### Scalanie i dzielenie
- **Scalanie komórek** — zaznacz prostokątny zakres myszką, scal jednym kliknięciem (obsługuje colspan + rowspan)
- **Dzielenie komórek** — rozdziel scaloną lub podziel na 2 kolumny
- **Dzielenie tabeli** — rozcina tabelę w miejscu kursora

### Wygląd
- Cieniowanie komórek — 46 kolorów + niestandardowy picker
- Siatka tabeli — pokaż / ukryj linie
- Równomierne rozkładanie wierszy / kolumn

### Zaznaczanie komórek
- Niestandardowy mechanizm: przeciągnij myszką prostokątny zakres (mousedown → mousemove)
- Zaznaczenie wielokomórkowe z wizualnym podświetleniem

---

## Slajd 13 — Nagłówki, stopki i ustawienia strony

### Nagłówek
Układ dokumentu

### Nagłówek i stopka
- Edytowalne regiony z własnym HTML
- Konfigurowalne marginesy (cm od góry / dołu)
- Opcje: inna pierwsza strona, różne parzyste/nieparzyste

### Ustawienia strony
- **Predefiniowane marginesy** — Normalny, Wąski, Szeroki (z wizualnymi miniaturkami)
- **Niestandardowe marginesy** — góra/dół/lewo/prawo (cm, krok 0.1)
- **Orientacja** — Pionowa / Pozioma (z ikonami podglądu)
- **Podgląd na żywo** — miniaturowa strona z liniami treści reagująca w czasie rzeczywistym

### Prowadnice marginesów
- Wizualne linie przerywane na krawędziach marginesów
- Przełącznik: Pokaż / Ukryj

### Paginacja
- Wielostronicowy podgląd (A4 — 1122px wysokości)
- Automatyczne liczenie stron
- Wskaźnik „Strona X z Y" pojawiający się przy przewijaniu

---

## Slajd 14 — Wstawianie multimediów

### Nagłówek
Obrazy, kody kreskowe i QR

### Obrazy
- Wstawianie z pliku (konwersja do Base64 po stronie klienta)
- Osadzone bezpośrednio w HTML jako data URI
- Zachowane przy konwersji HTML ↔ DOCX

### Kody kreskowe i QR
- Generowane na backendzie (ZXing.Net + SkiaSharp)
- Typy: QR Code, Code128, EAN-13, i inne
- Opcja: pokaż wartość pod kodem
- Wstawiane jako obraz Base64

### Inne elementy
- **Linia pozioma** (`<hr>`)
- **Podział strony** — manualny page break
- **Linia podpisu** — wizualny placeholder z imieniem, stanowiskiem, datą, markerem „✕ Podpis"

---

## Slajd 15 — Znajdź i zamień

### Nagłówek
Wyszukiwanie w dokumencie

### Toolbar — zintegrowane wyszukiwanie
- Pole szukania wbudowane w pasek narzędzi
- Nawigacja: następny / poprzedni wynik
- Licznik: „3 z 15 wyników"
- Podświetlanie dopasowań w dokumencie

### Dialog Znajdź i zamień (`Ctrl+H`)
- Pole „Znajdź" + pole „Zamień na"
- Akcje: **Znajdź następny** · **Zamień** · **Zamień wszystko**

---

## Slajd 16 — Właściwości dokumentu

### Nagłówek
Zarządzanie metadanymi DOCX

### Dialog właściwości — dwie kolumny

**Informacje ogólne:**
- Tytuł · Autor · Temat · Słowa kluczowe (oddzielane przecinkami) · Opis / Komentarze · Kategoria

**Organizacja:**
- Firma · Kierownik · Status (np. „Wersja robocza", „Zatwierdzony")

**Wersjonowanie:**
- Ostatnia modyfikacja przez · Rewizja · Wersja

**Statystyki (tylko do odczytu):**
- Data utworzenia · Data modyfikacji · Liczba słów

### Zapis do DOCX
- Core Properties → `PackageProperties` (tytuł, autor, temat, słowa kluczowe, opis, kategoria, rewizja...)
- Extended Properties → `app.xml` (firma, kierownik, aplikacja)
- Wszystko zapisywane **1:1** — otwiera się prawidłowo w Microsoft Word

---

## Slajd 17 — Podpisy cyfrowe

### Nagłówek
Podpisy elektroniczne X.509

### Przegląd podpisów
- Baner informacyjny: „Ten dokument zawiera N podpis(y) cyfrowy(e)"
- Lista kart podpisów z kolorowym statusem:
  - ✅ Ważny (zielony) / ❌ Nieważny (czerwony)
  - Dane: imię, stanowisko, email, powód, certyfikat, wystawca, ważność

### Podpisywanie dokumentu
- Upload certyfikatu `.pfx` / `.p12`
- Hasło do certyfikatu
- Dane podpisującego: imię, stanowisko, email, powód
- Kliknięcie → dokument podpisany i pobrany

### Jak to działa (backend)
1. HTML → DOCX (konwersja)
2. Obliczenie SHA-256 hash z `MainDocumentPart`
3. Podpis RSA (`X509Certificate2`)
4. Zapis podpisu jako **Custom XML Part** w DOCX (namespace: `schemas.importer.app/digitalsignatures`)
5. Weryfikacja: odczyt XML → walidacja hash + certyfikat

### Linia podpisu
- Wizualny blok w dokumencie: imię, stanowisko, data, marker „✕ Podpis"
- Wstawiana bezpośrednio w treść HTML

---

## Slajd 18 — Menu kontekstowe

### Nagłówek
Menu kontekstowe (prawy przycisk myszy)

### Standardowe opcje
Wytnij · Kopiuj · Wklej · Wklej bez formatowania · Pogrubienie · Kursywa · Podkreślenie · Wyrównanie akapitu ▸ · Interlinia ▸ · Zwiększ/Zmniejsz wcięcie · Zaznacz wszystko

### Kontekst tabeli
- Gdy kursor jest w komórce tabeli → dodatkowa opcja:
- **Kolor wypełnienia komórki** ▸ — paleta 46 kolorów + picker niestandardowy

### Inteligentne pozycjonowanie
- Menu automatycznie dostosowuje pozycję aby nie wystawać poza viewport

---

## Slajd 19 — Skróty klawiaturowe

### Nagłówek
Produktywność — skróty klawiaturowe

| Skrót | Akcja |
|---|---|
| `Ctrl+Z` | Cofnij |
| `Ctrl+Y` | Ponów |
| `Ctrl+X` | Wytnij |
| `Ctrl+C` | Kopiuj |
| `Ctrl+V` | Wklej |
| `Ctrl+Shift+V` | Wklej bez formatowania |
| `Ctrl+A` | Zaznacz wszystko |
| `Ctrl+B` | Pogrubienie |
| `Ctrl+I` | Kursywa |
| `Ctrl+U` | Podkreślenie |
| `Alt+Shift+5` | Przekreślenie |
| `Ctrl+.` | Indeks górny |
| `Ctrl+,` | Indeks dolny |
| `Ctrl+\` | Wyczyść formatowanie |
| `Ctrl+H` | Znajdź i zamień |

---

## Slajd 20 — Obsługa błędów i UX

### Nagłówek
Obsługa błędów i komunikaty

### Middleware (backend)
- Globalny `ExceptionHandlingMiddleware` — łapie wszystkie wyjątki
- Zwraca **RFC 7807 ProblemDetails** (standard JSON)
- Mapowanie: `ValidationException` → 400, `NotFoundException` → 404, inne → 500
- Logowanie krytycznych błędów

### Pipeline Behaviours
- **LoggingBehaviour** — loguje nazwę i payload każdego żądania MediatR (przed/po)
- **ValidationBehaviour** — automatycznie uruchamia walidatory FluentValidation; rzuca `ValidationException` przy błędach

### Frontend UX
- **Overlay ładowania** — spinner + „Przetwarzanie..." podczas operacji async
- **Toast błędu** — czerwony, ikona ✕, auto-znika po 5s
- **Toast sukcesu** — zielony, ikona ✓, auto-znika po 3s
- **Ostrzeżenie o niezapisanych zmianach** — `confirm()` przy tworzeniu nowego dokumentu
- **Zoom** — 50% do 200%, suwak + przyciski +/−

---

## Slajd 21 — Pasek statusu i zoom

### Nagłówek
Pasek statusu

### Elementy
- 📄 **Strona X z Y** — z ikoną strony
- 📝 **Słowa: N** — liczba słów w dokumencie
- 🔤 **Znaki: N** — liczba znaków
- 🕐 **Ostatnia modyfikacja: HH:mm:ss** — czas ostatniej zmiany
- 🔍 **Zoom** — suwak zakresowy (50%–200%), przyciski −/+, wyświetlacz procentowy

### Wskaźnik strony przy przewijaniu
- Pływający tooltip „Strona X z Y" pojawiający się przy scrollowaniu
- Auto-ukrywa się po 1,5 sekundy

---

## Slajd 22 — Szablony dokumentów

### Nagłówek
Szablony startowe

### Funkcjonalność
- Siatka kart szablonów (ikona + nazwa + opis)
- Kliknięcie ładuje szablon jako aktywny dokument
- Endpoint: `GET /api/Document/templates`

### Możliwe szablony
- Pusty dokument
- Notatka służbowa
- List formalny
- Raport
- *[do rozbudowy]*

---

## Slajd 23 — Bezpieczeństwo i walidacja

### Nagłówek
Bezpieczeństwo

### Walidacja wejść
- **FluentValidation** — każdy Command/Query ma dedykowany walidator
- `GenerateBarcodeValidator` — walidacja typów kodów i danych
- `ProcessZipFileValidator` — walidacja plików ZIP
- `SaveDocumentValidator` — walidacja danych dokumentu
- `UploadImageValidator` — limity rozmiaru (10 MB), dozwolone typy

### Limity plików
- Upload DOCX: max **50 MB**
- Upload obrazu: max **10 MB**

### CORS
- Skonfigurowane dla `localhost:4200` (frontend dev)

### Podpisy cyfrowe
- RSA + SHA-256 — standardowe algorytmy kryptograficzne
- Certyfikaty X.509 z plików `.pfx` / `.p12`

---

## Slajd 24 — Podsumowanie funkcji

### Nagłówek
Doc2 — kompletny edytor DOCX w przeglądarce

### Podsumowanie w punktach
✅ Pełne formatowanie tekstu (B/I/U, kolory, czcionki, style)
✅ Zaawansowane tabele (scalanie, cieniowanie, wielokomórkowe zaznaczanie)
✅ Nagłówki i stopki z konfiguracją
✅ Wielostronicowy podgląd A4 z paginacją
✅ Ustawienia strony (marginesy, orientacja, prowadnice)
✅ Obrazy, kody kreskowe, QR
✅ Znajdź i zamień
✅ Właściwości dokumentu (pełne metadane OOXML)
✅ Podpisy cyfrowe X.509
✅ Menu kontekstowe
✅ 15 skrótów klawiaturowych
✅ Dwukierunkowa konwersja DOCX ↔ HTML
✅ Clean Architecture + CQRS na backendzie
✅ Obsługa błędów RFC 7807

---

## Slajd 25 — Co dalej? (Roadmap)

### Nagłówek
Dalszy rozwój

### Propozycje
- 📤 Export do PDF (endpoint przygotowany jako stub — 501)
- 🔄 Kolaboracja w czasie rzeczywistym (WebSocket / SignalR)
- 📊 Wykresy i obiekty OLE
- 🔒 OPC Digital Signatures (standard Word)
- 📱 Responsywny widok mobilny
- 🗄️ Integracja z systemem zarządzania dokumentami
- 📝 Śledzenie zmian (Track Changes)
- 🧪 Testy E2E (katalog `e2e/` przygotowany)
- ⚡ Testy wydajnościowe (katalog `performence/` przygotowany)

---

## Slajd 26 — Dziękuję

### Tytuł
Dziękuję za uwagę!

### Elementy
- **Doc2** — Edytor dokumentów Word online
- Logo ING
- Link do repo / demo (jeśli jest)
- Dane kontaktowe: *[wpisz]*

---

> **Wskazówki do PPTX:**
> - Slajdy 3, 7 — idealnie nadają się na diagramy / schematy blokowe
> - Slajdy 8 — umieść screenshot aplikacji z numerowanymi strzałkami
> - Slajdy 10–14 — mogą mieć ikony/ilustracje obok wypunktowań
> - Slajdy 17 — wizualizacja flow podpisywania (strzałki: upload cert → hash → sign → download)
> - Slajdy 19 — tabela skrótów dobrze wygląda na ciemnym tle
> - Slajd 24 — ikony ✅ przy każdym punkcie dają efekt „checklisty"
> - Kolorystyka sugerowana: pomarańczowy ING (#FF6200) jako kolor akcentu, białe tło, ciemnoszary tekst
