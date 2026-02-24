# Doc2 — Edytor dokumentów Word online
## Materiały do prezentacji PPTX

> Każda sekcja `## Slajd X` odpowiada jednemu slajdowi w prezentacji.
> Podsekcje `###` to elementy do umieszczenia na slajdzie (nagłówki, wypunktowania, opisy).

---

## Slajd 1 — Strona tytułowa

### Tytuł
**Doc2 — Edytor dokumentów Word online**

### Podtytuł
Autorski edytor DOCX w przeglądarce oparty na Angular 20 i ASP.NET Core 8

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

## Slajd 2A — Zalety rozwiązania Doc2

### Nagłówek
Kluczowe zalety Doc2

### Zalety techniczne
- ✅ **Niezależność od licencji Microsoft** — brak kosztów Office 365 dla każdego użytkownika
- ✅ **Pełna kontrola nad infrastrukturą** — hosting on-premises lub w dowolnej chmurze
- ✅ **Bezpieczeństwo danych** — dokumenty nie opuszczają infrastruktury firmy
- ✅ **Elastyczność wdrożenia** — konteneryzacja Docker, skalowanie horyzontalne
- ✅ **Niskie wymagania klienckie** — tylko nowoczesna przeglądarka, brak instalacji
- ✅ **Cross-platform** — działa na Windows, Linux, macOS
- ✅ **API-first design** — łatwa integracja z istniejącymi systemami
- ✅ **Kod otwarty** — pełna kontrola nad rozwojem i customizacją

### Zalety biznesowe
- 💰 **Redukcja kosztów TCO** — brak licencji SharePoint Premium, Office 365 E5
- 📈 **Szybkie wdrożenie** — gotowe do użycia po skonfigurowaniu
- 🔒 **Compliance** — pełna kontrola nad lokalizacją danych (RODO, GDPR)
- 🎨 **Branding korporacyjny** — dostosowanie UI do identyfikacji wizualnej firmy
- 🚀 **Innowacyjność** — funkcje niedostępne w Word Online (kody QR, podpisy X.509)
- 📊 **Audytowalność** — pełne logi operacji w systemie

### Zalety użytkownika końcowego
- ⚡ **Szybkość działania** — brak opóźnień Office Online
- 🎯 **Intuicyjny interfejs** — znajome menu i skróty klawiszowe jak w Word
- 📱 **Dostępność** — z każdego miejsca przez przeglądarkę
- 💾 **Offline-first ready** — możliwość dodania trybu offline (PWA)
- 🔍 **Zaawansowane wyszukiwanie** — Znajdź i zamień z licznikiem wyników

---

## Slajd 2B — Porównanie: Doc2 vs MS Office Word w SharePoint

### Nagłówek
Doc2 kontra Microsoft Word Online (SharePoint)

### Tabela porównawcza

| Funkcja | Doc2 | Word Online (SharePoint) |
|---------|------|--------------------------|
| **Licencjonowanie** | Brak dodatkowych kosztów | Wymaga SharePoint + Office 365 (od €12.50/użytkownik/mc) |
| **Hosting** | On-premises lub dowolna chmura | Tylko Microsoft Cloud (Azure) |
| **Prywatność danych** | 100% kontrola, dane w firmie | Dane w chmurze Microsoft |
| **Podpisy cyfrowe** | ✅ X.509 (RSA-SHA256) z weryfikacją | ❌ Brak natywnego wsparcia w Word Online |
| **Kody QR/kreskowe** | ✅ 13 formatów, generowane na żądanie | ❌ Wymaga dodatków (płatnych) |
| **Konwersja DOCX↔HTML** | ✅ Pełna dwukierunkowa, 100% wierność | ⚠️ Ograniczona, utrata formatowania |
| **Metadane OOXML** | ✅ Pełne (Core + Extended Properties) | ✅ Podstawowe (Core Properties) |
| **Nagłówki/stopki** | ✅ Edytowalne z marginesami | ✅ Obsługiwane |
| **Tabele zaawansowane** | ✅ Scalanie, cieniowanie 46 kolorów | ⚠️ Ograniczone formatowanie |
| **Właściwości dokumentu** | ✅ 15+ pól (firma, kierownik, status) | ⚠️ Ograniczone pola |
| **Znajdź i zamień** | ✅ Z licznikiem wyników | ✅ Podstawowe |
| **Skróty klawiszowe** | ✅ 15+ skrótów (Ctrl+B/I/U, itp.) | ✅ Częściowe wsparcie |
| **Wielostronicowy podgląd** | ✅ Automatyczna paginacja A4 | ✅ Obsługiwane |
| **Zoom** | ✅ 50%-200% z suwakiem | ✅ Skalowanie |
| **Dialog akapitu** | ✅ Pełny (wcięcia, interlinia, wdowy) | ⚠️ Uproszczony |
| **Format Painter** | ✅ Kopiowanie formatowania | ✅ Obsługiwane |
| **Style dokumentu** | ✅ 10+ stylów (Normal, Heading 1-6) | ✅ Pełne wsparcie |
| **Export do PDF** | 🚧 W przygotowaniu (endpoint gotowy) | ✅ Natywny |
| **Kolaboracja real-time** | 🚧 Roadmap (WebSocket) | ✅ Natywna (multiple users) |
| **Track Changes** | 🚧 Roadmap | ✅ Pełne wsparcie |
| **Komentarze** | 🚧 Roadmap | ✅ Pełne wsparcie |
| **Integracja API** | ✅ REST API, Swagger docs | ⚠️ Microsoft Graph (skomplikowane) |
| **Customizacja UI** | ✅ Pełna kontrola (Angular components) | ❌ Brak możliwości |
| **Tryb offline** | 🚧 Możliwy (PWA) | ❌ Wymaga połączenia |
| **Czas ładowania** | ⚡ <2s (localhost), <5s (production) | ⚠️ 5-10s (zależne od Azure) |

### Legenda
- ✅ — Pełne wsparcie
- ⚠️ — Ograniczone wsparcie
- ❌ — Brak wsparcia
- 🚧 — W przygotowaniu

### Kluczowe przewagi Doc2
1. **Unikalne funkcje**: Podpisy X.509, kody QR/kreskowe, pełne metadane OOXML
2. **Bezpieczeństwo**: Dane w infrastrukturze firmy, brak zależności od Microsoft Cloud
3. **Koszty**: Zero licencji, hosting na własnych serwerach
4. **Kontrola**: Pełna customizacja, API-first, integracja z dowolnymi systemami

### Kluczowe przewagi Word Online
1. **Kolaboracja**: Natywny real-time editing dla wielu użytkowników
2. **Ekosystem**: Głęboka integracja z Microsoft 365 (Teams, OneDrive, Outlook)
3. **Track Changes**: Śledzenie zmian z komentarzami
4. **Dojrzałość**: Wieloletnie doświadczenie, stabilność

---

## Slajd 3 — Architektura systemu (diagram)

### Nagłówek
Architektura — Clean Architecture + CQRS

### Diagram (do narysowania / wklejenia)
```
┌─────────────────────────────────────────────────┐
│  FRONTEND  (Angular 20, Signals, Standalone)    │
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
- **Nazewnictwo projektów**: D2Tools.Api · D2Tools.Application · D2Tools.Domain · D2Tools.Infrastructure

---

## Slajd 4 — Stos technologiczny

### Nagłówek
Technologie i biblioteki

### Frontend
| Technologia | Zastosowanie |
|---|---|
| **Angular 20** | Framework UI (standalone components, signals) |
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
- **Standalone Components** (Angular 20) — brak NgModules, każdy komponent jest samowystarczalny
- **Signals** (Angular 20) — reaktywny stan bez RxJS tam, gdzie to możliwe

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
- **Pogrubienie** (`Ctrl+B`) — zwiększa wagę czcionki (font-weight: bold)
- *Kursywa* (`Ctrl+I`) — pochylenie tekstu (font-style: italic)
- <u>Podkreślenie</u> (`Ctrl+U`) — linia pod tekstem (text-decoration: underline)
- ~~Przekreślenie~~ (`Alt+Shift+5`) — przekreślenie środkowe (text-decoration: line-through)
- Indeks górny (`Ctrl+.`) — tekst powyżej linii bazowej (x², E=mc²)
- Indeks dolny (`Ctrl+,`) — tekst poniżej linii bazowej (H₂O, CO₂)
- **Kolor tekstu** — paleta 46 kolorów + picker RGB
- **Kolor tła** — podświetlenie tekstu (background-color)

### Czcionka
- **Rodzina czcionki** — dropdown z systemowymi czcionkami (Arial, Times New Roman, Calibri, Courier, Verdana, Georgia, Tahoma, Trebuchet)
- **Rozmiar czcionki** — od 8pt do 72pt, regulacja +/- krok 2pt
- **Zwiększ/Zmniejsz czcionkę** — przyciski toolbar do szybkiej zmiany

### Zmiana wielkości liter
- **WIELKIE LITERY** — uppercase transformation
- **małe litery** — lowercase transformation
- **Jak W Tytule** — capitalize każde słowo (Title Case)
- **zAmIaNa** — toggle case (inverse transformation)

### Inne
- **Format Painter** — kopiuj formatowanie z zaznaczenia i zastosuj do innego tekstu
  - Kliknij Format Painter → zaznacz tekst źródłowy → zaznacz tekst docelowy
- **Wyczyść formatowanie** (`Ctrl+\`) — usuń wszystkie style inline, przywróć domyślne
- **Style dokumentu** — predefiniowane style OOXML:
  - Normal (domyślny akapit)
  - Heading 1-6 (nagłówki hierarchiczne)
  - Title (tytuł dokumentu)
  - Subtitle (podtytuł)
  - Quote (blok cytatu)
  - ListParagraph (akapit listy numerowanej/wypunktowanej)

---

## Slajd 11 — Akapity i układ

### Nagłówek
Zaawansowane ustawienia akapitów

### Wyrównanie
- **Do lewej** (`Ctrl+L`) — tekst wyrównany do lewego marginesu (text-align: left)
- **Do środka** (`Ctrl+E`) — wyśrodkowanie tekstu (text-align: center)
- **Do prawej** (`Ctrl+R`) — tekst wyrównany do prawego marginesu (text-align: right)
- **Wyjustuj** (`Ctrl+J`) — rozciągnięcie tekstu do obu marginesów (text-align: justify)

### Wcięcia
- **Z lewej** — odsunięcie całego akapitu od lewego marginesu (0-10 cm)
- **Z prawej** — odsunięcie całego akapitu od prawego marginesu (0-10 cm)
- **Pierwszy wiersz** — wcięcie tylko pierwszej linii akapitu (typografia książkowa)
- **Wiszący** — pierwszy wiersz na marginesie, reszta wcięta (listy bibliograficzne)
- **Wcięcia lustrzane** — automatyczna zamiana lewy↔prawy dla stron parzystych (druk dwustronny)

### Interlinia (line-height)
- **Pojedyncza** — 1.0 (standardowa wysokość linii)
- **1,15** — domyślna w Word 2007+ (15% więcej przestrzeni)
- **1,5** — półtora wiersza (lepsza czytelność)
- **Podwójna** — 2.0 (wymóg akademicki)
- **Co najmniej** — minimalna wartość w pt (auto-adjust dla różnych czcionek)
- **Dokładnie** — fixed wartość w pt (bez auto-adjust)
- **Wielokrotność** — niestandardowy mnożnik (np. 1.25, 2.5)

### Odstępy między akapitami
- **Przed akapitem** — space-before w punktach (0-72pt)
- **Po akapicie** — space-after w punktach (0-72pt)
- Przydatne do separacji sekcji bez pustych linii

### Dialog akapitu (styl Microsoft Word)
**Zakładka 1: Wcięcia i odstępy**
- Wizualny konfigurator z trzema akapitami podglądu
- Slider dla każdego parametru z wartościami liczbowymi
- **Podgląd na żywo** — zmiany widoczne natychmiast w przykładowym tekście

**Zakładka 2: Podziały wiersza i strony**
- **Wdowy/sieroty** — zapobieganie samotnym liniom na początku/końcu strony
- **Trzymaj z następnym** — akapit nie rozdziela się od kolejnego (np. nagłówek + treść)
- **Trzymaj wiersze razem** — cały akapit na jednej stronie
- **Podział strony przed** — wymuszenie nowej strony przed akapitem

---

## Slajd 12 — Tabele

### Nagłówek
Zaawansowana obsługa tabel

### Wstawianie
- **Szybkie wstawienie** — siatka 5×5 z podglądem na żywo (2×2, 3×3, 4×4, 5×5)
- **Dialog niestandardowy** — do 63 kolumn × 500 wierszy (ograniczenie HTML)
- **Autodopasowanie**:
  - Do zawartości — kolumny rozszerzają się według tekstu
  - Do okna — tabela wypełnia szerokość strony (100%)
  - Stała szerokość — manualna kontrola szerokości kolumn

### Operacje na wierszach i kolumnach
- **Wstaw wiersz powyżej** — dodaj nowy wiersz nad kursorem
- **Wstaw wiersz poniżej** — dodaj nowy wiersz pod kursorem
- **Wstaw kolumnę z lewej** — dodaj kolumnę po lewej stronie
- **Wstaw kolumnę z prawej** — dodaj kolumnę po prawej stronie
- **Usuń wiersz** — usuń wiersz zawierający kursor
- **Usuń kolumnę** — usuń kolumnę zawierającą kursor
- **Usuń tabelę** — usuń całą tabelę

### Scalanie i dzielenie komórek
- **Scalanie komórek** — zaznacz prostokątny zakres myszką (mousedown→mousemove), kliknij „Scal komórki"
  - Obsługa **colspan** (scalanie poziome)
  - Obsługa **rowspan** (scalanie pionowe)
  - Możliwość scalania wielokomórkowych zakresów (np. 3×2)
- **Dzielenie komórek** — rozdziel scaloną komórkę lub podziel na 2 kolumny
- **Dzielenie tabeli** — rozcina tabelę w miejscu kursora, tworząc dwie osobne tabele

### Wygląd i formatowanie
- **Cieniowanie komórek** — paleta 46 predefiniowanych kolorów + niestandardowy color picker RGB
  - Kolory ING: pomarańczowy (#FF6200), szary (#4D4D4D), niebieski (#1F1F7A)
  - Standardowe: czerwony, zielony, niebieski, żółty, fioletowy, turkusowy, różowy
- **Siatka tabeli** — przełącznik Pokaż/Ukryj linie obramowania (border visibility)
- **Równomierne rozkładanie**:
  - Wierszy — wszystkie wiersze tej samej wysokości
  - Kolumn — wszystkie kolumny tej samej szerokości

### Zaznaczanie komórek (unikalna funkcja)
- **Niestandardowy mechanizm** — przeciągnij myszką prostokątny zakres (mousedown → mousemove)
- Zaznaczenie wielokomórkowe z wizualnym podświetleniem (background-color overlay)
- Działa cross-browser (bez standardowych kontrolek tabel HTML)
- Umożliwia zaznaczenie dowolnego prostokątnego bloku do scalenia

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
- **Wstawianie z pliku** — upload pliku obrazu (JPG, PNG, GIF, BMP)
- **Konwersja do Base64** — po stronie klienta przed wysłaniem na backend
- **Osadzone jako data URI** — `<img src="data:image/png;base64,...">`
- **Zachowane przy konwersji** — DOCX→HTML: ekstrahowane z ImagePart; HTML→DOCX: osadzane jako nowy ImagePart
- **Limit rozmiaru** — maksymalnie 10 MB na obraz (walidacja FluentValidation)
- **Kompresja** — opcjonalna kompresja przed osadzeniem (do implementacji)

### Kody kreskowe i QR (13 formatów)
**Generowane na backendzie** — ZXing.Net + SkiaSharp (cross-platform)

**Obsługiwane formaty**:
1. **QR Code** — dwuwymiarowy kod, do 4296 znaków alfanumerycznych
2. **Code128** — wysokogęsty kod liniowy, alfanumeryczny
3. **EAN-13** — 13-cyfrowy kod produktowy (europejski standard)
4. **EAN-8** — 8-cyfrowy skrócony EAN
5. **UPC-A** — 12-cyfrowy kod produktowy (USA/Kanada)
6. **Code39** — alfanumeryczny, stosowany w logistyce
7. **Code93** — kompaktowa wersja Code39
8. **Aztec** — 2D kod o wysokiej gęstości
9. **Data Matrix** — 2D kod do małych przedmiotów
10. **PDF417** — 2D wielowierszowy kod (paszporty, bilety)
11. **Interleaved 2of5** — numeryczny, przemysł/magazyny
12. **Maxicode** — 2D kod UPS (logistyka)
13. **Codabar** — biblioteki, banki krwi, przesyłki

**Opcje generowania**:
- **Wartość** — tekst do zakodowania (walidacja długości per format)
- **Szerokość/wysokość** — rozmiar obrazu w pikselach (default: 300×150)
- **Pokaż wartość** — opcjonalny tekst pod kodem (human-readable)
- **Format wyjściowy** — Base64 PNG

**Wstawianie** — wygenerowany kod jako `<img>` element w dokumencie

### Inne elementy
- **Linia pozioma** (`<hr>`) — separator sekcji dokumentu (border-top: 1px solid)
- **Podział strony** — manualny page break (`<div style="page-break-after: always">`)
  - Wymusza nową stronę w wielostronicowym podglądzie
  - Zachowywany w DOCX jako `<w:br w:type="page"/>`
- **Linia podpisu** — wizualny placeholder podpisu elektronicznego
  - Pola: Imię i nazwisko, Stanowisko, Data, Marker „✕ Podpis"
  - Stylizowana ramka z kropkowaną linią
  - Używana w połączeniu z funkcją podpisu cyfrowego

---

## Slajd 15 — Znajdź i zamień

### Nagłówek
Wyszukiwanie w dokumencie

### Toolbar — zintegrowane wyszukiwanie
- **Pole szukania** — input wbudowany w pasek narzędzi (zawsze dostępny)
- **Nawigacja wyników**:
  - Przycisk „Poprzedni" (←) — przejdź do poprzedniego dopasowania
  - Przycisk „Następny" (→) — przejdź do następnego dopasowania
- **Licznik dopasowań** — „3 z 15 wyników" (aktualizacja na żywo)
- **Podświetlanie** — wszystkie dopasowania podświetlone w dokumencie (background: yellow)
- **Aktywne dopasowanie** — obecnie wyświetlane wyróżnione innym kolorem (background: orange)
- **Wrap around** — automatyczny powrót do początku/końca dokumentu przy osiągnięciu krańca

### Dialog Znajdź i zamień (`Ctrl+H`)
- **Pole „Znajdź"** — tekst do wyszukania (case-sensitive default)
- **Pole „Zamień na"** — tekst zastępczy
- **Opcje**:
  - Wielkość liter ma znaczenie (case-sensitive toggle)
  - Tylko całe wyrazy (word boundaries)
- **Akcje**:
  - **Znajdź następny** — przejdź do kolejnego dopasowania (highlight + scroll)
  - **Zamień** — zamień aktualnie wyświetlone dopasowanie i przejdź do następnego
  - **Zamień wszystko** — zamień wszystkie dopasowania jednym kliknięciem (z potwierdzeniem)
- **Licznik** — „Znaleziono X wystąpień" po zakończeniu operacji
- **Historia wyszukiwań** — dropdown z ostatnimi 10 frazami (localStorage)

### Implementacja techniczna
- **TreeWalker API** — przechodzenie przez tekstowe węzły DOM
- **Regex matching** — wyrażenia regularne dla zaawansowanych wzorców
- **Highlight z `<mark>`** — owinięcie dopasowań w tagi `<mark>` z klasami CSS
- **ScrollIntoView** — automatyczne przewijanie do aktywnego dopasowania
- **Undo/Redo support** — zamiany można cofnąć (`Ctrl+Z`)

---

## Slajd 16 — Właściwości dokumentu

### Nagłówek
Zarządzanie metadanymi DOCX

### Dialog właściwości — dwie kolumny (layout)

**Kolumna lewa — Informacje ogólne:**
- **Tytuł** — tytuł dokumentu (Core.Title)
  - Przykład: „Raport kwartalny Q4 2025"
  - Maksymalnie 255 znaków
- **Autor** — główny autor (Core.Creator)
  - Przykład: „Jan Kowalski"
  - Auto-wypełniane z systemu (opcjonalne)
- **Temat** — krótki opis tematu (Core.Subject)
  - Przykład: „Wyniki finansowe"
- **Słowa kluczowe** — tagi do wyszukiwania (Core.Keywords)
  - Przykład: „finanse, raport, Q4, 2025"
  - Oddzielane przecinkami
- **Opis / Komentarze** — dłuższy opis (Core.Description)
  - Pole tekstowe wieloliniowe (textarea)
  - Maksymalnie 1000 znaków
- **Kategoria** — klasyfikacja dokumentu (Core.Category)
  - Przykład: „Raporty finansowe"

**Kolumna prawa — Organizacja i wersjonowanie:**
- **Firma** — nazwa organizacji (Extended.Company)
  - Przykład: „ING Bank Śląski S.A."
- **Kierownik** — menedżer odpowiedzialny (Extended.Manager)
  - Przykład: „Anna Nowak"
- **Status** — stan dokumentu (Extended.DocSecurity jako tekst)
  - Dropdown: „Wersja robocza", „Weryfikacja", „Zatwierdzony", „Finalny"
- **Ostatnia modyfikacja przez** — user (Core.LastModifiedBy)
  - Auto-wypełniane przy zapisie
- **Rewizja** — numer wersji (Core.Revision)
  - Przykład: „3"
  - Auto-inkrementacja przy zapisie (opcjonalne)
- **Wersja** — semantyczna wersja (Extended.AppVersion)
  - Przykład: „1.2.0"

**Statystyki — tylko do odczytu (read-only):**
- **Data utworzenia** — timestamp (Core.Created)
  - Format: „24.02.2026 15:22:00"
- **Data modyfikacji** — ostatni zapis (Core.Modified)
  - Format: „24.02.2026 16:45:30"
- **Liczba słów** — word count (Extended.Words)
  - Obliczane przy zapisie
  - Przykład: „2847 słów"

### Zapis do DOCX — mapowanie OOXML
**Core Properties** (`/docProps/core.xml` — PackageProperties):
- `dc:title` ← Tytuł
- `dc:creator` ← Autor
- `dc:subject` ← Temat
- `cp:keywords` ← Słowa kluczowe
- `dc:description` ← Opis
- `cp:category` ← Kategoria
- `cp:lastModifiedBy` ← Ostatnia modyfikacja przez
- `cp:revision` ← Rewizja
- `dcterms:created` ← Data utworzenia
- `dcterms:modified` ← Data modyfikacji

**Extended Properties** (`/docProps/app.xml` — ExtendedFilePropertiesPart):
- `Company` ← Firma
- `Manager` ← Kierownik
- `AppVersion` ← Wersja
- `Application` ← „Doc2" (automatyczne)
- `Words` ← Liczba słów (auto-obliczane)

### Zgodność z Microsoft Word
- **100% zgodność** — wszystkie pola zapisywane w standardowych lokacjach OOXML
- **Otwieranie w Word** — właściwości widoczne w Plik → Informacje → Właściwości
- **Edycja w Word** — zmiany w Word zachowane przy ponownym otwarciu w Doc2
- **Search integration** — SharePoint/Windows Search indeksuje metadane

---

## Slajd 17 — Podpisy cyfrowe

### Nagłówek
Podpisy elektroniczne X.509

### Przegląd podpisów (Signature Verification)
- **Baner informacyjny** — „Ten dokument zawiera N podpis(y) cyfrowy(e)"
- **Lista kart podpisów** z kolorowym statusem:
  - ✅ **Ważny** (zielony) — hash dokumentu zgodny z podpisem, certyfikat zweryfikowany
  - ❌ **Nieważny** (czerwony) — hash niezgodny lub certyfikat nieprawidłowy
- **Dane wyświetlane na karcie**:
  - Imię i nazwisko podpisującego
  - Stanowisko (Title)
  - Email
  - Powód podpisania (Reason)
  - Data podpisania (timestamp ISO 8601)
  - Nazwa certyfikatu (Subject CN)
  - Wystawca certyfikatu (Issuer CN)
  - Ważność certyfikatu (NotBefore — NotAfter)
  - Algorytm: RSA-SHA256
- **Paginacja** — jeśli więcej niż 10 podpisów

### Podpisywanie dokumentu (Digital Signing)
**Kroki**:
1. **Upload certyfikatu** — plik `.pfx` lub `.p12` z kluczem prywatnym
2. **Hasło certyfikatu** — input type="password" (wymagane dla pfx)
3. **Dane podpisującego** — formularz:
   - Imię i nazwisko (required, max 100 znaków)
   - Stanowisko (optional, max 100 znaków)
   - Email (optional, format validation)
   - Powód podpisania (optional, max 200 znaków, np. "Zatwierdzenie faktury")
4. **Podpisz i pobierz** — kliknięcie wysyła żądanie, zwraca podpisany DOCX

### Jak to działa — backend (DigitalSignatureService)
**Proces podpisywania**:
1. **Konwersja HTML→DOCX** — najpierw tworzymy dokument DOCX z aktualnej treści
2. **Obliczenie hash** — SHA-256 z `MainDocumentPart.GetStream()`
   - Hash z treści dokumentu (body.xml)
   - Nie hash całego pliku ZIP (nie zależy od metadanych)
3. **Podpis RSA** — użycie `X509Certificate2.GetRSAPrivateKey().SignData()`
   - Algorytm: RSA z padding PKCS#1
   - Hash: SHA-256
4. **Serializacja do XML** — Custom XML Part:
   ```xml
   <DigitalSignature xmlns="schemas.D2Tools.app/digitalsignatures">
     <SignerName>Jan Kowalski</SignerName>
     <SignerTitle>Dyrektor Finansowy</SignerTitle>
     <SignerEmail>jan.kowalski@ing.pl</SignerEmail>
     <Reason>Zatwierdzenie dokumentu</Reason>
     <SignedDate>2026-02-24T15:22:00Z</SignedDate>
     <CertificateSubject>CN=Jan Kowalski, O=ING Bank</CertificateSubject>
     <CertificateIssuer>CN=ING Root CA, O=ING Bank</CertificateIssuer>
     <CertificateValidFrom>2025-01-01T00:00:00Z</CertificateValidFrom>
     <CertificateValidTo>2030-01-01T00:00:00Z</CertificateValidTo>
     <SignatureValue>base64-encoded-signature</SignatureValue>
     <DocumentHash>sha256-hash-hex</DocumentHash>
   </DigitalSignature>
   ```
5. **Osadzenie w DOCX** — CustomXmlPart dodany do WordprocessingDocument
6. **Zwrot pliku** — podpisany DOCX gotowy do pobrania

**Proces weryfikacji**:
1. **Odczyt Custom XML Part** — parsowanie XML z namespace `schemas.D2Tools.app/digitalsignatures`
2. **Rekonstrukcja certyfikatu** — X509Certificate2 z CertificateSubject/Issuer
3. **Obliczenie aktualnego hash** — SHA-256 z MainDocumentPart
4. **Weryfikacja podpisu** — RSA.VerifyData() porównuje hash z SignatureValue
5. **Walidacja certyfikatu** — sprawdzenie dat ważności (NotBefore/NotAfter)
6. **Zwrot wyniku** — lista SignatureVerificationResult (Valid/Invalid)

### Linia podpisu w dokumencie
- **Wizualny element** — ramka z przeznaczeniem na podpis
- **Pola**: Imię i nazwisko, Stanowisko, Data, Marker „✕ Podpis"
- **Wstawiana przez menu** — Wstaw → Linia podpisu
- **Renderowana jako HTML** — stylizowany `<div>` z border-top: dotted
- **Związek z podpisem cyfrowym** — linia podpisu jest wizualna, podpis X.509 jest kryptograficzny (osobne funkcje)

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

### Formatowanie i edycja tekstu
✅ **Pełne formatowanie tekstu** — Bold/Italic/Underline, kolory (46 predefiniowanych + picker), czcionki (8 rodzin), rozmiary (8-72pt)
✅ **Style dokumentu** — 10 predefiniowanych stylów OOXML (Normal, Heading 1-6, Title, Subtitle, Quote)
✅ **Format Painter** — kopiowanie formatowania między fragmentami tekstu
✅ **Zmiana wielkości liter** — UPPERCASE, lowercase, Title Case, toggle
✅ **Indeksy** — górny (x²) i dolny (H₂O)

### Akapity i układ strony
✅ **Zaawansowane akapity** — wyrównanie (4 typy), wcięcia (lewe/prawe/pierwsze/wiszące), interlinia (7 trybów), odstępy przed/po
✅ **Dialog akapitu** — pełny konfigurator z podglądem na żywo (3 akapity próbne) + kontrola wdów/sierot
✅ **Ustawienia strony** — marginesy (predefiniowane + niestandardowe 0.1cm krok), orientacja (pionowa/pozioma), prowadnice
✅ **Wielostronicowy podgląd** — automatyczna paginacja A4 (1122px), licznik „Strona X z Y"
✅ **Zoom** — 50%-200% z suwakiem i przyciskami +/−

### Tabele
✅ **Zaawansowane tabele** — wstawianie (szybkie 5×5 + niestandardowe do 63×500), scalanie/dzielenie komórek (colspan+rowspan)
✅ **Operacje wierszy/kolumn** — wstawianie, usuwanie, równomierne rozkładanie
✅ **Cieniowanie komórek** — 46 kolorów + picker RGB, siatka tabeli (pokaż/ukryj)
✅ **Niestandardowe zaznaczanie** — przeciągnij myszką prostokątny zakres (mousedown→mousemove)

### Nagłówki, stopki i multimedia
✅ **Nagłówki i stopki** — edytowalne regiony z konfigurowalnymi marginesami, opcja różnych dla pierwszej strony
✅ **Obrazy** — wstawianie z pliku (max 10MB), konwersja Base64, zachowane przy DOCX↔HTML
✅ **Kody kreskowe i QR** — 13 formatów (QR, Code128, EAN-13, UPC-A, Aztec, PDF417...), generowane backend (ZXing+SkiaSharp)
✅ **Linia podpisu** — wizualny placeholder z imieniem, stanowiskiem, datą, markerem „✕ Podpis"

### Wyszukiwanie i narzędzia
✅ **Znajdź i zamień** — toolbar search + dialog (`Ctrl+H`), licznik wyników, podświetlanie, wrap-around
✅ **Menu kontekstowe** — prawy przycisk myszy, 15+ opcji kontekstowych, kolor wypełnienia komórek w tabelach
✅ **15 skrótów klawiaturowych** — Ctrl+B/I/U, Ctrl+Z/Y, Ctrl+H, Ctrl+\, Ctrl+./,, Alt+Shift+5

### Metadane i podpisy
✅ **Właściwości dokumentu** — pełne metadane OOXML (15+ pól: tytuł, autor, firma, kierownik, rewizja, status, kategoria)
✅ **Podpisy cyfrowe X.509** — podpisywanie RSA-SHA256, weryfikacja certyfikatów, Custom XML Part storage
✅ **Weryfikacja podpisów** — kolorowe karty statusu (✅ Ważny / ❌ Nieważny), dane certyfikatu, issuer, ważność

### Konwersja i architektura
✅ **Dwukierunkowa konwersja DOCX ↔ HTML** — pełna wierność, zachowanie stylów, tabel, obrazów, metadanych, nagłówków/stopek
✅ **Clean Architecture** — 4 warstwy z unidirectional dependency flow (Domain → Application → Infrastructure)
✅ **CQRS + MediatR** — Commands/Queries pattern, Pipeline Behaviours (logging + validation)
✅ **Result Pattern** — Result<T> zamiast wyjątków, railway-oriented error handling
✅ **Obsługa błędów RFC 7807** — standardowe ProblemDetails, globalny middleware, mapowanie statusów HTTP

### Pasek statusu i UX
✅ **Pasek statusu** — strona X/Y, liczba słów, liczba znaków, timestamp ostatniej modyfikacji
✅ **Overlay ładowania** — spinner + „Przetwarzanie..." podczas async operations
✅ **Toast notifications** — sukces (zielony, 3s) i błąd (czerwony, 5s)

---

## Slajd 24A — Unikalne funkcje Doc2 (przewaga konkurencyjna)

### Nagłówek
Funkcje niedostępne w konkurencji

### 1. Podpisy cyfrowe X.509 (RSA-SHA256)
- **Standardy kryptograficzne** — NIST-approved algorithm
- **Custom XML storage** — podpisy osadzone w DOCX (namespace: schemas.D2Tools.app/digitalsignatures)
- **Pełna weryfikacja** — hash dokumentu + validacja certyfikatu
- **Karty podpisów** — wizualizacja z danymi certyfikatu, issuer, ważnością
- ❌ **Brak w Word Online** — wymaga desktop Word + dodatków

### 2. Generowanie kodów kreskowych i QR (13 formatów)
- **ZXing.Net + SkiaSharp** — cross-platform, działa na Linux/GCP
- **13 formatów** — QR, Code128, EAN-13, Aztec, PDF417, Data Matrix, UPC-A, Code39...
- **Backend rendering** — Base64 PNG, opcja show-value-below
- **Bezproblemowe osadzanie** — jako `<img>` w HTML/DOCX
- ❌ **Brak w Word Online** — wymaga płatnych dodatków z Marketplace

### 3. Pełne metadane OOXML (Core + Extended Properties)
- **15+ pól metadanych** — tytuł, autor, firma, kierownik, rewizja, status, kategoria, wersja...
- **Extended Properties** — firma, kierownik, aplikacja (app.xml)
- **Zgodność 100%** — otwiera się w Word z pełnymi właściwościami
- ⚠️ **Word Online** — tylko podstawowe Core Properties

### 4. Dialog akapitu z podglądem na żywo
- **3 akapity próbne** — wizualizacja zmian w czasie rzeczywistym
- **Pełna kontrola** — wcięcia, interlinia, odstępy, podziały strony
- **Zakładka 2** — wdowy/sieroty, trzymaj z następnym, podział strony przed
- ⚠️ **Word Online** — uproszczony dialog bez podglądu

### 5. Niestandardowe zaznaczanie komórek tabel
- **Mousedown→mousemove** — przeciągnij prostokątny zakres
- **Cross-browser** — działa bez standardowych kontrolek HTML table
- **Wizualne podświetlenie** — overlay background-color
- ⚠️ **Word Online** — standardowe zaznaczanie, czasem buggy

### 6. API-first design z Swagger
- **REST API** — wszystkie funkcje dostępne jako HTTP endpoints
- **Swagger UI** — interaktywna dokumentacja (localhost:5190/swagger)
- **Integracja** — łatwa integracja z istniejącymi systemami
- ⚠️ **Microsoft Graph** — skomplikowane API, wymaga OAuth

### 7. Pełna kontrola nad hostingiem
- **On-premises** — hosting na własnych serwerach
- **Dowolna chmura** — GCP, AWS, Azure, DigitalOcean...
- **Docker** — konteneryzacja, skalowanie horyzontalne
- **Kubernetes** — orkiestracja, auto-scaling
- ❌ **Word Online** — tylko Microsoft Cloud (Azure)

### 8. Zero kosztów licencjonowania
- **Brak Office 365** — nie wymaga subskrypcji Microsoft
- **Brak SharePoint** — działa samodzielnie
- **Open Source możliwy** — kod może być udostępniony wewnętrznie
- 💰 **Word Online** — wymaga od €12.50/użytkownik/miesiąc

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
- ✅ Testy E2E — Playwright + pytest-bdd (Python) w katalogu `D2E2ETools/` (Page Object Model, scenariusze BDD po polsku, testy UI + API)
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

## Slajd 27 — Porównanie: Doc2 vs alternatywne edytory

### Nagłówek
Doc2 kontra konkurencyjne edytory dokumentów

### Wprowadzenie
Porównanie Doc2 z innymi rozwiązaniami do edycji dokumentów w przeglądarce:
- **ONLYOFFICE** — open-source suite biurowy
- **Syncfusion** — komercyjna biblioteka UI
- **Apryse (dawniej PDFTron)** — SDK do dokumentów PDF/Office
- **CKEditor 5** — edytor rich-text z pluginami
- **TinyMCE** — popularny edytor WYSIWYG

### Tabela porównawcza

| Funkcja / Cecha | **Doc2** | **ONLYOFFICE** | **Syncfusion** | **Apryse** | **CKEditor 5** | **TinyMCE** |
|----------------|----------|----------------|----------------|------------|----------------|-------------|
| **Typ rozwiązania** | Pełny edytor DOCX | Suite biurowy | Biblioteka UI | PDF/Office SDK | Edytor rich-text | Edytor WYSIWYG |
| **Licencjonowanie** | ✅ Brak kosztów | ⚠️ AGPL v3 (open-source) / €1,500+ (commercial) | ❌ Płatna (od $995/dev/rok) | ❌ Płatna (od $3,000+/rok) | ⚠️ GPL / Płatna (od $3,499/rok) | ⚠️ MIT (core) / Płatna (premium od $69/mc) |
| **Hosting** | ✅ On-premises lub chmura | ✅ On-premises lub chmura | ✅ On-premises lub chmura | ✅ On-premises lub chmura | ✅ Dowolny | ✅ Dowolny |
| **Format natywny** | ✅ DOCX (OOXML) | ✅ DOCX, XLSX, PPTX | ✅ DOCX, PDF, itp. | ✅ PDF, DOCX, PPT | ❌ HTML tylko | ❌ HTML tylko |
| **Konwersja DOCX↔HTML** | ✅ Pełna dwukierunkowa | ✅ Pełna dwukierunkowa | ✅ Dobra | ✅ Bardzo dobra | ⚠️ Import przez plugin | ⚠️ Import przez plugin |
| **Zachowanie formatowania** | ✅ 100% kompatybilność z Word | ✅ Bardzo dobra (~95%) | ✅ Bardzo dobra | ✅ Doskonała | ❌ Ograniczone | ❌ Ograniczone |
| **Podpisy cyfrowe X.509** | ✅ Natywne (RSA-SHA256) | ✅ Obsługiwane | ⚠️ Wymaga custom impl. | ✅ Pełne wsparcie PDF | ❌ Brak | ❌ Brak |
| **Kody QR/kreskowe** | ✅ 13 formatów (ZXing.Net) | ❌ Wymaga dodatków | ⚠️ Wymaga custom impl. | ✅ Możliwe przez API | ❌ Wymaga pluginów | ❌ Wymaga pluginów |
| **Metadane OOXML** | ✅ Pełne (Core + Extended) | ✅ Pełne | ✅ Pełne | ✅ Pełne | ❌ Brak | ❌ Brak |
| **Nagłówki/stopki** | ✅ Pełna edycja | ✅ Pełna edycja | ✅ Pełna edycja | ✅ Pełna edycja | ❌ Brak natywnego | ❌ Brak natywnego |
| **Tabele zaawansowane** | ✅ Scalanie, 46 kolorów | ✅ Pełne wsparcie | ✅ Pełne wsparcie | ✅ Pełne wsparcie | ⚠️ Podstawowe | ⚠️ Podstawowe |
| **Kolaboracja real-time** | 🚧 Roadmap (WebSocket) | ✅ Natywna (WebSocket) | ✅ Obsługiwana | ✅ Możliwa | ✅ Przez plugin | ✅ Przez plugin |
| **Track Changes** | 🚧 Roadmap | ✅ Pełne wsparcie | ✅ Pełne wsparcie | ✅ Pełne wsparcie | ⚠️ Przez plugin | ⚠️ Przez plugin |
| **Export do PDF** | 🚧 W przygotowaniu | ✅ Natywny | ✅ Natywny | ✅ Natywny (core) | ⚠️ Wymaga backend | ⚠️ Wymaga backend |
| **Komentarze** | 🚧 Roadmap | ✅ Pełne wsparcie | ✅ Pełne wsparcie | ✅ Pełne wsparcie | ✅ Natywne | ✅ Natywne |
| **API REST** | ✅ Pełne (Swagger) | ✅ Dostępne | ✅ Dostępne | ✅ Rozbudowane | ⚠️ Ograniczone | ⚠️ Ograniczone |
| **Responsywność** | ✅ Skalowanie 50%-200% | ✅ Pełna | ✅ Pełna | ✅ Pełna | ✅ Pełna | ✅ Pełna |
| **Wielojęzyczność** | ✅ Polish + English | ✅ 40+ języków | ✅ Lokalizowalne | ✅ Lokalizowalne | ✅ Lokalizowalne | ✅ 50+ języków |
| **Wymagania backend** | .NET Core 8 | Node.js / Java / .NET | .NET / Java / JS | Różne platformy | Node.js (opcja) | PHP/Node (opcja) |
| **Wymagania frontend** | Angular 20 | Vanilla JS | Angular/React/Vue | Vanilla JS / React | Vanilla JS / React | Vanilla JS / React |
| **Rozmiar paczki (frontend)** | ~2.5 MB (gzip) | ~5 MB (gzip) | ~3 MB (gzip) | ~8 MB+ (gzip) | ~200 KB (core) | ~150 KB (core) |
| **Obsługa makr VBA** | ❌ Brak | ⚠️ Ograniczona | ❌ Brak | ⚠️ Częściowa | ❌ Brak | ❌ Brak |
| **Obsługa OLE** | 🚧 Roadmap | ✅ Częściowa | ✅ Częściowa | ✅ Dobra | ❌ Brak | ❌ Brak |

### Kluczowe różnice

#### 1️⃣ Doc2 — Autorskie rozwiązanie
- ✅ **Zero kosztów licencjonowania** — bez opłat za użytkownika
- ✅ **Pełna kontrola nad kodem** — możliwość customizacji
- ✅ **Dedykowany dla DOCX** — 100% focus na Word documents
- ✅ **Unikalne funkcje** — podpisy X.509, kody QR/kreskowe
- ✅ **Branding ING** — dostosowane do potrzeb korporacyjnych
- ⚠️ **Mniej dojrzałe** — brak niektórych zaawansowanych funkcji (track changes, kolaboracja)

#### 2️⃣ ONLYOFFICE — Kompleksowy suite
- ✅ **Pełny suite biurowy** — Word, Excel, PowerPoint
- ✅ **Dojrzały produkt** — rozwijany od 2009 roku
- ✅ **Kolaboracja** — real-time editing z wieloma użytkownikami
- ⚠️ **Licencja AGPL** — wymaga open-source całego projektu lub płatnej licencji commercial
- ⚠️ **Większe wymagania** — więcej zasobów serwera
- ⚠️ **Mniejsza elastyczność** — trudniejsza customizacja UI

#### 3️⃣ Syncfusion — Biblioteka komercyjna
- ✅ **Profesjonalne wsparcie** — płatna pomoc techniczna
- ✅ **Bogaty ekosystem** — 1,800+ komponentów UI
- ✅ **Dojrzałość** — stabilny produkt enterprise-grade
- ❌ **Wysokie koszty** — $995/developer/rok (minimum)
- ❌ **Vendor lock-in** — uzależnienie od dostawcy
- ⚠️ **Trudna integracja** — wymaga zakupu całej suite

#### 4️⃣ Apryse (PDFTron) — SDK premium
- ✅ **Najlepsza jakość** — doskonałe renderowanie PDF/DOCX
- ✅ **Zaawansowane funkcje** — annotacje, formularze, OCR
- ✅ **Wsparcie enterprise** — dedykowane dla korporacji
- ❌ **Bardzo drogie** — od $3,000+/rok + runtime fees
- ❌ **Licencjonowanie runtime** — opłaty za użytkowników końcowych
- ⚠️ **Over-engineering** — zbyt rozbudowane dla prostych przypadków

#### 5️⃣ CKEditor 5 — Edytor rich-text
- ✅ **Lekki i szybki** — mała paczka (~200 KB)
- ✅ **Popularne** — używane przez Wikipedia, GitHub
- ✅ **Modułowe** — pluginowa architektura
- ❌ **Nie DOCX-native** — tylko HTML z konwerterami
- ❌ **Ograniczone formatowanie** — brak pełnej kompatybilności Word
- ⚠️ **Wymaga pluginów** — większość funkcji to płatne dodatki

#### 6️⃣ TinyMCE — Edytor WYSIWYG
- ✅ **Najpopularniejszy** — używany przez WordPress, Jira
- ✅ **Łatwa integracja** — prosty setup (kilka linii JS)
- ✅ **Licencja MIT (core)** — darmowy dla podstawowych funkcji
- ❌ **Nie DOCX-native** — tylko HTML editor
- ❌ **Funkcje premium płatne** — PowerPaste, MergeFields (od $69/mc)
- ⚠️ **Brak zaawansowanych tabel** — ograniczone scalanie, formatowanie

### Podsumowanie

**Doc2 jest najlepszym wyborem gdy:**
- ✅ Potrzebujesz **pełnej kompatybilności z Word** (DOCX native)
- ✅ Chcesz **uniknąć kosztów licencjonowania** (zero per-user fees)
- ✅ Wymagasz **pełnej kontroli** nad hostingiem i kodem
- ✅ Potrzebujesz **unikalnych funkcji** (podpisy X.509, kody QR)
- ✅ Akceptujesz **roadmap** dla niektórych funkcji (kolaboracja, track changes)

**Alternatywy są lepsze gdy:**
- ONLYOFFICE — potrzebujesz **pełnego suite** (Excel, PowerPoint) z **kolaboracją** już teraz
- Syncfusion — potrzebujesz **wsparcia enterprise** i **SLA**
- Apryse — wymagana **najwyższa jakość** renderowania i **zaawansowane PDF**
- CKEditor 5 / TinyMCE — wystarcza **prosty edytor HTML** bez pełnej kompatybilności DOCX

---

> **Wskazówki do PPTX:**
> - Slajdy 3, 7 — idealnie nadają się na diagramy / schematy blokowe
> - Slajdy 8 — umieść screenshot aplikacji z numerowanymi strzałkami
> - Slajdy 10–14 — mogą mieć ikony/ilustracje obok wypunktowań
> - Slajdy 17 — wizualizacja flow podpisywania (strzałki: upload cert → hash → sign → download)
> - Slajdy 19 — tabela skrótów dobrze wygląda na ciemnym tle
> - Slajd 24 — ikony ✅ przy każdym punkcie dają efekt „checklisty"
> - Slajd 27 — tabela porównawcza idealnie nadaje się na wykres radarowy lub heatmap z kolorami (zielony=✅, żółty=⚠️, czerwony=❌)
> - Kolorystyka sugerowana: pomarańczowy ING (#FF6200) jako kolor akcentu, białe tło, ciemnoszary tekst
> - Struktura repozytoriów: `D2ApiTools/` (backend .NET), `D2GuiTools/` (frontend Angular 20), `D2E2ETools/` (testy E2E Playwright+BDD)
