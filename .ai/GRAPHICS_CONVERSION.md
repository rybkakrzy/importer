# Konwersja grafik dokumentowych (VML / EMF / WMF / TIFF → web)

Usługa konwertująca legacy/nie-natywne grafiki Worda do formatu wyświetlanego w przeglądarce, z
łańcuchem strategii, **przezroczystym blankiem zamiast widocznego placeholdera** i zachowaniem
wierności w DOCX przez pass-through. Wprowadzone 2026-06-04; przebudowa „no-placeholder" + cache +
TIFF/WEBP/ICO 2026-06-22 (**ADR-0020**).

> **Bez widocznego placeholdera.** Gdy żadna realna strategia nie zwróci rastra, wynik to
> **przezroczysty, pusty SVG** (`IsBlankFallback`, wymiary z layoutu) — NIGDY szare tło ani tekst
> „… — podgląd w Word". Niepowodzenie raportowane strukturalnie w `GraphicConversionDiagnostics`
> (`Status` / `AttemptedStrategies` / `FailureReason`), nie w obrazie. Oryginał jedzie do DOCX.

## 1. Problem (stan przed)

`DocxToHtmlConverter.ConvertDrawingToHtml` (a:blip) i `ConvertPictureToHtml` (VML v:imagedata)
emitowały `<img src="data:{contentType};base64,…">` z **surowym** content-type media partu. Dla
EMF/WMF dawało to `data:image/x-emf;base64,…` / `image/x-wmf` — **przeglądarka tego nie renderuje**
→ złamany obraz w edytorze. VML kształty wektorowe (bez imagedata) nie były obsługiwane.

## 2. Architektura

```
Domain/Interfaces/IGraphicConversionService.cs   ← port (Detect / ConvertForEditor / ConvertVmlShapeForEditor / SanitizeSvg)
Domain/Models/GraphicConversionModels.cs          ← GraphicSource, WebGraphicRepresentation, GraphicConversionResult, GraphicConversionDiagnostics, enumy
Infrastructure/Services/GraphicConversionService.cs ← implementacja (pure-managed)
DocxToHtmlConverter.WebGraphicForLegacy(...)        ← hook: EMF/WMF media part → renderowalny data:URL
```

- **Wykrywanie** (`Detect`): magic bytes (PNG/JPEG/GIF/BMP/EMF/WMF/SVG) z fallbackiem na content-type.
- **Wymiary**: parsery nagłówków bez zależności — PNG IHDR, JPEG SOFx, GIF/BMP header, **EMF**
  `rclFrame` (0.01 mm), **WMF** Aldus placeable bbox/inch. EMU→px = `/9525`.
- **Wynik**: `WebGraphicRepresentation` (mime + bajty + wymiary + `IsPlaceholder`) → `ToDataUrl()`.
- **Diagnostyka**: `InputKind`, `OutputMimeType`, `Status` (Converted/PassThrough/Fallback/Unsupported/Rejected),
  `Fidelity` (Lossless/Lossy/Fallback/Unsupported), `ElapsedMs`, `Warnings`, `LostProperties`.

## 3. Formaty

| Wejście | Strategia | Wyjście | Fidelity |
|---|---|---|---|
| PNG/JPEG/GIF/BMP/**WEBP/ICO** | passthrough (web-native) | ten sam raster (data URL) | Lossless |
| SVG | sanitizacja | SVG bezpieczny | Lossless/blank |
| **TIFF** | dekoder **SkiaSharp** → PNG (gdy kodek dostępny), inaczej **przezroczysty blank** | PNG / blank | Lossy/Fallback |
| **EMF/WMF** z osadzonym PNG/JPEG | ekstrakcja PNG/JPEG z bajtów | osadzony raster | Lossy |
| **EMF/WMF** z osadzonym DIB (StretchDIBits itp.) | DIB→BMP→**SkiaSharp**→PNG (od 2026-06-11) | PNG | Lossy |
| **EMF/WMF** bez rastra (czysty wektor) | **własny tłumacz wektorowy → SVG** (etap 1, ADR-0027: kształty/poly/ścieżki/transformacje/pióra/pędzle); gdy brak obsługiwanych rekordów → przezroczysty blank. Zawsze + **pass-through oryginału** | SVG / blank | Lossy / Fallback |
| **VML** rect/oval/line/roundrect | mapowanie bezpiecznego podzbioru | SVG (fill/stroke) | Lossy |
| **VML** v:imagedata | rozwiązanie partu → ścieżka EMF/WMF/raster | jw. | jw. |
| **EMZ/WMZ** (metafile GZIP, `image/x-emz`/`x-wmz`) | dekompresja GZIP (bounded do `MaxInputBytes`) → dalej jak EMF/WMF | jw. | jw. |
| VML inne / nieznane media | raster-rescue SkiaSharp, inaczej **przezroczysty blank** + diagnostyka | blank (niewidoczny) | Unsupported |

Łańcuch strategii dla EMF/WMF (kolejność, raportowany w `AttemptedStrategies`):
`embedded-raster` → `dib-rasterize` → **`vector-translate`** (etap 1 własnego tłumacza
`MetafileVectorTranslator`, ADR-0027) → blank. Dla TIFF/Unknown: `skia-decode` → blank.
Dla GZIP: `gzip-decompress` poprzedza detekcję (2026-07-05). Reader (`WebGraphicForLegacy`)
kieruje do konwersji także `GraphicKind.Unknown` — `src` NIGDY nie dostaje nierenderowalnego
`data:{contentType}` (wcześniej EMZ/nieznane = ikona złamanego obrazka w edytorze).

## 4. Dlaczego BEZ LibreOffice / GDI / System.Drawing

- **LibreOffice/soffice** — zakazane (wymaga procesu zewnętrznego, ciężka instalacja, niestabilne w kontenerze).
- **System.Drawing.Common** — od .NET 6 **Windows-only**; na Linux/GCP rzuca `PlatformNotSupported`. Reguła 7.
- **GDI/EMF rendering** — brak bezpiecznej, wieloplatformowej, pure-managed ścieżki rasteryzacji EMF/EMF+/WMF
  o akceptowalnej licencji. Magick.NET/ImageMagick deleguje WMF/EMF do natywnych delegatów (libwmf/GDI),
  nieobecnych w obrazie kontenera → nieprzewidywalne.

**Decyzja:** **zero nowych zależności graficznych.** Wszystko pure-managed (`System.Buffers.Binary`,
`System.Xml`). EMF/WMF **nie są rasteryzowane** — pokazujemy osadzony raster (gdy jest) albo
placeholder, a **oryginalny part jedzie do DOCX bez zmian** (pass-through, przez `ConvertPreservingPackage`),
więc Word renderuje prawdziwą grafikę i dokument nie zostaje uszkodzony. To honest fallback, nie udawanie.

### USUNIĘTO poprzednią ścieżkę EMF (soffice + System.Drawing)
Reader **wcześniej** konwertował EMF/WMF przez `TryConvertMetafileToPng` → **LibreOffice `soffice`**
(zakazane) → fallback **`System.Drawing`** (Windows-only; na .NET 8 rzuca `PlatformNotSupported` na
Linux). Działało wyłącznie na Windows deva, **nie na GCP**. Cały ten kod (`TryConvertMetafileToPng`,
`ResolveSofficeBinary`, `ProbeBinary`, `TryConvertWithLibreOffice`, `ConvertMetafileToPngWindows`,
`LooksLikeWmf`) **usunięto**, a oba miejsca wywołań (`LoadImageFromPart`, picture-bullet loader)
przepięto na pure-managed `GraphicConversionService`. **Pakiet `System.Drawing.Common` usunięty z
`.csproj`.** Pozostały `SkiaSharp`(+`NativeAssets.Linux`) jest GCP-safe i używany tylko przez generator
kodów kreskowych. Efekt: identyczne zachowanie EMF/WMF na Windows i Linux/GCP.

### Roadmapa rasteryzacji (poza MVP)
Out-of-process sidecar/worker z permisywnym rendererem EMF (np. dedykowana usługa konwersji w
osobnym kontenerze) — uruchamiany asynchronicznie, cache wynikowego PNG. Nie dodano (infrastruktura).

## 5. Zapis zwrotny do DOCX

- **Nowe/edytowane grafiki z edytora** (PNG/JPEG): istniejąca ścieżka writera → media part + relacja
  + `w:drawing/wp:inline` (rozmiar z `data-*-emu`). SVG z edytora: strategia = osadzić PNG fallback
  (Word wymaga `a:blip` rastra obok SVG) — **roadmapa** (obecnie edytor wstawia rastry).
- **Legacy EMF/WMF/VML nieedytowane**: **pass-through** — `ConvertPreservingPackage` zachowuje oryginalny
  pakiet (media parts + relacje + content types), więc metafile zostaje nietknięty i Word go renderuje.
  Placeholder w edytorze jest tylko podglądem (atrybut `data-legacy-graphic="placeholder"` na `<img>`),
  **nie** trafia do zapisu jako treść.
- Brak gubienia relacji/rozmiaru/pozycji: hook zmienia wyłącznie `src` na renderowalny; `data-*-emu`,
  anchor/wrap/border/crop pozostają (round-trip jak dla rastrów).

## 6. Bezpieczeństwo (niezaufane wejście)

- XML/SVG: `XmlReaderSettings { DtdProcessing=Prohibit, XmlResolver=null, MaxCharactersFromEntities=0 }`
  → blokada XXE / billion-laughs / zewnętrznych encji.
- `SanitizeSvg`: usuwa `script/foreignObject/iframe/use/animate*/handler`, atrybuty `on*`,
  `href/xlink:href/src` spoza `data:`/`#`, wartości `javascript:`.
- Limity: `MaxInputBytes` (32 MB), `MaxPlaceholderWidthPx/HeightPx` (2000, clamp), `Timeout` (5 s,
  `CancellationTokenSource.CancelAfter`). Oversize/empty → `Rejected`. Wyjątki na danych → `Fallback`
  (nigdy nie wywraca importu). Brak logowania pełnych dokumentów/danych.

## 7. Cache / deduplikacja (zaimplementowane 2026-06-22)

Konwersja jest deterministyczna i czysta → cache po **hashu treści media partu**. Klucz =
`SHA-256(bytes) + content-type + origin + TargetW/H EMU + limity opcji` (`BuildCacheKey`). Magazyn:
`ConcurrentDictionary` w obrębie instancji `GraphicConversionService` (brak globalnego locka),
**bounded** do `MaxCacheEntries=1024` (po limicie przestajemy dokładać → ograniczona pamięć dla
dokumentów z setkami unikalnych grafik). Identyczny asset wstawiony N razy w dokumencie konwertowany
**raz** (ten sam, niezmienny obiekt wyniku). Klucz dostępny w `Diagnostics.CacheKey`. Invalidacja
naturalna (hash treści). Test: `IdenticalInput_IsDeduplicated_FromContentHashCache`.

## 8. Testy

Uruchom: `dotnet test D2ViewerEditor.Infrastructure.UnitTests --filter GraphicConversion` (35).

| Plik | Zakres |
|---|---|
| `GraphicConversionServiceTests` | detekcja PNG/JPEG/GIF/BMP/EMF/WMF/SVG/**TIFF/WEBP/ICO**/Unknown, PNG passthrough, EMF/WMF wymiary+**przezroczysty blank (regresja no-placeholder: brak `<text`/`fill`/„requires conversion")**, ekstrakcja osadzonego PNG, DIB→PNG, **dedup po hashu + CacheKey + AttemptedStrategies/FailureReason**, TIFF→blank, VML rect/oval/line + imagedata→null + unsupported→null, limity |
| `GraphicConversionSecurityTests` | sanitizacja SVG (script/on*/external/javascript), XXE/DTD blok, malformed SVG/EMF → fallback bez wyjątku, clamp wymiarów |
| `GraphicConversionIntegrationTests` | DOCX EMF/**WMF**/**mixed native+EMF**/**duplicated EMF**: renderowalny `src` (NIGDY raw x-emf/x-wmf), `data-legacy-graphic="blank"`, skan HTML/SVG bez widocznego placeholdera, round-trip EMF→valid DOCX z partem EMF |
| `GraphicConversionIntegrationTests` | DOCX z EMF a:blip → reader emituje `data:image/svg+xml` + `data-legacy-graphic="placeholder"`, **nigdy** `data:image/x-emf` |
| `ImageImportRegressionTests` (2026-07-05) | kolizja rId body↔header (obraz nagłówka ≠ obraz body), `mc:AlternateContent` Choice/Fallback renderowany, EMZ (gzip) z osadzonym PNG → PNG w src / czysty wektor → blank, gunzip w serwisie (`gzip-decompress`), zerowy `wp:extent` → wymiary intrinsic, writer zapisuje prawdziwy content type (TIFF/WEBP), `r:link`-only nie wywraca importu |
| `MetafileVectorTranslationTests` (2026-07-05, ADR-0027) | EMF: rect+pen/brush → `<rect>` z kolorami, POLYGON16, MoveTo/LineTo → `<line>`, ścieżka Begin/Close/StrokePath → `<path>`, SetWorldTransform (translacja), header-only → nadal blank (tłumacz nie fabrykuje treści), nieznane rekordy → blank bez wyjątku; WMF: rect+pen/brush, polygon; integracja: DOCX z wektorowym EMF → SVG w `src`, brak atrybutu blank, eksport = oryginalny part EMF (nie SVG) |

## 9. Benchmarki

`GraphicConversionBenchmarks` (BenchmarkDotNet, `[MemoryDiagnoser]`). Ścieżki: `Detect_Emf`,
`Png_PassThrough`, `Emf_Placeholder`, `Emf_EmbeddedRasterExtraction`, `Vml_ToSvg`, `Svg_Sanitize`;
param `GraphicCount = 1/10/100` (przepustowość dokumentu).

Uruchomienie (Release):
```
cd D2ApiViewerEditor/D2ViewerEditor.Benchmarks
dotnet run -c Release -- --filter *GraphicConversion*
```
CI: osobny job (jak istniejące benchmarki, report-only — patrz `BENCHMARKS.md`). Interpretacja:
`Mean` (czas/op), `Allocated` (alokacje/op). Progi regresji: placeholder/detekcja w µs; ekstrakcja
rastra skaluje się z rozmiarem (skan sygnatur). Brak natywnych alokacji (pure-managed).

## 10. Ograniczenia

- **EMF/WMF z osadzonym rastrem (PNG/JPEG/DIB) SĄ rasteryzowane do PNG** w przeglądarce
  (od 2026-06-11, przez SkiaSharp — `TryRasterizeMetafileToPng`/`TryExtractEmfDib`/`TryFindDibGeneric`
  w `GraphicConversionService`). Pokrywa najczęstszy realny przypadek (EMF/WMF opakowujący bitmapę).
  Eksport nadal niesie oryginalny metafile (`data-original-src`) → Word renderuje wektor.
- **Czysto wektorowe EMF/WMF** — od 2026-07-05 (ADR-0027) tłumaczone na SVG własnym pure-managed
  parserem rekordów (`MetafileVectorTranslator`, strategia `vector-translate`): pióra/pędzle,
  linie, prostokąty/elipsy, poly*, ścieżki, transformacje świata, StretchDIBits jako `<image>`.
  Poza podzbiorem etapu 1: **tekst (ExtTextOut), clipping, ROP-y, pędzle wzorkowe, rekordy EMF+**
  (pliki dual EMF+ niosą fallback EMF, który tłumaczymy) — takie rekordy są pomijane z licznikiem
  w diagnostyce, a metafile bez żadnych obsługiwanych rekordów degraduje się do przezroczystego
  blanku jak dotąd. SVG to WYŁĄCZNIE podgląd — do DOCX zawsze wraca oryginalny metafile.
  Pełna wierność (EMF+/tekst) = ewentualny sidecar (roadmapa, §4).
- **TIFF**: rasteryzowany do PNG tylko gdy SkiaSharp ma kodek na danej platformie (libtiff często
  nieobecny w obrazie kontenera) — inaczej przezroczysty blank + pass-through oryginału.
- **VML**: tylko bezpieczny podzbiór kształtów (rect/roundrect/oval/line); paths/gradients/cienie/
  textboxy/rotacja → przezroczysty blank. Integracja `ConvertVmlShapeForEditor` w rendererze readera = roadmapa
  (obecnie reader obsługuje VML **v:imagedata**, najczęstszy realny przypadek; kształty wektorowe
  serwowane przez usługę i pokryte testami, hook do readera do dołożenia).
- **SVG do DOCX**: bez PNG-fallback (roadmapa); edytor wstawia rastry, więc zapis rastrów działa.
- Wymiary EMF z `rclFrame`; metafile bez ramki → wymiar z `TargetWidthEmu` (DOCX extent) lub default.
