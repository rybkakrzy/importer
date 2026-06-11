# Konwersja grafik dokumentowych (VML / EMF / WMF → web)

Usługa konwertująca legacy grafiki Worda do formatu wyświetlanego w przeglądarce, z bezpiecznym
fallbackiem i zachowaniem wierności w DOCX przez pass-through. Wprowadzone 2026-06-04.

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
| PNG/JPEG/GIF/BMP | passthrough | ten sam raster (data URL) | Lossless |
| SVG | sanitizacja | SVG bezpieczny | Lossless/Fallback |
| **EMF/WMF** z osadzonym PNG/JPEG | ekstrakcja PNG/JPEG z bajtów | osadzony raster | Lossy |
| **EMF/WMF** z osadzonym DIB (StretchDIBits itp.) | DIB→BMP→**SkiaSharp**→PNG (od 2026-06-11) | PNG | Lossy |
| **EMF/WMF** bez rastra (czysty wektor) | **placeholder SVG** (wymiary z headera) + **pass-through oryginału** | placeholder | Fallback |
| **VML** rect/oval/line/roundrect | mapowanie bezpiecznego podzbioru | SVG (fill/stroke) | Lossy |
| **VML** v:imagedata | rozwiązanie partu → ścieżka EMF/WMF/raster | jw. | jw. |
| VML inne / nieznane | placeholder + diagnostyka | placeholder | Unsupported |

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

## 7. Cache (projekt)

Konwersja jest deterministyczna i czysta → kandydat do cache po **hashu bajtów media partu / VML XML**
(klucz = SHA-256(bytes) + wymiary docelowe). MVP nie cache'uje (konwersje są tanie: rząd µs–ms);
cache dodać przy batchach 100+ grafik. Invalidacja naturalna (hash treści).

## 8. Testy

Uruchom: `dotnet test D2ViewerEditor.Infrastructure.UnitTests --filter GraphicConversion` (21).

| Plik | Zakres |
|---|---|
| `GraphicConversionServiceTests` | detekcja, PNG passthrough, EMF/WMF wymiary+placeholder, ekstrakcja osadzonego PNG, VML rect/oval/line + imagedata→null + unsupported→null, limity |
| `GraphicConversionSecurityTests` | sanitizacja SVG (script/on*/external/javascript), XXE/DTD blok, malformed SVG/EMF → fallback bez wyjątku, clamp wymiarów |
| `GraphicConversionIntegrationTests` | DOCX z EMF a:blip → reader emituje `data:image/svg+xml` + `data-legacy-graphic="placeholder"`, **nigdy** `data:image/x-emf` |

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
- **Czysto wektorowe EMF/WMF (bez osadzonego rastra)** wciąż nie są rasteryzowane (placeholder) —
  pełny podgląd dopiero w Word (pass-through). Pełny interpreter wektora EMF = roadmapa (sidecar).
- **VML**: tylko bezpieczny podzbiór kształtów (rect/roundrect/oval/line); paths/gradients/cienie/
  textboxy/rotacja → placeholder. Integracja `ConvertVmlShapeForEditor` w rendererze readera = roadmapa
  (obecnie reader obsługuje VML **v:imagedata**, najczęstszy realny przypadek; kształty wektorowe
  serwowane przez usługę i pokryte testami, hook do readera do dołożenia).
- **SVG do DOCX**: bez PNG-fallback (roadmapa); edytor wstawia rastry, więc zapis rastrów działa.
- Wymiary EMF z `rclFrame`; metafile bez ramki → wymiar z `TargetWidthEmu` (DOCX extent) lub default.
