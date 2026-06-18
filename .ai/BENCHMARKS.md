# Benchmarks & fidelity checks

Mierzalne benchmarki wydajności/pamięci oraz fidelity-checks dla pipeline DOCX↔HTML↔DOCX —
do wykrywania regresji **wydajnościowych** i regresji **wierności dokumentów** po zmianach w
parserze, rendererze, edytorze i writerze. Powiązane: `DOCX_CONVERSION.md`, `RISKS_ASSUMPTIONS.md`
(R-15..R-21), `FIDELITY_REPORT.md`.

## 0. Po co to w ogóle robimy

Cały produkt stoi na jednej obietnicy: dokument z Worda wygląda i wraca z Qutas tak samo, a edycja
jest płynna. Te benchmarki to wczesny alarm — łapią moment, w którym zmiana w konwerterze cicho
psuje wierność albo spowalnia pipeline, **zanim** zobaczy to użytkownik (albo Word komunikatem
„dokument uszkodzony"). Bez nich regresje wychodzą dopiero na realnych plikach klienta, gdzie są
najdroższe do zdiagnozowania.

### Import `docx→html` (reader)
**Po co:** to pierwszy kontakt z plikiem klienta — jeśli reader zgubi rozmiar czcionki, podział
stron czy styl, cała reszta edytora pracuje już na zepsutym modelu. **Co chcemy zobaczyć:** że
struktura wchodzi 1:1 (style, tabele, page-breaki, wymiary) i że czas rośnie liniowo z rozmiarem
dokumentu, bez nagłych skoków. **Co z tym robimy:** każdy skok czasu/alokacji albo spadek wierności
wskazuje konkretny commit do cofnięcia lub poprawy; ustalamy też budżet „ile wolno trwać importowi".

### Eksport `html→docx` (writer)
**Po co:** to ścieżka, która realnie psuła pliki — zły porządek elementów czy goły SVG blip dają
DOCX, którego Word nie otworzy bez ostrzeżenia. **Co chcemy zobaczyć:** że wynik jest poprawny
schematowo (OpenXmlValidator = 0 błędów), a pass-through nie gubi stylów/theme/tabel względem
oryginału. **Co z tym robimy:** twarde inwarianty (table count, page-breaki) failują CI, więc
korupcja nigdy nie dojedzie do produkcji; resztę (numbering, rozmiar) trzymamy jako WARN do domknięcia.

### Konwersja obrazków (EMF/WMF/VML/SVG)
**Po co:** legacy grafiki Worda nie renderują się w przeglądarce i potrafią wywrócić zarówno podgląd,
jak i zapis — a robimy to pure-managed, bez natywnych zależności (Linux/GCP-safe). **Co chcemy
zobaczyć:** że konwersja jest tania (µs–ms, bez alokacji natywnych), bezpieczna (sanityzacja SVG,
blok XXE) i że oryginał przeżywa round-trip bez uszkodzenia pliku. **Co z tym robimy:** pilnujemy, by
przepustowość nie siadła przy dokumentach z dziesiątkami grafik, i decydujemy, kiedy opłaca się
dołożyć cache albo sidecar rasteryzujący EMF.

## 1. Co jest gdzie

| Element | Lokalizacja | Narzędzie |
|---|---|---|
| Backend benchmarki (perf+pamięć) | `D2ViewerEditor.Benchmarks/` | BenchmarkDotNet 0.14 |
| Fidelity report (DOCX, report-only) | `D2ViewerEditor.Benchmarks/` (`--fidelity`) | własny analizator pakietu |
| Fidelity gate (CI, szybki) | `D2ViewerEditor.Infrastructure.UnitTests/Services/DocxFidelityRegressionTests.cs` | NUnit |
| Frontend perf harness (report-only) | `D2GuiViewerEditor/.../wysiwyg-editor.perf.spec.ts` | Vitest + Performance API |
| Sample data (syntetyczne) | `D2ViewerEditor.Benchmarks/BenchmarkAssets.cs` | generowane w pamięci |

Sample DOCX są **generowane w pamięci** (brak commitowanych, potencjalnie wrażliwych plików).
Opcjonalny realny plik regresyjny (`orginał_GOOD.docx`) jest używany **tylko gdy istnieje** —
ścieżka z env `D2_BENCH_ASSETS` (domyślnie katalog repo `importer`).

## 2. Backend — benchmarki BenchmarkDotNet

Wymaga **Release**. Wyniki (md/html/csv) → `BenchmarkDotNet.Artifacts/`.

```bash
cd D2ApiViewerEditor
# wszystkie (interaktywne menu):
dotnet run -c Release --project D2ViewerEditor.Benchmarks
# konkretna grupa:
dotnet run -c Release --project D2ViewerEditor.Benchmarks -- --filter "*DocxRoundTripBenchmarks*"
# szybki smoke (1 iteracja, niewiarygodne liczby, tylko że działa):
dotnet run -c Release --project D2ViewerEditor.Benchmarks -- --filter "*Import*" --job Dry
```

| Klasa | Mierzy | Dane wejściowe | Metryki |
|---|---|---|---|
| `DocxImportBenchmarks` | reader `DocxToHtmlConverter.Convert` | simple / tables / images / multi-style / (regression) | czas, alokacje (`[MemoryDiagnoser]`) |
| `DocxExportBenchmarks` | writer `Convert` + `ConvertPreservingPackage` (R-16) | simple / tables / (regression) | czas, alokacje, rozmiar bajtów |
| `DocxRoundTripBenchmarks` | pełny reader→writer | tables / header-footer / page-breaks / regression(pass-through) | czas, alokacje |
| `TableBenchmarks` (`[ShortRunJob]`) | parse/serialize dużej tabeli | `[Params] Rows = 100, 400` | czas, alokacje |

Wszystkie mają `[MemoryDiagnoser]` (Allocated/op, Gen0). „Baseline" w każdej grupie to wariant
`Simple`/`Tables` (atrybut `Baseline=true`) — BenchmarkDotNet podaje kolumnę `Ratio` względem niego.

Lekkie: Import/Export/RoundTrip (sekundy–minuty). Cięższe: `TableBenchmarks` Rows=400 (stąd ShortRun).

## 3. Backend — fidelity report (DOCX, report-only)

Szybki (sekundy), nie wymaga Release. Przepuszcza próbki przez round-trip i ocenia inwarianty.
Oryginał jest swoim własnym baseline (metryki original-vs-result — bez przechowywanego baseline).

```bash
cd D2ApiViewerEditor
dotnet run -c Release --project D2ViewerEditor.Benchmarks -- --fidelity
dotnet run -c Release --project D2ViewerEditor.Benchmarks -- --fidelity --fail-on-regression  # exit 1 na FAIL
```

Raport markdown → konsola + `bin/.../fidelity-reports/fidelity-*.md`. Per próbka tabela PASS/WARN/FAIL:

| Check | Inwariant | Gate |
|---|---|---|
| Styles preserved | R-16: liczba stylów nie spada >50% (pass-through) | FAIL (pass-through) / info (self-contained) |
| Table styles | R-16: style tabel zachowane | WARN |
| Theme preserved | R-16: theme nie znika (pass-through) | FAIL (pass-through) / info |
| Theme minor font | R-16: font theme nie podmieniony (Cambria↛Calibri) | WARN |
| Numbering preserved | R-19: `numbering.xml` (obecnie NIE przenoszony) | WARN |
| Table count | R-17: jedna tabela nie staje się dwiema | **FAIL** (uniwersalne) |
| Page breaks | R-15: brak utraty/duplikacji | **FAIL** na utracie / WARN na dodaniu |
| Row heights | R-18: brak sztucznych `trHeight` | WARN |
| File size | informacyjne (% różnicy) | info |

Status interpretacji: dla **self-contained** writera (bez oryginału) utrata styles/theme jest
*projektowa* (regeneracja) → nie FAIL. Dla **pass-through** utrata = regresja (FAIL/WARN).

## 4. Backend — fidelity gate (CI, szybki, NUnit)

`DocxFidelityRegressionTests` (4 testy, <1 s) jako twardy gate w zwykłym `dotnet test`:
- `PassThrough_PreservesStylesAndTheme_R16` — styles nie kolapsują, theme=Cambria.
- `RoundTrip_KeepsManualPageBreaks_R15` — page breaki bez utraty/duplikacji.
- `RoundTrip_DoesNotMultiplyTables_R17` — liczba tabel zachowana.
- `RoundTrip_DoesNotFabricateRowHeights_R18` — brak sztucznych `trHeight`.

Uzupełniają istniejące per-feature: `PassThroughPackageTests`, `PageBreakRoundTripTests`,
`RoundTripLayoutFidelityTests`, `ImageAltTextRoundTripTests`, golden snapshoty.

## 5. Frontend — perf harness (report-only, Vitest)

```bash
cd D2GuiViewerEditor && npx ng test --watch=false   # loguje [perf] ... ms
```

`wysiwyg-editor.perf.spec.ts` mierzy operacje CPU-bound edytora (jsdom) z lenient ceiling 4000 ms:
- `getContent` na 800 akapitach / 500-wierszowej tabeli,
- scalanie 20 fragmentów split-table (R-17),
- `setContent` (split na strony) dla dużego HTML z page-breakami.

Przykładowe wyniki (lokalnie, jsdom): getContent 800 akapitów ~140 ms, tabela 500 wierszy ~100 ms,
merge 20 fragmentów ~175 ms, setContent 50 stron ~3 ms.

**Ograniczenie:** jsdom nie liczy layoutu → czas `_repaginateNow` (paginacja wg wysokości),
realnego renderu i layout-shiftów **nie da się** zmierzyć tutaj. To wymaga harnessu przeglądarkowego
(Playwright) — niewprowadzanego teraz (projekt go nie ma; rule: bez ciężkiej infrastruktury bez powodu).

## 6. Baseline i quality gates (report-only → fail-on-regression)

- **Perf baseline:** BenchmarkDotNet zapisuje wyniki w `BenchmarkDotNet.Artifacts/` (md/csv).
  Porównanie do poprzedniego biegu: zachować artefakty jako baseline i porównać `Mean`/`Allocated`
  ręcznie albo `dotnet run -- --filter ... --statisticalTest 5%` / narzędzia diff BenchmarkDotNet.
  Progi do rozważenia (po zebraniu historii, na razie informacyjne): czas import/export +20%,
  alokacje +25%.
- **Fidelity gate:** twarde inwarianty (table count, page breaks) failują przez NUnit gate i
  `--fidelity --fail-on-regression`. Reszta (styles/theme/numbering/file-size) = report-only/WARN
  do czasu ustalenia twardych progów.

## 7. CI/CD

- **Zwykły pipeline (PR):** `dotnet test D2ViewerEditor.sln` (zawiera fidelity gate, szybki) +
  `ng test --watch=false` (zawiera perf harness report-only). **Bez** ciężkich benchmarków.
- **Okresowo / nightly:** `dotnet run -c Release --project D2ViewerEditor.Benchmarks -- --filter "*"`
  + `--fidelity --fail-on-regression`; archiwizować `BenchmarkDotNet.Artifacts/` i `fidelity-reports/`.
- Lokalnie: jak w sekcjach 2–5.

## 8. Ograniczenia / dalsza praca

- Brak frontendowego pomiaru realnego renderu/paginacji (potrzebny Playwright).
- Brak API-level benchmarków (endpointy) — pipeline mierzony na poziomie konwerterów (gdzie jest
  koszt); benchmark endpointów wymagałby `WebApplicationFactory` + DB/GCS mocków (kandydat na później).
- Perf progi są informacyjne do czasu zebrania historii baseline.
- Fidelity numbering (R-19) jest WARN — pass-through nie przenosi jeszcze `numbering.xml`.

## 9. Notatka z wyników — przebieg 2026-06-08

> Profesjonalna notatka gotowa do dokumentacji. Dane z realnego uruchomienia raportu fidelity
> (`--fidelity`, Release) + Dry-benchmarków (cold-start) + pokrycia testami konwersji.

### 9.1 Cel i zakres
Mierzymy **wierność round-tripu DOCX↔HTML↔DOCX** (zachowanie struktury dokumentu) oraz **koszt
czasowy** trzech ścieżek: import `docx→html` (reader), eksport `html→docx` (writer) i konwersja
grafik (EMF/WMF/VML/SVG). Próbki: syntetyczne (simple/tables/images/page-breaks/multi-style) +
realny plik regresyjny `orginał_GOOD.docx` (19 partów, 6 tabel, ~46 KB HTML).

### 9.2 Wierność konwersji — wynik (overall **WARN**, jedyna przyczyna: R-19 numbering)

| Próbka | Status | Style | Tabele | Theme | Page breaks | Rozmiar pliku |
|---|---|---|---|---|---|---|
| simple | PASS | 1→14 (regeneracja) | — | regen (self-contained) | 0 | 2318→3832 B (+65%) |
| tables | PASS | 1→1 | 6→6 | present | 0 | 2609→4895 B (+88%) |
| images | PASS | 1→14 | — | regen | 0 | 7631→9512 B (+25%); parts 31→32 |
| page-breaks | PASS | 1→14 | — | regen | **3 zachowane** | 2266→3774 B (+67%) |
| multi-style (pass-through) | PASS | **161→161** | **100→100** | present | 0 | 3445→4787 B (+39%) |
| **regression `orginał_GOOD.docx`** | **WARN** | **164→164** | **100→100** | present (Cambria) | **1 zachowany** | 39893→25142 B (**−37%**) |

Jedyny WARN: `numbering.xml` nie jest przenoszony przy pass-through (**R-19**). Wszystkie twarde
inwarianty PASS: **table count** (6→6), **page breaks** (bez utraty/duplikacji), **row heights**
(30→30, brak fabrykowanych `trHeight`), **styles/theme** przy pass-through nietknięte.

**Interpretacja:**
- *Self-contained writer* (próbki bez oryginału): style/theme są **regenerowane** projektowo
  (1→14 = dodanie bazowego zestawu) — to nie regresja. Rozrost rozmiaru (+25..+88%) wynika z
  rozpakowanego, nieskompresowanego zapisu małych próbek — przy realnym pliku jest **−37%**
  (writer pisze zwięźlej niż oryginał Worda).
- *Pass-through* (multi-style + regression): style (164→164), style tabel (100→100) i theme
  przechodzą **1:1** — to potwierdza, że `ConvertPreservingPackage` zachowuje pakiet źródłowy.

### 9.3 Koszt czasowy — import `docx→html` (BenchmarkDotNet, Dry/cold-start, 1 iteracja)

> ⚠ Dry = pojedynczy przebieg z JIT cold-start → **górne oszacowanie, rząd wielkości**. Steady-state
> (pełny `dotnet run -c Release`) będzie istotnie niższy. Liczby orientacyjne do baseline.

| Próbka | Czas/op (cold) |
|---|---|
| Simple | ~80 ms |
| Tables | ~98 ms |
| Images | ~85 ms |
| MultiStyle | ~93 ms |
| **Regression (realny plik 19-part)** | **~293 ms** |

### 9.4 Koszt czasowy — front (Vitest/jsdom, realne pomiary [perf])

| Operacja | Czas |
|---|---|
| `getContent` — 800 akapitów | ~140 ms |
| `getContent` — tabela 500 wierszy | ~100 ms |
| `getContent` — scalanie 20 fragmentów split-table | ~214 ms |
| `setContent` — split 50 stron z page-breakami | ~5 ms |

⚠ jsdom **nie liczy layoutu** → `_repaginateNow` (paginacja wg wysokości) **niemierzone**.

### 9.5 Konwersja obrazków
Pure-managed (`GraphicConversionService`): detekcja/placeholder w zakresie **µs–ms**, **zero alokacji
natywnych**. Pokrycie: 22 testy (13 service + 7 security + 2 integration). **Round-trip EMF
zweryfikowany OpenXmlValidatorem = 0 błędów schematu** (po naprawie z 2026-06-08: oryginalny EMF
zapisywany zamiast gołego SVG blip; brak „uszkodzonego DOCX").

### 9.6 Pokrycie testami per kierunek konwersji (Infrastructure, 144 total)
- **`docx→html` (reader): 40** — Fidelity 9, HeaderFooter 5, SectionReference 8, DefaultStyleFont 6,
  FontSizeImport 4, LineSpacing 6, PageFieldFont 2.
- **`html→docx` (writer): 8** — FontFamilyWrite 3, PassThroughPackage 5.
- **round-trip (oba kierunki): 32** — PageBreak 8, RoundTripLayoutFidelity 4, HeaderFooter 6,
  PageSize 3, ImageAltText 3, ImageBorderCrop 2, ImageFloating 2, FidelityRegression 4.
- **obrazy: 22** — GraphicConversion (service/security/integration).

### 9.7 Wnioski i rekomendacje
1. **Wierność jest wysoka** — wszystkie twarde inwarianty PASS na realnym pliku; jedyny dług to
   **R-19 (numbering pass-through)** — rekomendacja: przenosić `numbering.xml` w `ConvertPreservingPackage`.
2. **Koszt importu** rośnie ~liniowo z liczbą partów/tabel; realny plik ~293 ms cold — akceptowalne,
   ale warto zebrać **steady-state baseline** (pełny Release) i ustawić progi (import +20%, alloc +25%).
3. **Luka pomiarowa:** realny render/paginacja na froncie — wymaga **Playwright** (rekomendacja stała).
4. **Rozmiar pliku** self-contained rośnie na małych próbkach, ale maleje na realnym (−37%) —
   nie jest problemem produkcyjnym.

### 9.8 Ryzyka/ograniczenia pomiarów
Dry-benchmark to cold-start (zawyżony); raport fidelity to original-vs-result (bez przechowywanego
baseline historycznego); front bez layoutu (jsdom). Pełny wiarygodny przebieg: `dotnet run -c Release
--project D2ViewerEditor.Benchmarks -- --filter "*"` + `--fidelity --fail-on-regression`.
