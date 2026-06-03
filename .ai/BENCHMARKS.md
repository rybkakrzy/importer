# Benchmarks & fidelity checks

Mierzalne benchmarki wydajności/pamięci oraz fidelity-checks dla pipeline DOCX↔HTML↔DOCX —
do wykrywania regresji **wydajnościowych** i regresji **wierności dokumentów** po zmianach w
parserze, rendererze, edytorze i writerze. Powiązane: `DOCX_CONVERSION.md`, `RISKS_ASSUMPTIONS.md`
(R-15..R-21), `FIDELITY_REPORT.md`.

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
