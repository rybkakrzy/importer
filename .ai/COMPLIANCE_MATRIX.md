# Macierz zgodności — audyt import/edycja/eksport DOCX

> Data: 2026-07-06 (ADR-0028). Zakres: 9 obszarów z promptu „Principal SWE".
> Uzupełnia `AUDIT_WORD_COMPATIBILITY.md` (pełna lista ~60 wad) i `FIDELITY_REPORT.md`.
> Poziomy: **full** · **partial** · **preserved-not-editable** · **unsupported** · **data-loss-risk**.

## Obszary z promptu

| # | Obszar | Stan po tej sesji | Poziom | Dowód |
|---|---|---|---|---|
| 1 | Rozmiar strony / geometria / zoom | Model rozdziela fizyczny rozmiar / marginesy / content; zoom = wizualny `transform:scale` (nie dotyka modelu ani eksportu). Stała px/cm ujednolicona (`units.util`, koniec dryfu linijka 37.795 vs strona 37.8). | full | `units.util.ts`, `units.util.spec.ts`, wysiwyg `pageWidthPx`/`headerBandPx` specy |
| 2 | Stopki per sekcja | Zrobione wcześniej (ADR-0025): stopki/nagłówki per sekcja z własnymi refs; diagnostyka przez `inspect-docx`. | full (round-trip) | `MultiSectionFidelityTests`, `inspect-docx` |
| 3 | Pola numeracji stron w stopce | PAGE/NUMPAGES = dynamiczny placeholder; literalny tekst stopki zachowany; wartości innych pól nie giną (KR-05). | full (PAGE/NUMPAGES) | `FieldValuePreservationTests`, `PageFieldFontTests` |
| 4 | Tabele (gridSpan/vMerge/obramowania/style) | Zrobione wcześniej (ADR-0026): scalanie, style tabel, pozycyjne krawędzie, banding. Niejednolite obramowania NIE normalizowane. | partial→full | `TableStyleFidelityTests`, `TableWriteFidelityTests` |
| 5 | Interlinia / odstępy (auto/exact/atLeast) | Zrobione wcześniej (ADR-0023): before/after/line/lineRule round-trip; `atLeast` przez marker. Limity CSS udokumentowane (PG-09/10). | partial (CSS limit) | `LineSpacing` testy, AUDIT PG-09/10 |
| 6 | Główny toolbar — wybór fontu | Combobox (input+datalist): nie białło, wpisywanie/wyszukiwanie, stan mieszany, commit dopiero po zatwierdzeniu, Escape przywraca. | full | `editor-toolbar.font.spec.ts` (6) |
| 7 | Toolbar kontekstowy — font firmowy | `FontProviderService` wspólny dla obu toolbarów; font firmowy z `--corporate-font-family`. | full | `font-provider.service.spec.ts` (4), `editor-toolbar.font.spec.ts` |
| 8 | Znaczniki podziału strony | Semantyczny niedrukowalny marker (nie grafika); eksport `w:br type=page` (nie drawing). Podgląd tylko w trybie znaków formatowania. | full | `wysiwyg-editor.page-break.spec.ts` (2), writer `CreatePageBreak` |
| 9 | Kształty / kotwice / pola tekstowe | Tekst z `wps:txbx`/`v:textbox` ekstrahowany (KR-06 minimum, wcześniejsza sesja). Pełny model zakotwiczonego kształtu (tabela+logo, wrap, z-order) — niezrobione. | partial / data-loss-risk | AUDIT KR-06, IM-04 |

## Funkcje tabel (prompt §4)

| Funkcja | Import | Render | Edycja | Eksport | Round-trip | Poziom |
|---|---|---|---|---|---|---|
| Łączenie kolumn (`gridSpan`) | ✓ | ✓ | ✓ | ✓ | ✓ | full |
| Łączenie wierszy (`vMerge`) | ✓ | ✓ | ✓ | ✓ | ✓ | full |
| Niejednolite obramowania | ✓ | ✓ | ~ | ✓ | ✓ | partial (edycja per-krawędź ograniczona) |
| Style tabel / banding | ✓ | ✓ | ~ | ✓ | ✓ | partial |
| Zagnieżdżone tabele | ✓ | ✓ | ✓ | ✓ | ✓ | full |
| Tabela w stopce | ✓ | ✓ | ~ | ✓ | ✓ | partial |
| Tabela w kształcie | ~ | ✗ | ✗ | ~ | ✗ | data-loss-risk |
| Przekątne obramowania | ✗ | ✗ | ✗ | ✗ | ✗ | unsupported (TB-01) |

## Kształty / obiekty pływające (prompt §9)

| Element | Poziom | Uwaga |
|---|---|---|
| Text box (tekst) | partial | tekst ekstrahowany (KR-06); pozycja/obramowanie kształtu nie |
| Obraz inline | full | wymiary/rId per część pakietu |
| Obraz zakotwiczony | partial | offset bezwzględny; `relativeFrom/align` gubione (IM-04) |
| Kształt z tabelą + logo | data-loss-risk | brak modelu zagnieżdżonej treści kształtu |
| VML `v:shape`/`v:textbox` | partial | tekst tak, geometria nie |
| Wrap tight/through | unsupported | tylko square/przód/tył (IM-03) |
| Rotacja obrazu | unsupported | IM-01 |

## Pola Word (prompt §3)

| Pole | Poziom | Uwaga |
|---|---|---|
| PAGE / NUMPAGES / SECTIONPAGES | full | dynamiczny placeholder, round-trip jako `w:fld` |
| DATE / TIME | partial | wartość cache zachowana (nie `DateTime.Now`); format instrukcji nie renderowany |
| REF / TOC / MERGEFIELD / SEQ | preserved-not-editable | wartość jako tekst („unlink"), KOD pola nie round-tripuje (KR-05) |

## Elementy nadal nieobsługiwane (skrót)

Przypisy · komentarze · track-changes (ins/del) · zakładki · równania (oMath) ·
wykresy/SmartArt/OLE · kolumny (`w:cols`) · `pgNumType` · `vanish` · ochrona
dokumentu · RTL/bidi · watermark · rotacja obrazu · `tblpPr` · kod pól (TOC/REF).
Pełna lista + priorytety: `AUDIT_WORD_COMPATIBILITY.md` §9.

## Ryzyka utraty danych

Każdy nieobsługiwany element = cichy drop w readerze + regeneracja pakietu w
autosave ⇒ trwała utrata w v2 po 1. autosave (oryginał v1 immutable = jedyna
mitygacja). `inspect-docx` raportuje te elementy (`unsupportedElements`, exit ≠ 0),
`compare-docx` wykrywa je jako `lost`/`unsupported` w round-trip.
