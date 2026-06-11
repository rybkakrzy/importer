# Editor — klawiatura, formatowanie, paste, interlinia, linki, pliki

Referencja zachowania edytora WYSIWYG (`wysiwyg-editor.ts`) i powiązanego pipeline DOCX.
Dotyczy poprawek z 2026-06-04 (8 błędów). Testy regresji na końcu.

## 1. Obsługa klawiatury

Edytor renderuje każdą stronę jako **osobny `contenteditable`** (multi-page, Wariant A). To
wymusza ręczną obsługę nawigacji/kasowania na granicach stron — przeglądarka nie przenosi
karetki ani nie scala treści między osobnymi `contenteditable`.

| Klawisz | Zachowanie | Kod |
|---|---|---|
| ENTER | Brak custom handlera — domyślne `contenteditable` (nowy blok). Karetka zostaje w nowym akapicie dzięki **block-aware** kotwicy karetki przy repaginacji (patrz niżej). | `handleKeyboard` (brak case Enter) + `_saveGlobalCaret`/`_restoreGlobalCaret` |
| Backspace (zaznaczony obraz) | Usuwa obraz. | `handleKeyboard` |
| Backspace na **początku strony** | Usuwa **manualny page-break** kończący poprzednią stronę (nie wrapper strony) i repaginuje → treść scala się w górę. | `_tryDeletePageBreakBackwards`, `_isCaretAtEditorStart`, `_removeTrailingPageBreak` |
| Backspace (reszta) | Domyślne kasowanie znaku/zaznaczenia (przeglądarka). | — |
| ArrowDown/Up na granicy strony | Przy zwiniętej karetce na skrajnej linii i bez modyfikatorów — przeniesienie do sąsiedniej strony. Shift/Ctrl/Alt i zaznaczenia nietknięte. | `_tryMoveCaretAcrossPages`, `_isCaretOnEdgeLine`, `_placeCaretAtEditorEdge` |
| ArrowLeft/Right | Domyślne (przeglądarka) w obrębie strony. | — |
| Ctrl/Cmd+B/I/U | Bold/Italic/Underline. | `handleKeyboard` → `executeCommand` |
| Ctrl/Cmd+Z / +Y / +Shift+Z | Undo / Redo. | `undo`/`redo` |
| Ctrl/Cmd+Shift+V | Następne wklejenie jako **czysty tekst**. | `plainTextPasteUntil` + `handlePaste` |
| Tab / Shift+Tab | Wcięcie / cofnięcie wcięcia. | `executeCommand('indent'/'outdent')` |

### Kotwica karetki przy repaginacji (root cause Issue 1)

Po wpisaniu znaku edytor planuje repaginację (debounce ~600 ms, `_schedulePaginate` →
`_repaginateNow`), która **przebudowuje DOM stron** (`pageContents.set` → `[innerHTML]`).
Karetkę zachowujemy kotwicą **`{ block, offset }`**:

- `block` = indeks bloku najwyższego poziomu w spłaszczonej sekwencji wszystkich stron
  (`_flattenTopBlocks`, ta sama logika co paginacja — indeksy się zgadzają).
- `offset` = offset tekstowy w obrębie tego bloku.

Czysty **globalny offset tekstowy** był niejednoznaczny na granicy bloków: nowy **pusty**
akapit po ENTER ma 0 znaków, więc restore lądował na końcu poprzedniego akapitu — objaw
„kursor wraca do poprzedniej linii". Block-index rozróżnia pusty nowy akapit od końca
poprzedniego. Pusty blok → karetka na początku bloku (`_placeCaretAtTextOffset`).
Ograniczenie: gdy tabela **przed** karetką dzieli się inaczej w danym przebiegu, indeks może
się przesunąć (rzadkie podczas pisania w akapicie).

## 2. Formatowanie tekstu

| Właściwość | Edytor (HTML) | DOCX (writer) | Uwagi |
|---|---|---|---|
| font-size | inline `style="font-size:Npt"` (span) | `w:sz` (half-points) | `setFontSize` odrzuca `<1`/`>400`/NaN (inaczej `0pt` = niewidoczny tekst — Issue 2). Toolbar też waliduje. |
| font-family | inline `style="font-family:…"` | `w:rFonts` (Ascii+HighAnsi) | Writer **HTML-decode'uje** styl przed parsowaniem: `innerHTML` serializuje nazwy wielowyrazowe jako `&quot;`, a `;` wewnątrz encji ucinał nazwę → Word wracał do domyślnego fontu (Issue 3). Pierwsza rodzina wygrywa, fallback generyczny (`,serif`/`,sans-serif`) odcinany. |
| computed-style | reader rozwiązuje docDefaults → styl akapitu → `w:rStyle` → direct; font fallback generyczny wg nazwy (serif/mono/sans). | — | patrz `DOCX_CONVERSION.md`. |

Reguła: **firmowy/domyślny font nie nadpisuje explicit font-family** — explicit inline wygrywa
w CSS (specyficzność) i jest zapisywany w `w:rFonts`. Fallback fontu (gdy krój niedostępny w
przeglądarce) **nie** jest zapisywany jako font dokumentu — w DOCX zostaje oryginalna nazwa.

## 3. Paste handling

- **Zwykłe wklejenie (Ctrl+V):** `handlePaste` — sanitizuje `text/html` (usuwa script/style/
  komentarze/`on*`), wstawia przez `insertHtml`. Brak HTML → `text/plain` znormalizowany.
- **Wklej bez formatowania:** Ctrl+Shift+V (okno czasowe `plainTextPasteUntil`) **lub** menu
  „Edytuj → Wklej bez formatowania" / menu kontekstowe → `pasteWithoutFormatting()` →
  `navigator.clipboard.readText()` (tylko `text/plain`) → `editor.insertText`.
- **Root cause Issue 4:** klik w pozycję menu **zabierał fokus** edytorowi, więc
  `execCommand('insertText')` nie miał karetki i nic nie wstawiał. `insertText` **odtwarza
  zapisaną selekcję i fokus** przed wstawieniem (jak `setFontSize`). Brak dostępu do schowka →
  ciche niepowodzenie (zostaje skrót Ctrl+Shift+V).
- Clipboard API `readText()` wymaga bezpiecznego kontekstu (HTTPS/localhost) i zgody.

## 4. Interlinia (line spacing) — mapowanie Word → CSS

Reader (`ConvertParagraphPropertiesToCss`) mapuje `w:spacing`:

| OOXML | CSS | Wartość |
|---|---|---|
| `w:line` + `lineRule="auto"` (lub brak) | `line-height:<n>` (bezjednostkowe) | `line/240` → 240=pojedyncze (1.0), 360=1.5, 480=podwójne (2.0) |
| `w:line` + `lineRule="exact"`/`"atLeast"` | `line-height:<n>pt` | `line/20` (twips→pt) |
| `w:before` / `w:beforeLines` | `margin-top:<n>pt` | twips→pt; `beforeAutoSpacing` pomijane |
| `w:after` / `w:afterLines` | `margin-bottom:<n>pt` | twips→pt; `afterAutoSpacing` pomijane |
| `w:contextualSpacing` | `--w-contextual-spacing:1` | round-trip do writera |

**Mapowanie jest poprawne** (testy `LineSpacingMappingTests`). Spacing before/after **nie jest
liczony podwójnie** (osobno margin-top/bottom; brak nakładania z innym źródłem CSS — inline
wygrywa).

### Ograniczenia (HTML/CSS vs Word — nieusuwalne mapowaniem)
1. **„Single" ≠ `line-height:1`.** Word w trybie single (auto, 240) dolicza *line-gap* metryki
   fontu (~1.15 dla Calibri); CSS `line-height:1` to dokładnie rozmiar fontu → przeglądarka
   renderuje ciaśniej. Mapowanie jest wierne wartości OOXML; różnica wynika z metryk fontu.
2. **Kolaps marginesów.** Sąsiednie `margin-bottom` + `margin-top` w CSS **kolabują** (większy
   wygrywa), a Word **sumuje** before+after między akapitami → odstęp w przeglądarce bywa
   mniejszy. Zmiana na padding ryzykowałaby tła/obramowania/round-trip — świadomie zostawione.

`atLeast` jest mapowane jak `exact` (CSS nie ma „co najmniej" dla line-height) — akceptowalne
przybliżenie.

## 5. Linki (hyperlink)

- **UI:** toolbar „Wstaw link" → dialog (`showLinkDialog`, `linkUrl`/`linkText`) →
  `confirmInsertLink` → `insertLink(url, text)`.
- **Root cause Issue 6:** pole URL dialogu zabierało fokus, selekcja edytora kolabowała →
  `createLink` nie miał na czym działać. `insertLink` **odtwarza zapisaną selekcję** (zapisaną
  na mouseup/keyup edytora), **normalizuje URL** (`normalizeLinkUrl`: schematy http/mailto/tel/
  `#`/`/` zostają; goła domena → `https://`) i **escapuje** etykietę. Zaznaczenie → `createLink`;
  zwinięta karetka → wstawia `<a href target=_blank rel=noopener>`.
- **Model/HTML:** `<a href>` w treści; reader czyta hyperlinki (`ConvertHyperlinkToHtml`).
- **DOCX:** writer `ConvertAnchorElement` → `w:hyperlink` + relacja w `document.xml.rels`.
- jsdom nie ma `execCommand`/`createLink`, więc wstawienie w DOM jest weryfikowane manualnie;
  testy automatyczne pokrywają normalizację URL + escaping + no-op dla pustego URL.

## 6. Obsługa plików

| Format | Status | Ścieżka |
|---|---|---|
| `.docx` | Obsługiwany | upload → `OpenDocument` (`/open`) → parser OOXML |
| `.pdf` | Obsługiwany | upload → podgląd PDF |
| `.doc` (binarny, OLE/CFBF) | **Konwertowany do .docx** (od 2026-06-11) | `/open` → `DocumentInputNormalizer` → `LegacyDocBinaryConverter` (FIB + piece table → tekst+akapity → DOCX) |
| `.doc` mislabeled (faktycznie DOCX) | Obsługiwany | normalizer wykrywa ZIP → pass-through |
| inne | Odrzucane | „Obsługiwane są pliki DOCX i PDF." |

**Decyzja `.doc` (Issue 7, zaktualizowana 2026-06-11):** `.doc` to **stary binarny format** (OLE/CFBF),
nie OOXML. Zamiast odrzucać, **czysto zarządzany `LegacyDocBinaryConverter`** parsuje FIB
(wIdent 0xA5EC, `fcClx`/`lcbClx` @0x01A2/0x01A6, `fWhichTblStm`) i piece table (CLX/PlcPcd; PCD.fc
compressed=CP1252/Latin1 vs 16-bit Unicode), wyciąga **tekst + podział akapitów** i buduje DOCX
przez OpenXML — bez LibreOffice/GDI (Linux/GCP-safe). **Świadome ograniczenie:** odzyskiwany jest
tekst, NIE bogate formatowanie/tabele/obrazy (pełna wierność wymagałaby dużego parsera MS-DOC lub
konwertera zewnętrznego). Gdy struktura jest niespójna → **fallback** na kontrolowane odrzucenie
(`UnsupportedLegacyDoc` → 400 z instrukcją konwersji), nigdy śmieci. Walidacja:
- Backend `DocumentController.OpenDocument` — rozszerzenie + detekcja po zawartości; binarny `.doc`
  → konwersja, a przy niepowodzeniu 400 z komunikatem.
- Frontend `document-editor.openDocument` — `accept=.docx,.doc,.pdf`, ładowanie przez `/open`.
- `dashboard.openFile` (utwórz nowy z dysku, bez normalizacji) — kieruje `.doc` do „Plik → Otwórz".

**Roadmapa:** pełna wierność (formatowanie/tabele/obrazy) przez LibreOffice headless / sidecar.

## 7. Testy regresji

Uruchomienie:
- Backend: `dotnet test D2ViewerEditor.Infrastructure.UnitTests` (i `.Api.UnitTests`).
- GUI: `cd D2GuiViewerEditor && npx ng test --watch=false`.

| Obszar | Test | Pokrycie |
|---|---|---|
| ENTER / karetka | `wysiwyg-editor.enter-caret.spec.ts` | block-aware save/restore; pusty akapit ≠ poprzednia linia |
| font-size | `wysiwyg-editor.toolbar-actions.spec.ts` | odrzucenie 0/NaN/zakresu; akceptacja poprawnego |
| font-family | `FontFamilyWriteTests.cs` | `&quot;`/wielowyraz/single-quote+fallback/unquoted → `w:rFonts` bez cudzysłowów |
| font-family (front) | `wysiwyg-editor.toolbar-actions.spec.ts` | `setFontFamily` odtwarza zapisaną selekcję **przed** `focus()` (wybór z `<select>` gubi selekcję → bez tego ZWS-span lądował na początku dokumentu i nowy tekst dziedziczył domyślną czcionkę) |
| paste plain | `wysiwyg-editor.toolbar-actions.spec.ts` | `insertText` odtwarza selekcję po utracie fokusu |
| line spacing | `LineSpacingMappingTests.cs` | single/1.5/double/exact/atLeast/before/after |
| linki | `wysiwyg-editor.toolbar-actions.spec.ts` | normalizacja URL + escaping + no-op pustego |
| `.doc` | `DocumentControllerTests.cs` | `.doc` → 400 z instrukcją; `.pdf`/inne → 400 |
| Backspace page-break | `wysiwyg-editor.caret-field.spec.ts` | usunięcie page-break na początku strony; brak na środku linii; brak na 1. stronie |
| nawigacja strzałkami | `wysiwyg-editor.caret-field.spec.ts` | ArrowDown/Up między stronami; Shift/range bez ruchu |

### Scenariusze manualne (jsdom nie pokrywa contenteditable/execCommand/layout)
1. ENTER w akapicie/pustym akapicie/tabeli/nagłówku/stopce/przy page-breaku — karetka zostaje w
   nowej linii również po ~600 ms (repaginacja).
2. Font-size 0/puste/litery → tekst nie znika; poprawna wartość zmienia rozmiar zaznaczenia.
3. Font-family „Times New Roman" → zapis DOCX ma `w:rFonts ascii="Times New Roman"`, po
   ponownym otwarciu font się trzyma.
4. „Wklej bez formatowania" (menu + Ctrl+Shift+V) → wkleja czysty tekst w miejscu karetki.
5. Link na zaznaczeniu / pustej karetce → `<a>` + relacja w DOCX; po round-tripie działa.
6. Backspace na początku 2. strony usuwa twardy podział; strzałki góra/dół przechodzą między
   stronami bez losowych skoków.
