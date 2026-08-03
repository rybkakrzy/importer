# Ściąga: Claude Code CLI

> Dokładny zestaw flag zależy od wersji — pełną, aktualną listę dla swojej instalacji
> zobaczysz przez `claude --help` w terminalu, w którym masz `claude` w PATH.

## Uruchamianie i wznawianie sesji

| Komenda / flaga | Przeznaczenie |
|---|---|
| `claude` | Start interaktywnej sesji (REPL) w bieżącym katalogu. |
| `claude "zadanie"` | Start sesji od razu z podanym promptem. |
| `claude -c` / `--continue` | Wznawia **ostatnią** rozmowę z tego katalogu (bez pytania). |
| `claude -r` / `--resume [id]` | Pokazuje listę poprzednich sesji do wyboru; z podanym ID wznawia konkretną. |
| `claude --session-id <uuid>` | Wymusza konkretny identyfikator sesji (przydatne w automatyzacji). |
| `claude --fork-session` | Przy `--resume`/`--continue` tworzy odgałęzienie zamiast dopisywać do starej sesji. |

## Tryb nieinteraktywny (skrypty / CI)

| Flaga | Przeznaczenie |
|---|---|
| `-p` / `--print "zadanie"` | Jednorazowe wykonanie: odpowiedź na stdout i koniec, bez REPL-a. |
| `--output-format text\|json\|stream-json` | Format wyjścia dla `-p` (JSON przydatny do parsowania w skryptach). |
| `--input-format text\|stream-json` | Format wejścia (stream-json do pipe'owania wiadomości). |
| `--max-turns N` | Limit tur agenta w trybie `-p` (bezpiecznik przed zapętleniem). |
| `--verbose` | Pełne logowanie każdej tury — debug. |

## Model i uprawnienia

| Flaga | Przeznaczenie |
|---|---|
| `--model sonnet\|opus\|haiku\|<pełne-id>` | Wybór modelu na sesję (np. `claude-sonnet-5`). |
| `--permission-mode default\|plan\|acceptEdits\|bypassPermissions` | Tryb uprawnień od startu (np. `plan` = tylko planowanie, bez edycji). |
| `--dangerously-skip-permissions` | Pomija **wszystkie** pytania o zgodę (tzw. YOLO mode) — tylko w zaufanym/odizolowanym środowisku. |
| `--allowedTools "Bash(git *)" "Read"` | Biała lista narzędzi dozwolonych bez pytania. |
| `--disallowedTools "Bash(rm *)"` | Czarna lista narzędzi. |
| `--add-dir <ścieżka>` | Dodaje dodatkowy katalog roboczy poza bieżącym repo. |

## Konfiguracja i dodatki

| Flaga | Przeznaczenie |
|---|---|
| `--append-system-prompt "..."` | Dokleja własny tekst do system promptu (tylko z `-p`). |
| `--settings <plik.json>` | Ładuje dodatkowy plik ustawień. |
| `--mcp-config <plik.json>` | Ładuje serwery MCP z podanego pliku. |
| `--strict-mcp-config` | Używa **tylko** serwerów z `--mcp-config`, ignoruje globalne. |
| `--ide` | Automatycznie łączy się z otwartym IDE przy starcie. |
| `--agents '{...}'` | Definiuje subagentów JSON-em z linii komend. |

## Podkomendy

| Komenda | Przeznaczenie |
|---|---|
| `claude update` | Aktualizacja Claude Code do najnowszej wersji. |
| `claude doctor` | Diagnostyka instalacji (problemy z auto-update, środowiskiem itd.). |
| `claude mcp` | Zarządzanie serwerami MCP (`add`, `list`, `remove`, `add-json`...). |
| `claude config` | Odczyt/zapis ustawień z CLI (`get`, `set`, `list`). |
| `claude setup-token` | Generuje długoterminowy token auth do użycia w CI. |
| `claude install` | Instalacja natywnego buildu. |

## Najważniejsze komendy `/` w trakcie sesji

| Komenda | Przeznaczenie |
|---|---|
| `/resume` | Przełączenie na inną sesję bez wychodzenia. |
| `/rewind` | Cofnięcie rozmowy i/lub zmian w plikach do wcześniejszego punktu (checkpointy). |
| `/clear` | Czyści kontekst — świeży start w tej samej sesji. |
| `/compact` | Kompresuje historię rozmowy (oszczędza kontekst przy długiej pracy). |
| `/model` | Zmiana modelu w locie. |
| `/permissions` | Podgląd i edycja reguł uprawnień (allow/deny). |
| `/config` | Ustawienia (motyw, powiadomienia itd.). |
| `/memory` | Edycja plików pamięci (CLAUDE.md). |
| `/init` | Generuje CLAUDE.md dla projektu. |
| `/agents` | Zarządzanie subagentami. |
| `/hooks` | Konfiguracja hooków (komendy odpalane przy zdarzeniach). |
| `/mcp` | Status i autoryzacja serwerów MCP. |
| `/cost` / `/usage` | Zużycie tokenów / limity planu. |
| `/status` | Wersja, model, konto, katalog roboczy. |
| `/export` | Eksport rozmowy (do pliku/schowka). |
| `/vim` | Tryb vim w polu edycji. |
| `/doctor`, `/bug` | Diagnostyka / zgłoszenie błędu. |

## Najczęstszy praktyczny zestaw

- `claude -c` — wróć do wczorajszej pracy,
- `claude -r` — wybierz starszą sesję z listy,
- `claude -p "..." --output-format json` — skrypty,
- `--permission-mode plan` — bezpieczne planowanie bez dotykania plików.
