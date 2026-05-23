# Reusable Prompts for AI Agents

## Start pracy nad zadaniem

```text
Przeczytaj CLAUDE.md lub AGENTS.md oraz .ai/INDEX.md, PROJECT_CONTEXT.md, CURRENT_STATE.md i TASK_HANDOFF.md.
Następnie przeanalizuj obszar zadania w repozytorium.
Nie zmieniaj kodu, dopóki nie ustalisz istniejącej architektury i źródeł prawdy.
```

## Pierwsze dostosowanie `.ai/` po wrzuceniu do repo

```text
Przejrzyj repozytorium i dostosuj katalog .ai do realnego projektu.
Uzupełnij faktyczny stack, komendy build/test/run, strukturę backendu, strukturę frontendu, bazę danych, deployment i główne funkcje.
Nie zmieniaj kodu aplikacji. Zmieniaj tylko pliki dokumentacyjne .ai, CLAUDE.md i AGENTS.md.
```

## Analiza funkcji

```text
Przeanalizuj tę funkcję od strony backendu, frontendu, API, testów i ryzyk.
Najpierw opisz istniejący stan, potem zaproponuj minimalny plan zmiany.
```

## Implementacja bez chaosu

```text
Wykonaj minimalną zmianę wymaganą przez zadanie.
Nie refaktoryzuj niepowiązanego kodu.
Po zmianie uruchom dostępne testy/build i zaktualizuj .ai/CURRENT_STATE.md, .ai/TASK_HANDOFF.md oraz .ai/CHANGELOG.md.
```

## Review kodu

```text
Zrób code review zmian pod kątem:
- poprawności biznesowej,
- bezpieczeństwa,
- testowalności,
- zgodności z architekturą,
- jakości TypeScript/C#,
- potencjalnych breaking changes w API.

Nie poprawiaj automatycznie wszystkiego. Najpierw wypisz konkretne problemy i ryzyka.
```

## Handoff

```text
Przygotuj przekazanie pracy kolejnemu agentowi.
Zaktualizuj .ai/TASK_HANDOFF.md tak, aby kolejna sesja mogła kontynuować bez ponownego odkrywania kontekstu.
```

## Aktualizacja decyzji technicznej

```text
Jeżeli podczas pracy została podjęta decyzja techniczna, dopisz ją do .ai/DECISIONS.md w formacie ADR.
Uwzględnij kontekst, decyzję, konsekwencje i rozważone alternatywy.
```
