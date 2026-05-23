# AI Agent Workflow

## Cel

Ten plik mówi agentowi AI, jak ma pracować w repozytorium.

## Start każdej sesji

Agent powinien wykonać następujące kroki:

1. Przeczytaj `CLAUDE.md` albo `AGENTS.md`.
2. Przeczytaj `.ai/INDEX.md`.
3. Przeczytaj `.ai/PROJECT_CONTEXT.md`.
4. Przeczytaj `.ai/CURRENT_STATE.md`.
5. Przeczytaj `.ai/TASK_HANDOFF.md`.
6. Otwórz pliki związane z obszarem zadania.
7. Sprawdź fakty w repozytorium.
8. Dopiero potem zaproponuj lub wykonaj zmianę.

## Przed zmianą kodu

Ustal:

- który obszar systemu dotykasz,
- jakie pliki są źródłem prawdy,
- jakie testy chronią ten obszar,
- jaki kontrakt API lub UI może zostać naruszony,
- czy zadanie wymaga decyzji technicznej,
- czy istnieje ryzyko bezpieczeństwa,
- czy zmiana wymaga aktualizacji `.ai/`.

## Podczas pracy

- Rób małe, spójne zmiany.
- Nie poprawiaj niepowiązanego kodu.
- Nie zmieniaj stylu całego modułu.
- Nie dodawaj zależności bez uzasadnienia.
- Jeżeli napotkasz niejasność, zapisz ją w `RISKS_ASSUMPTIONS.md`.
- Jeżeli podejmiesz decyzję techniczną, zapisz ją w `DECISIONS.md`.
- Jeżeli zmieniasz API, sprawdź frontend.
- Jeżeli zmieniasz frontend DTO, sprawdź backend API.
- Jeżeli zmieniasz model danych, sprawdź migracje i wpływ na dane.

## Po zmianie kodu

Zrób możliwą weryfikację:

- build,
- testy,
- lint,
- type-check,
- manual smoke test,
- analiza wpływu na API.

Jeżeli czegoś nie można uruchomić, napisz dlaczego.

## Koniec sesji

Zaktualizuj:

- `.ai/CURRENT_STATE.md`,
- `.ai/TASK_HANDOFF.md`,
- `.ai/CHANGELOG.md`,
- `.ai/DECISIONS.md`, jeżeli była decyzja,
- `.ai/RISKS_ASSUMPTIONS.md`, jeżeli były założenia lub ryzyka,
- `.ai/FEATURES.md`, jeżeli zmieniła się funkcja,
- `.ai/API_CONTRACTS.md`, jeżeli zmieniło się API.

## Format raportu końcowego

```md
## Changed

- ...

## Verified

- `command` — result

## Risks

- ...

## Next

- ...
```

## Reguły anty-chaos

- Nie twórz masowych zmian formatowania.
- Nie zmieniaj lockfile bez zmiany zależności.
- Nie migruj frameworka bez zadania.
- Nie usuwaj testów, żeby build przeszedł.
- Nie wyciszaj błędów kompilacji przez `any`, `!`, `#nullable disable` albo puste `catch`.
- Nie dodawaj TODO zamiast implementacji, chyba że jest to świadoma część planu.
- Nie zmieniaj plików środowiskowych z sekretami.
- Nie wykonuj destructive commands bez zgody.
