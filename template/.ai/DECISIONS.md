# Technical Decisions

Ten plik jest lekkim rejestrem decyzji architektonicznych i technicznych.

## Format wpisu

```md
## ADR-0001: <tytuł>

- Date: YYYY-MM-DD
- Status: Proposed / Accepted / Superseded / Rejected

### Context

<jaki problem rozwiązujemy>

### Decision

<co zdecydowaliśmy>

### Consequences

<pozytywne i negatywne skutki>

### Alternatives considered

<alternatywy>
```

## Decyzje

## ADR-0001: Use `.ai/` as project memory for AI agents

- Date: 2026-05-22
- Status: Accepted

### Context

Projekt ma być rozwijany z pomocą różnych agentów AI. Agenci tracą kontekst między sesjami, dlatego potrzebują trwałego, jawnego źródła wiedzy o projekcie.

Claude korzysta z `CLAUDE.md`, ale projekt nie powinien być uzależniony od jednego narzędzia.

### Decision

Używamy `.ai/` jako głównego katalogu pamięci projektu.

Dodatkowo utrzymujemy:

- `CLAUDE.md` — punkt wejścia dla Claude,
- `AGENTS.md` — neutralny punkt wejścia dla innych agentów.

### Consequences

Pozytywne:

- Agent ma jedno stabilne miejsce z kontekstem.
- Łatwiej kontynuować pracę między sesjami.
- Zmniejsza się ryzyko przypadkowych decyzji technicznych.
- Handoff po pracy staje się standardem.

Negatywne:

- Pliki trzeba aktualizować.
- Nieaktualna pamięć projektu może wprowadzić agenta w błąd.
- Wymaga dyscypliny po stronie ludzi i agentów.

### Alternatives considered

- Trzymanie wszystkiego w `CLAUDE.md`.
- Trzymanie notatek tylko w README.
- Trzymanie wiedzy tylko w issue trackerze.
- Brak trwałej pamięci projektu.

## ADR-0002: Prefer minimal, safe, reviewable AI changes

- Date: 2026-05-22
- Status: Accepted

### Context

Agenci AI mogą wprowadzać zbyt szerokie refaktoryzacje albo mieszać zmianę funkcjonalną z porządkowaniem kodu.

### Decision

Każdy agent powinien preferować małe, spójne, reviewowalne zmiany. Refaktoryzacja powinna być ograniczona do obszaru wymaganego przez zadanie.

### Consequences

Pozytywne:

- Mniejsze ryzyko regresji.
- Łatwiejsze review.
- Łatwiejszy rollback.

Negatywne:

- Niektóre większe porządki trzeba planować jako osobne zadania.
