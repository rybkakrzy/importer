# AI Project Index

To jest główny indeks pamięci projektu. Agent AI powinien przeczytać ten plik jako pierwszy po `CLAUDE.md` lub `AGENTS.md`.

## Tryb pracy agenta

1. Odczytaj aktualny kontekst z `.ai/`.
2. Zweryfikuj fakty w repozytorium.
3. Ustal minimalny zakres zmiany.
4. Wprowadź małą, spójną zmianę.
5. Zweryfikuj ją buildem, testem, lintem lub analizą.
6. Zaktualizuj `.ai/`, jeżeli zmiana ma znaczenie dla dalszej pracy.

## Domyślna interpretacja projektu

Ten projekt traktuj jako aplikację fullstack:

```text
Angular frontend -> REST API -> ASP.NET Core backend -> database / external services
```

Preferowane zasady:

- backend nie ufa frontendowi,
- frontend nie zawiera krytycznych reguł bezpieczeństwa,
- API ma jawne DTO,
- logika biznesowa nie siedzi w kontrolerach ani komponentach,
- kontrakty API są kompatybilne wstecz, o ile nie ustalono inaczej,
- konfiguracja środowiskowa nie jest sekretem, ale sekrety nigdy nie trafiają do repo.

## Najważniejsze pliki

| Plik | Kiedy czytać |
|---|---|
| `PROJECT_CONTEXT.md` | Zawsze przy starcie nowej sesji |
| `PRODUCT_GOALS.md` | Przy funkcjach, UX, zakresie i priorytetach |
| `TECH_STACK.md` | Przy zależnościach, wersjach, buildzie, strukturze projektu |
| `ARCHITECTURE.md` | Przed zmianami strukturalnymi |
| `FEATURES.md` | Przy pracy nad funkcjami użytkownika |
| `DOMAIN.md` | Przy logice biznesowej i nazewnictwie domenowym |
| `CURRENT_STATE.md` | Zawsze przed kontynuacją pracy |
| `TASK_HANDOFF.md` | Zawsze przed kontynuacją przerwanej pracy |
| `DECISIONS.md` | Przed i po decyzjach technicznych |
| `CODING_STANDARDS.md` | Przed pisaniem lub refaktoryzacją kodu |
| `BACKEND_DOTNET.md` | Przy pracy w backendzie |
| `FRONTEND_ANGULAR.md` | Przy pracy we frontendzie |
| `API_CONTRACTS.md` | Przy zmianach API, DTO, walidacji, HTTP |
| `DATABASE.md` | Przy modelu danych, migracjach i zapytaniach |
| `TESTING_QUALITY.md` | Przy testach, buildzie, CI i jakości |
| `SECURITY.md` | Przy auth, sekretach, danych, walidacji |
| `DEVOPS_DEPLOYMENT.md` | Przy Dockerze, CI/CD, GCP, środowiskach |
| `AI_AGENT_WORKFLOW.md` | Gdy agent nie wie, jak zacząć lub zakończyć pracę |
| `RISKS_ASSUMPTIONS.md` | Gdy pojawia się niepewność lub ryzyko |
| `CHANGELOG.md` | Po każdej istotnej zmianie |
| `GLOSSARY.md` | Przy terminologii technicznej i domenowej |
| `PROMPTS.md` | Gotowe prompty do pracy z agentami |
| `DOCX_CONVERSION.md` | Reference techniczny konwersji DOCX↔HTML: pipeline, model `DocumentContent`, macierz statusów, roadmapa (R-10, computed-style, domknięcia), diagramy |
| `EDITOR_KEYBOARD.md` | Edytor: obsługa klawiatury (ENTER/Backspace/strzałki, granice stron), formatowanie (font-size/family), paste (zwykły/bez formatowania), interlinia (Word→CSS + ograniczenia), linki, obsługa plików (DOCX/PDF/.doc), testy regresji |
| `BENCHMARKS.md` | Benchmarki wydajności/pamięci (BenchmarkDotNet) + fidelity-checks DOCX (R-15..R-21), perf harness frontu, jak uruchomić, progi, CI |
| `FIDELITY_REPORT.md` | Zgodność odwzorowania DOCX→HTML z Wordem (macierz wierności, testy, kolejne kroki) |

## Reguła aktualizacji

Po istotnej zmianie aktualizuj:

```text
.ai/CURRENT_STATE.md
.ai/TASK_HANDOFF.md
.ai/CHANGELOG.md
```

Jeżeli zmiana wpływa na architekturę lub techniczną strategię:

```text
.ai/DECISIONS.md
```

Jeżeli zmiana opiera się na założeniu albo ma ryzyko:

```text
.ai/RISKS_ASSUMPTIONS.md
```
