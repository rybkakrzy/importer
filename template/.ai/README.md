# `.ai` — Project Memory for AI Agents

Ten katalog jest trwałą pamięcią projektu dla agentów AI: Claude, ChatGPT, Cursor, Copilot, Windsurf i innych.

Nie jest to dokumentacja marketingowa. To operacyjny kontekst, który ma pozwolić agentowi wrócić do pracy bez ponownego odkrywania całego repozytorium.

## Domyślny typ projektu

Ten starter jest wstępnie ustawiony pod projekt:

- backend: .NET / ASP.NET Core / C#,
- frontend: Angular / TypeScript,
- API: REST,
- lokalne środowisko: Docker / Docker Compose,
- deployment: kontenery, potencjalnie GCP,
- styl pracy: małe zmiany, dobra jakość, testowalność, brak przypadkowych refaktorów.

Dostosuj pliki do realiów swojego repozytorium po wrzuceniu katalogu do projektu.

## Jak agent ma używać tego katalogu

Agent powinien zaczynać od:

```text
.ai/INDEX.md
```

Następnie powinien przeczytać:

```text
.ai/CURRENT_STATE.md
.ai/TASK_HANDOFF.md
```

Potem powinien otworzyć pliki specyficzne dla zadania, np.:

- backend: `BACKEND_DOTNET.md`,
- frontend: `FRONTEND_ANGULAR.md`,
- API: `API_CONTRACTS.md`,
- baza danych: `DATABASE.md`,
- deployment: `DEVOPS_DEPLOYMENT.md`.

## Co aktualizować po pracy

Po istotnej zmianie agent powinien zaktualizować minimum:

- `CURRENT_STATE.md`,
- `TASK_HANDOFF.md`,
- `CHANGELOG.md`.

Jeżeli podjęto decyzję techniczną:

- `DECISIONS.md`.

Jeżeli pojawiło się ryzyko albo założenie:

- `RISKS_ASSUMPTIONS.md`.

## Pliki startowe dla narzędzi

W katalogu głównym repozytorium są też:

- `CLAUDE.md` — punkt wejścia dla Claude / Claude Code,
- `AGENTS.md` — neutralny punkt wejścia dla innych agentów.
