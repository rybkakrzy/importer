# Observability standard (ELK / GCP)

Standard logowania dla `D2ViewerEditor.Api`. Logi to **structured JSON, jedna linia = jeden wpis**,
pisane na **stdout** (zbierane przez Filebeat/Logstash → Elasticsearch → Kibana; ten sam format niesie
`severity` dla Google Cloud Logging). Bez nowych zależności — własny `GcpJsonConsoleFormatter`.

## Włączanie

`AddGcpStructuredLogging()` (Program.cs) włącza formatter JSON poza Development (albo jawnie
`Logging:UseGcpFormat=true`). Lokalnie zostaje czytelny tekst. Poziomy: `Logging:LogLevel` w appsettings.

## Format wpisu

| Pole | Źródło | Uwagi |
|---|---|---|
| `timestamp` | UTC ISO-8601 (`o`) | |
| `severity` | LogLevel → GCP severity | DEBUG/INFO/WARNING/ERROR/CRITICAL |
| `level` | `LogLevel.ToString()` | ELK-friendly (Information/Warning/…) |
| `message` | sformatowana wiadomość (+ pełny stack trace przy wyjątku) | |
| `category` | kategoria loggera (typ) | |
| `service` | `IHostEnvironment.ApplicationName` | enrichment |
| `environment` | `IHostEnvironment.EnvironmentName` | enrichment |
| `traceId` / `spanId` | `Activity.Current` | trace W3C |
| `eventId` | gdy ≠ 0 | |
| `exceptionType` | FQN wyjątku | przy błędach |
| `correlationId` | scope (middleware) | korelacja w obrębie requestu |
| `requestId` | scope (`HttpContext.TraceIdentifier`) | |
| `httpMethod`,`httpPath`,`statusCode`,`elapsedMs`,`userId` | access-log (argumenty strukturalne) | jeden wpis na request |
| `masterId`,`versionId`,`deliveryId`,`attempt` | scope w workerze wysyłki | business context |

Dodatkowo: wszystkie pary klucz/wartość z **scope'ów** (`BeginScope`) oraz **argumenty szablonu**
wiadomości (`LogInformation("... {Foo}", foo)`) trafiają jako osobne pola.

## Korelacja

- `RequestObservabilityMiddleware` (outermost) czyta nagłówek **`X-Correlation-ID`** (lub generuje:
  `Activity.TraceId`, fallback GUID), zapisuje go w `HttpContext.Items`, zwraca w nagłówku odpowiedzi
  i otwiera scope `{ correlationId, requestId }` — więc **każdy** log requestu (w tym wyjątki z
  `ExceptionHandlingMiddleware`) niesie `correlationId`.
- Błędy zwracane jako ProblemDetails zawierają `correlationId` w `extensions` → klient/support może go
  podać do wyszukania w Kibanie.
- Wysyłka asynchroniczna (`DeliveryAttemptRunner`) loguje w scope `{ deliveryId, masterId, versionId,
  correlationId, attempt }`.

## Poziomy logowania

- `Information` — istotne zdarzenia (zakończony request 2xx/3xx, wysłano dostarczenie).
- `Warning` — obsłużone anomalie (request 4xx, retry/permanent-fail wysyłki).
- `Error` — błędy wymagające diagnostyki (request 5xx, wyjątki, nieudana próba wysyłki).
- `Debug`/`Trace` — wyłącznie szczegóły developerskie.

## Czego NIE logujemy

Zawartości plików/dokumentów, tokenów, haseł, sekretów, pełnych danych osobowych (imię/email),
dużych payloadów request/response. `userId` to **subject/oid** (nieosobowy identyfikator), nie dane
osobowe. `corporateKey` (identyfikator korporacyjny) logowany tylko jako flaga obecności w wysyłce.

## Przykładowe zapytania Kibana (KQL)

- Wszystko z jednego requestu: `correlationId : "abc123"`
- Błędy serwera ostatniej godziny: `level : "Error" and statusCode >= 500`
- Wolne requesty: `elapsedMs > 1000`
- Ścieżka + status: `httpPath : "/api/documentstorage/*" and statusCode : 403`
- Diagnostyka wysyłki dokumentu: `masterId : "<guid>"` lub `deliveryId : "<guid>"`
- Wyjątki konkretnego typu: `exceptionType : "System.InvalidOperationException"`
