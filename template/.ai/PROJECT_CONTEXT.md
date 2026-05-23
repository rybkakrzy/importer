# Project Context

## Nazwa projektu

`TODO: wpisz nazwę projektu`

## Typ projektu

Fullstack web application.

Domyślny kontekst technologiczny:

- backend w .NET / ASP.NET Core,
- frontend w Angular,
- komunikacja przez REST API,
- konteneryzacja przez Docker,
- potencjalny deployment do GCP lub innego środowiska kontenerowego,
- praca wspierana przez agentów AI.

## Krótki opis

`TODO: wpisz konkretny opis domeny`

Przykładowy opis do podmiany:

> Projekt jest aplikacją webową do obsługi procesów biznesowych. System składa się z backendu .NET, frontendu Angular oraz warstwy API. Celem projektu jest szybkie dostarczanie funkcji przy zachowaniu czytelnej architektury, testowalności i bezpieczeństwa.

## Główni użytkownicy

Dostosuj do projektu:

- użytkownik końcowy,
- administrator,
- operator systemu,
- właściciel biznesowy,
- integrator API,
- developer / maintainer.

## Najważniejsze scenariusze

Wstępny szkielet:

1. Użytkownik loguje się do aplikacji.
2. Użytkownik przegląda listę zasobów.
3. Użytkownik tworzy lub edytuje zasób.
4. System waliduje dane po stronie backendu.
5. Frontend prezentuje loading, success, validation error i failure state.
6. Administrator zarządza konfiguracją lub dostępem.
7. System zapisuje zdarzenia i błędy w logach.

Dostosuj scenariusze do realnej domeny.

## Cele techniczne

- Utrzymywalna architektura.
- Czytelne granice między frontendem, API, application layer, domeną i infrastrukturą.
- Stabilne kontrakty API.
- Minimalny dług techniczny.
- Możliwość bezpiecznej pracy przez różnych agentów AI.
- Przewidywalny local development.
- Testy dla logiki krytycznej.
- Brak sekretów w repozytorium.

## Główne ograniczenia

- Nie zmieniać publicznych kontraktów API bez świadomej decyzji.
- Nie przepisywać architektury przy okazji małych tasków.
- Nie dodawać zależności bez uzasadnienia.
- Nie mieszać logiki domenowej z warstwą prezentacji.
- Nie przechowywać sekretów w repozytorium.
- Nie robić masowego formatowania niezwiązanego z zadaniem.
- Nie usuwać testów tylko po to, aby build przeszedł.

## Definicja sukcesu

Projekt jest prowadzony dobrze, jeżeli:

- nowe funkcje można dodać bez dużego długu technicznego,
- kod jest czytelny i testowalny,
- backend i frontend mają jasny kontrakt,
- agent AI może kontynuować pracę bez każdorazowego odkrywania projektu od zera,
- zmiany są małe, odwracalne i łatwe do review,
- deployment i lokalne uruchomienie są opisane i powtarzalne.
