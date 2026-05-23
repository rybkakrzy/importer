# Glossary

## Cel

Ten plik opisuje terminy techniczne i domenowe używane w projekcie.

## Terminy domenowe

Do uzupełnienia po analizie realnego projektu.

| Termin | Znaczenie | Uwagi |
|---|---|---|
| User | Użytkownik systemu | Potwierdzić z domeną |
| Role | Rola lub zestaw uprawnień | Potwierdzić z auth |
| Resource | Ogólna nazwa zasobu biznesowego | Zastąpić realną nazwą |
| Status | Stan zasobu lub procesu | Zdefiniować możliwe wartości |
| Audit | Historia operacji | Potwierdzić, czy istnieje |

## Terminy techniczne

| Termin | Znaczenie | Uwagi |
|---|---|---|
| API | Kontrakt komunikacji między frontendem i backendem | Używać zgodnie z `API_CONTRACTS.md` |
| DTO | Obiekt transferu danych | Nie utożsamiać z encją domenową |
| Use case | Operacja aplikacyjna realizująca konkretny scenariusz | Powinna być testowalna |
| Entity | Obiekt z tożsamością | W domenie lub ORM zależnie od architektury |
| Value Object | Obiekt wartości bez tożsamości | Powinien pilnować własnej poprawności |
| Migration | Zmiana schematu bazy danych | Wymaga ostrożności |
| Guard | Mechanizm ochrony routingu w Angularze | Nie zastępuje backend auth |
| Interceptor | Mechanizm przechwytywania HTTP w Angularze | Typowo auth/error handling |
| Service | Klasa usługowa | Powinna mieć konkretną odpowiedzialność |
| Component | Element UI w Angularze | Nie powinien zawierać ciężkiej logiki biznesowej |

## Zasada nazewnictwa

Jeżeli agent wprowadza nowy ważny termin, powinien dodać go do tego pliku lub `DOMAIN.md`.
