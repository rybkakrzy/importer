6. Co jeśli inny użytkownik ma link i też chce otworzyć plik?

Tu musisz podjąć świadomą decyzję produktowo-architektoniczną.

Masz trzy główne opcje.

Opcja A — jeden aktywny edytujący

To jest najprostsze i najbezpieczniejsze podejście.

Jeżeli użytkownik A otworzył dokument, backend zakłada lock:

documentId -> locked by editSessionId/userId until leaseExpiresAt

Gdy użytkownik B otwiera ten sam link, backend odpowiada np.:

423 Locked

Albo zwraca dokument tylko do odczytu:

{
  "status": "Locked",
  "mode": "ReadOnly",
  "lockedBy": "jan.kowalski@example.com",
  "lockExpiresAt": "2026-05-19T12:15:00Z"
}

Możesz dać opcje:

- Otwórz tylko do odczytu
- Spróbuj ponownie później
- Przejmij edycję, jeśli poprzednia sesja wygasła

To podejście polecam, jeśli Twój edytor nie ma prawdziwej współbieżnej edycji opartej o operacje, konfliktowanie, merge, OT albo CRDT.

Opcja B — wielu użytkowników może otworzyć, ale tylko jeden zapisuje

To jest wariant podobny, tylko UX jest trochę łagodniejszy.

Użytkownik B może zobaczyć dokument, ale nie może edytować, dopóki A trzyma locka.

To ma sens, jeżeli dokumenty mogą być konsultowane, ale modyfikować może jedna osoba.

Opcja C — prawdziwa współbieżna edycja

To jest najtrudniejsze.

Wtedy autosave pełnego pliku co 30 sekund zwykle nie wystarcza. Potrzebujesz modelu zmian na poziomie operacji, np.:

User A inserted text at position X
User B deleted paragraph Y
User C changed style Z

Do tego potrzebujesz mechanizmu scalania zmian, rozwiązywania konfliktów, broadcastu zmian do innych klientów itd.

Jeżeli obecnie masz po prostu:

co 30 sekund wyślij aktualny plik do backendu

to nie masz jeszcze bezpiecznej współbieżnej edycji. W takim modelu dwóch użytkowników może sobie nadpisywać zmiany.

Dlatego na start rekomendowałbym:

jeden aktywny edytujący + lock z lease/heartbeat
