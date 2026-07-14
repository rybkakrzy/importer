import { Injectable, signal } from '@angular/core';

/** Migawka ostatniego błędu HTTP — na potrzeby diagnostyki zgłoszeń problemów. */
export interface HttpErrorSnapshot {
  /** Kod statusu HTTP (0 = brak połączenia). */
  status: number;
  /** Metoda żądania (GET/POST/…). */
  method: string;
  /** URL żądania. */
  url: string;
  /** Komunikat/detal błędu wyliczony przez interceptor. */
  detail: string;
  /** Czas wystąpienia (lokalny, pl-PL). */
  at: string;
}

/**
 * Serwis przechowujący migawkę OSTATNIEGO błędu HTTP.
 *
 * Interceptor błędów zapisuje tu każdy błąd; funkcja „Zgłoś problem" dołącza go
 * do treści zgłoszenia. Wcześniej ta informacja przepadała (tylko console.error + toast).
 *
 * Lekki (same sygnały, brak zależności) — bezpieczny do wstrzyknięcia wszędzie.
 */
@Injectable({ providedIn: 'root' })
export class LastHttpErrorService {
  /** Ostatni zarejestrowany błąd HTTP (null, jeśli sesja przebiegła bez błędów). */
  readonly lastError = signal<HttpErrorSnapshot | null>(null);

  /** Zapisuje migawkę błędu (wołane z interceptora). */
  record(snapshot: HttpErrorSnapshot): void {
    this.lastError.set(snapshot);
  }

  /** Czyści migawkę. */
  clear(): void {
    this.lastError.set(null);
  }
}
