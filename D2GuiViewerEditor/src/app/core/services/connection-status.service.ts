import { Injectable, signal, computed, OnDestroy, inject, NgZone } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { timeout, TimeoutError } from 'rxjs';
import { ApiConfigService } from './api-config.service';

export interface HealthResponse {
  status: string;
  environment: string;
  buildNumber: string;
  buildDate: string;
  timestamp: string;
}

/**
 * Serwis monitorujący połączenie z API.
 *
 * Wykrywa utratę łączności na dwa sposoby:
 *   1. Zdarzenia przeglądarki `online` / `offline`
 *   2. Cykliczny health-check do backendu (co 30 s)
 *
 * Parsuje odpowiedź health i udostępnia dane o buildzie API.
 */
@Injectable({ providedIn: 'root' })
export class ConnectionStatusService implements OnDestroy {
  private readonly http = inject(HttpClient);
  private readonly apiConfig = inject(ApiConfigService);
  private readonly ngZone = inject(NgZone);

  /** Flaga: przeglądarka nie ma sieci */
  private readonly _browserOffline = signal(!navigator.onLine);

  /** Flaga: API nie odpowiada (health-check) */
  private readonly _apiUnreachable = signal(false);

  /** Pasek został zamknięty ręcznie przez użytkownika */
  private readonly _dismissed = signal(false);

  /** Dane z ostatniej udanej odpowiedzi /api/health */
  readonly apiEnvironment = signal<string | null>(null);
  readonly apiBuildNumber = signal<string | null>(null);
  readonly apiBuildDate = signal<string | null>(null);

  /** Czy aplikacja jest offline (brak sieci LUB API nie odpowiada) */
  readonly isOffline = computed(() =>
    (this._browserOffline() || this._apiUnreachable()) && !this._dismissed()
  );

  /** Odwrotność — wygodny helper */
  readonly isOnline = computed(() => !this._browserOffline() && !this._apiUnreachable());

  private healthCheckInterval: ReturnType<typeof setInterval> | null = null;

  /** Strażnik: pojedyncze żądanie health w locie — zapobiega duplikatom przy nakładających
   *  się wyzwalaczach (start aplikacji + zdarzenie `online` + interwał). */
  private _checkInFlight = false;

  private onOnline = () => {
    this._browserOffline.set(false);
    this._dismissed.set(false);
    this.checkNow();
  };

  private onOffline = () => {
    this._browserOffline.set(true);
    this._dismissed.set(false);
  };

  constructor() {
    window.addEventListener('online', this.onOnline);
    window.addEventListener('offline', this.onOffline);

    // Uruchom health-check poza strefą Angulara, żeby nie wywoływać CD
    this.ngZone.runOutsideAngular(() => {
      this.healthCheckInterval = setInterval(() => this.checkNow(), 30_000);
    });

    // Pierwszy check natychmiast (przy konstrukcji). Eager instancjacja w app.config
    // (provideAppInitializer → checkNow) gwarantuje, że dzieje się to zaraz po starcie
    // aplikacji, niezależnie od tego, który komponent pierwszy wstrzyknie serwis.
    this.checkNow();
  }

  ngOnDestroy(): void {
    window.removeEventListener('online', this.onOnline);
    window.removeEventListener('offline', this.onOffline);
    if (this.healthCheckInterval) {
      clearInterval(this.healthCheckInterval);
    }
  }

  /**
   * Wywołane przez interceptor HTTP, gdy żądanie zakończy się status === 0
   * (brak połączenia z serwerem).
   */
  reportApiError(): void {
    this.ngZone.run(() => {
      this._apiUnreachable.set(true);
      this._dismissed.set(false);
    });
  }

  /**
   * Zamknij pasek offline (użytkownik kliknął X).
   * Pasek pojawi się ponownie, gdy stan się zmieni.
   */
  dismiss(): void {
    this._dismissed.set(true);
  }

  /**
   * Wykonaj health-check natychmiast (pierwszy strzał po starcie aplikacji oraz każdy kolejny).
   * Idempotentne: jeśli żądanie już trwa, kolejne wywołanie jest pomijane (brak duplikatów).
   * Publiczne — wołane przez `provideAppInitializer` zaraz po inicjalizacji aplikacji.
   */
  checkNow(): void {
    if (this._checkInFlight) return;            // dedup: jedno żądanie w locie
    if (!this.apiConfig.baseUrl) return;        // nie strzelaj zanim URL API jest znany (brak fałszywego offline)
    this._checkInFlight = true;

    this.http
      .get<HealthResponse>(`${this.apiConfig.baseUrl}/health`)
      .pipe(timeout(5_000))
      .subscribe({
        next: (res) => {
          this.ngZone.run(() => {
            this._checkInFlight = false;
            this._apiUnreachable.set(false);
            this.apiEnvironment.set(res.environment);
            this.apiBuildNumber.set(res.buildNumber);
            this.apiBuildDate.set(res.buildDate);
          });
        },
        error: (err) => {
          this.ngZone.run(() => {
            this._checkInFlight = false;
            if (err.status === 0 || err instanceof TimeoutError) {
              this._apiUnreachable.set(true);
              this._dismissed.set(false);
            } else {
              // API odpowiedziało (np. 404/500) — serwer działa
              this._apiUnreachable.set(false);
            }
          });
        }
      });
  }
}
