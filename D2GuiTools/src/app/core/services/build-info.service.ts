import { Injectable, inject, computed } from '@angular/core';
import { ConnectionStatusService } from './connection-status.service';
import { environment as env } from '../../../environments/environment';

/**
 * Serwis informacji o buildzie.
 *
 * Dane z API (via ConnectionStatusService health-check) mają priorytet.
 * Gdy API niedostępne — fallback na wersję/datę builda frontu.
 * Aktualizuje się automatycznie co 30 s wraz z health-checkiem.
 */
@Injectable({ providedIn: 'root' })
export class BuildInfoService {
  private readonly connectionStatus = inject(ConnectionStatusService);

  readonly environment = computed(() =>
    this.connectionStatus.apiEnvironment() ?? env.environmentName
  );

  readonly buildNumber = computed(() =>
    this.connectionStatus.apiBuildNumber() ?? env.buildVersion
  );

  readonly buildDate = computed(() =>
    this.connectionStatus.apiBuildDate() ?? env.buildDate
  );

  /** Czy dane pochodzą z API (true) czy z lokalnego builda frontu (false) */
  readonly isApiData = computed(() =>
    this.connectionStatus.apiBuildNumber() !== null
  );
}
