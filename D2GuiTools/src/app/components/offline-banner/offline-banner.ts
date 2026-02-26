import { Component, inject } from '@angular/core';
import { ConnectionStatusService } from '../../core/services/connection-status.service';
import { BuildInfoService } from '../../core/services/build-info.service';

/**
 * Pasek informujący o pracy w trybie OFFLINE.
 *
 * Wyświetlany, gdy aplikacja nie ma połączenia z API
 * (brak internetu lub awaria backendu).
 * Można go zamknąć przyciskiem X — pojawi się ponownie,
 * jeśli stan połączenia zmieni się.
 */
@Component({
  selector: 'd2-offline-banner',
  standalone: true,
  templateUrl: './offline-banner.html',
  styleUrl: './offline-banner.scss'
})
export class OfflineBannerComponent {
  readonly connectionStatus = inject(ConnectionStatusService);
  readonly buildInfo = inject(BuildInfoService);
}
