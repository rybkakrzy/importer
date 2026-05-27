import { Component } from '@angular/core';
import { EnvironmentBannerComponent } from '../environment-banner/environment-banner';
import { OfflineBannerComponent } from '../offline-banner/offline-banner';

/**
 * Container for the application's global top banners.
 *
 * Lives in the normal flex flow at the top of the app shell, so the layout
 * reserves the banners' height and they never overlap the editor / toolbar /
 * side panel. Stacking order is fixed here: environment banner on top, the
 * API/offline banner directly below it.
 */
@Component({
  selector: 'd2-global-banners',
  standalone: true,
  imports: [EnvironmentBannerComponent, OfflineBannerComponent],
  templateUrl: './global-banners.html',
  styleUrl: './global-banners.scss',
})
export class GlobalBannersComponent {}
