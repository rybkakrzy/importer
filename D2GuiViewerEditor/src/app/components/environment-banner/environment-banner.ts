import { Component, computed, inject } from '@angular/core';
import { BuildInfoService } from '../../core/services/build-info.service';
import { resolveEnvironmentBanner } from '../../core/utils/environment-banner.util';

/**
 * Global banner indicating which environment the GUI is talking to.
 *
 * Presentational: the environment name comes from BuildInfoService (the single
 * source of truth — API health-check env, falling back to the front-end build
 * config). Visibility and styling are derived by the pure mapping util.
 */
@Component({
  selector: 'd2-environment-banner',
  standalone: true,
  templateUrl: './environment-banner.html',
  styleUrl: './environment-banner.scss',
})
export class EnvironmentBannerComponent {
  private readonly buildInfo = inject(BuildInfoService);

  protected readonly view = computed(() =>
    resolveEnvironmentBanner(this.buildInfo.environment()),
  );
}
