import { Component, OnInit, inject, DestroyRef, signal } from '@angular/core';
import { takeUntilDestroyed } from '@angular/core/rxjs-interop';
import { RouterOutlet } from '@angular/router';
import { MsalService, MsalBroadcastService } from '@azure/msal-angular';
import { InteractionStatus } from '@azure/msal-browser';
import { filter } from 'rxjs';
import { GlobalBannersComponent } from './components/global-banners/global-banners';
import { MSAL_CUSTOM_CONFIG } from './core/config/runtime-config';
import { primeApiToken } from './core/auth/api-token-primer';

@Component({
  selector: 'd2-root',
  imports: [RouterOutlet, GlobalBannersComponent],
  template: `
    <d2-global-banners />
    @if (!outletActivated()) {
      <div class="app-boot-loading" aria-live="polite">
        <span class="app-boot-spinner" aria-hidden="true"></span>
        <span class="app-boot-text">Ładowanie aplikacji…</span>
      </div>
    }
    <router-outlet (activate)="onOutletActivate()" />
  `,
  // Shell flex layout lives in global styles.scss (under `d2-root`): the routed
  // page components are inserted by <router-outlet> and don't inherit this
  // component's encapsulation attributes, so sizing them here would not apply.
  styles: [`
    :host {
      display: flex;
      flex-direction: column;
      height: 100vh;
      overflow: hidden;
    }

    .app-boot-loading {
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 14px;
      color: #6b7280;
      opacity: 0;
      animation: app-boot-fade-in .25s ease .2s forwards;
    }

    .app-boot-spinner {
      width: 36px;
      height: 36px;
      border-radius: 50%;
      border: 3px solid #e5e7eb;
      border-top-color: #ff6200;
      animation: app-boot-spin .8s linear infinite;
    }

    .app-boot-text {
      font-size: 14px;
      letter-spacing: .01em;
    }

    @keyframes app-boot-spin { to { transform: rotate(360deg); } }
    @keyframes app-boot-fade-in { to { opacity: 1; } }
  `]
})
export class App implements OnInit {
  private readonly msal = inject(MsalService);
  private readonly broadcast = inject(MsalBroadcastService);
  private readonly destroyRef = inject(DestroyRef);
  private readonly authConfig = inject(MSAL_CUSTOM_CONFIG);

  readonly outletActivated = signal(false);

  onOutletActivate(): void {
    this.outletActivated.set(true);
  }

  ngOnInit(): void {
    // Redirect completion is handled by MsalRedirectComponent (<app-redirect>), which owns the
    // single handleRedirectObservable() call. Here we only keep an active account selected once
    // MSAL is idle (post-redirect / on reload), so token acquisition always has an account.
    this.msal.instance.enableAccountStorageEvents();
    this.broadcast.inProgress$
      .pipe(
        filter((status) => status === InteractionStatus.None),
        takeUntilDestroyed(this.destroyRef),
      )
      .subscribe(() => {
        const accounts = this.msal.instance.getAllAccounts();
        if (accounts.length > 0 && !this.msal.instance.getActiveAccount()) {
          this.msal.instance.setActiveAccount(accounts[0]);
        }
        // MSAL is idle and (after a fresh redirect login) an account is now selected. Prime the API
        // token here too, so api_access_token exists immediately after login — not only on reload
        // (handled by the app initializer) or the first backend call (handled by the interceptor).
        void primeApiToken(this.msal.instance, this.authConfig);
      });
  }
}
