import { Component, OnInit, inject, DestroyRef } from '@angular/core';
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
    <router-outlet />
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
  `]
})
export class App implements OnInit {
  private readonly msal = inject(MsalService);
  private readonly broadcast = inject(MsalBroadcastService);
  private readonly destroyRef = inject(DestroyRef);
  private readonly authConfig = inject(MSAL_CUSTOM_CONFIG);

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
