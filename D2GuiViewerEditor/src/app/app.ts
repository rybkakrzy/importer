import { Component, OnInit, inject } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { MsalService } from '@azure/msal-angular';
import { GlobalBannersComponent } from './components/global-banners/global-banners';

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

  ngOnInit(): void {
    // Complete any redirect login and keep an active account selected for token acquisition.
    this.msal.instance.enableAccountStorageEvents();
    this.msal.handleRedirectObservable().subscribe(() => {
      const accounts = this.msal.instance.getAllAccounts();
      if (accounts.length > 0 && !this.msal.instance.getActiveAccount()) {
        this.msal.instance.setActiveAccount(accounts[0]);
      }
    });
  }
}
