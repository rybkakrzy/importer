import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
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
export class App {}
