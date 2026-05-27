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
  styles: [`
    :host {
      display: flex;
      flex-direction: column;
      height: 100vh;
      overflow: hidden;
    }
    d2-global-banners {
      flex-shrink: 0;
    }
    d2-dashboard,
    d2-document-editor,
    d2-pdf-maintenance,
    d2-pdf-viewer,
    d2-admin-shell {
      flex: 1;
      min-height: 0;
      overflow: hidden;
    }
  `]
})
export class App {}
