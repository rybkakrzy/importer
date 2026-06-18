import { Component } from '@angular/core';
import { RouterLink } from '@angular/router';

/**
 * Dedicated "no permissions" (403) view. Shown when a route guard denies access (missing role /
 * resource) or the backend returns HTTP 403. Distinct from technical errors (404/500) which surface
 * as toasts, not a page.
 */
@Component({
  selector: 'd2-access-forbidden',
  standalone: true,
  imports: [RouterLink],
  templateUrl: './access-forbidden.html',
  styleUrl: './access-forbidden.scss',
})
export class AccessForbiddenComponent {}
