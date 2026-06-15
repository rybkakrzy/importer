import { Injectable, inject } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { RUNTIME_AUTH_CONFIG } from '../config/runtime-config';

/**
 * UX-facing identity derived from the signed-in Entra ID account (ID token claims).
 * For presentation only (e.g. hiding the admin menu) — the backend is the source of truth
 * for every authorization decision (access token claims).
 */
@Injectable({ providedIn: 'root' })
export class CurrentUserService {
  private readonly msal = inject(MsalService);
  private readonly authConfig = inject(RUNTIME_AUTH_CONFIG);

  private claims(): Record<string, unknown> | undefined {
    const account =
      this.msal.instance.getActiveAccount() ?? this.msal.instance.getAllAccounts()[0];
    return account?.idTokenClaims as Record<string, unknown> | undefined;
  }

  isAuthenticated(): boolean {
    return this.msal.instance.getAllAccounts().length > 0;
  }

  isAdmin(): boolean {
    const roles = (this.claims()?.['roles'] as string[] | undefined) ?? [];
    return roles.includes(this.authConfig.adminRole);
  }

  /** UX only; backend reads CorporateKey from the access token, never trusts the client. */
  corporateKey(): string | null {
    const value = this.claims()?.['ck'];
    return typeof value === 'string' && value.trim() ? value.trim() : null;
  }
}
