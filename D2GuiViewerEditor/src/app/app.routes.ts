import { Routes } from '@angular/router';
import { DashboardComponent } from './pages/dashboard/dashboard';
import { DocumentEditorComponent } from './components/document-editor/document-editor';
import { PdfMaintenanceComponent } from './pages/pdf-maintenance/pdf-maintenance';
import { documentAccessGuard } from './guards/document-access.guard';
import { resourceGuard } from './guards/resource.guard';
import { authGuard } from './guards/auth.guard';

// Cała aplikacja wymaga logowania (authGuard → MsalGuard; bypass gdy auth.enabled=false). Dostęp do
// zasobów (editor/viewer/admin) gatinguje resourceGuard po nazwie trasy względem backendowej listy
// (GET /api/identity/resources); dokumenty dokładają documentAccessGuard (per-dokument 403).
// Backend = źródło prawdy; guardy frontowe są tylko UX.
export const routes: Routes = [
  { path: '', component: DashboardComponent, canActivate: [authGuard] },
  {
    path: 'editor', component: DocumentEditorComponent, canActivate: [authGuard, resourceGuard, documentAccessGuard] },
  {
    path: 'viewer',
    canActivate: [authGuard, resourceGuard, documentAccessGuard],
    loadComponent: () =>
      import('./pages/pdf-viewer/pdf-viewer').then((m) => m.PdfViewerComponent),
  },
  {
    path: 'access-denied',
    loadComponent: () =>
      import('./pages/document-access-denied/document-access-denied').then((m) => m.DocumentAccessDeniedComponent),
  },
  { path: 'pdf-maintenance', component: PdfMaintenanceComponent, canActivate: [authGuard] },
  {
    path: 'admin',
    canActivate: [authGuard, resourceGuard],
    loadComponent: () =>
      import('./pages/admin/admin-shell/admin-shell').then((m) => m.AdminShellComponent),
    children: [
      { path: '', redirectTo: 'files', pathMatch: 'full' },
      {
        path: 'files',
        loadComponent: () =>
          import('./pages/admin/admin-files/admin-files').then((m) => m.AdminFilesComponent),
      },
      {
        path: 'deliveries',
        loadComponent: () =>
          import('./pages/admin/admin-deliveries/admin-deliveries').then((m) => m.AdminDeliveriesComponent),
      },
    ],
  },
  { path: '**', redirectTo: '' },
];
