import { Routes } from '@angular/router';
import { DashboardComponent } from './pages/dashboard/dashboard';
import { DocumentEditorComponent } from './components/document-editor/document-editor';
import { PdfMaintenanceComponent } from './pages/pdf-maintenance/pdf-maintenance';

export const routes: Routes = [
  { path: '', component: DashboardComponent },
  { path: 'editor', component: DocumentEditorComponent },
  {
    path: 'viewer', loadComponent: () =>
      import('./pages/pdf-viewer/pdf-viewer').then((m) => m.PdfViewerComponent),
  },
  { path: 'pdf-maintenance', component: PdfMaintenanceComponent },
  {
    path: 'admin',
    loadComponent: () =>
      import('./pages/admin/admin-shell/admin-shell').then((m) => m.AdminShellComponent),
    children: [
      { path: '', redirectTo: 'files', pathMatch: 'full' },
      {
        path: 'files',
        loadComponent: () =>
          import('./pages/admin/admin-files/admin-files').then((m) => m.AdminFilesComponent),
      },
    ],
  },
  { path: '**', redirectTo: '' },
];
