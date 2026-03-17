import { Routes } from '@angular/router';
import { DashboardComponent } from './pages/dashboard/dashboard';
import { DocumentEditorComponent } from './components/document-editor/document-editor';
import { PdfMaintenanceComponent } from './pages/pdf-maintenance/pdf-maintenance';

export const routes: Routes = [
  { path: '', component: DashboardComponent },
  { path: 'editor', component: DocumentEditorComponent },
  { path: 'pdf-maintenance', component: PdfMaintenanceComponent },
  { path: '**', redirectTo: '' }
];
