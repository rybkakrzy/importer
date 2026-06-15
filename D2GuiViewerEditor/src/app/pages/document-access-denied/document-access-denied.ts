import { Component } from '@angular/core';
import { RouterLink } from '@angular/router';

/**
 * Shared "access denied" view for document routes (DOCX editor and PDF viewer).
 * Shown by the route guard when the backend returns 403 for a restricted document,
 * so the editor/viewer never initializes.
 */
@Component({
  selector: 'd2-document-access-denied',
  standalone: true,
  imports: [RouterLink],
  templateUrl: './document-access-denied.html',
  styleUrl: './document-access-denied.scss'
})
export class DocumentAccessDeniedComponent {}
