import { Component, computed, input } from '@angular/core';

/**
 * Presentational security label for a document's classification (from metadata).
 * Shared by the DOCX editor and the PDF viewer so the badge looks and behaves
 * identically regardless of file type.
 *
 * Renders nothing when the classification is absent/blank — never an empty badge.
 * Unknown values are shown verbatim (no allow-list enforced here); only the color
 * mapping recognises the project dictionary C1..C4. This is purely presentational
 * and does not affect document access.
 */
@Component({
  selector: 'd2-document-classification-badge',
  standalone: true,
  templateUrl: './document-classification-badge.html',
  styleUrl: './document-classification-badge.scss'
})
export class DocumentClassificationBadgeComponent {
  classification = input<string | null | undefined>(undefined);

  /** Trimmed value, or null when absent/blank/whitespace — single source for "should render". */
  readonly value = computed(() => {
    const raw = this.classification();
    if (typeof raw !== 'string') return null;
    const trimmed = raw.trim();
    return trimmed.length > 0 ? trimmed : null;
  });

  /** CSS modifier per known level (C1..C4); unknown values get a neutral style. */
  readonly levelClass = computed(() => {
    const v = this.value();
    if (!v) return '';
    const known: Record<string, string> = {
      C1: 'lvl-c1',
      C2: 'lvl-c2',
      C3: 'lvl-c3',
      C4: 'lvl-c4'
    };
    return known[v.toUpperCase()] ?? 'lvl-unknown';
  });
}
