import { Component, EventEmitter, Input, Output } from '@angular/core';

/**
 * Side panel surfacing the header/footer editing controls — the previous
 * floating toolbar over the document was overlapping the workspace. This panel
 * lives in the same dock as Wyszukiwanie / d2-table-properties-panel and is
 * driven by the parent's `editingSection` signal.
 *
 * Stateless on purpose: every action is forwarded to the document editor / the
 * wysiwyg editor as an output so the existing logic stays the single source of
 * truth (rule 9 — no parallel mechanism).
 */
@Component({
  selector: 'd2-header-footer-panel',
  standalone: true,
  templateUrl: './header-footer-panel.html',
  styleUrl: './header-footer-panel.scss',
})
export class HeaderFooterPanelComponent {
  /** 'header' | 'footer' — drives the contextual labels. */
  @Input() section: 'header' | 'footer' | 'body' = 'header';
  /** Reflects the editor's `differentFirstPage()` so the checkbox is two-way. */
  @Input() differentFirstPage = false;

  @Output() close = new EventEmitter<void>();
  @Output() toggleDifferentFirstPage = new EventEmitter<void>();
  @Output() openFormatDialog = new EventEmitter<void>();
  @Output() insertImage = new EventEmitter<void>();
  @Output() insertPageNumbers = new EventEmitter<void>();
  @Output() removeSection = new EventEmitter<void>();

  protected sectionLabel(): string {
    return this.section === 'footer' ? 'Stopka' : 'Nagłówek';
  }

  protected formatLabel(): string {
    return this.section === 'footer' ? 'Format stopki' : 'Format nagłówka';
  }

  protected removeLabel(): string {
    return this.section === 'footer' ? 'Usuń stopkę' : 'Usuń nagłówek';
  }
}
