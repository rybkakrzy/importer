import { Component, EventEmitter, Input, Output, computed, signal } from '@angular/core';
import { FormsModule } from '@angular/forms';

/**
 * Stateless side panel for the currently selected image. Lives in the same dock as
 * Wyszukiwanie / d2-table-properties-panel / d2-header-footer-panel. The parent
 * (document-editor) owns the actual editor mutation — this component only surfaces
 * controls and emits intent (rule 9 — no parallel logic).
 *
 * Word-like positioning roadmap: this MVP covers size, lock aspect, alignment and
 * remove for inline images. Floating (wp:anchor) wrap modes are deferred to the
 * next iteration — see .ai/FEATURES.md for the limitation note.
 */
export type ImagePositionMode = 'inline' | 'front' | 'behind';

export interface ImageSelectionState {
  widthPx: number;
  heightPx: number;
  aspectRatio: number;
  alignment: 'left' | 'center' | 'right' | null;
  positionMode: ImagePositionMode;
}

@Component({
  selector: 'd2-image-properties-panel',
  standalone: true,
  imports: [FormsModule],
  templateUrl: './image-properties-panel.html',
  styleUrl: './image-properties-panel.scss',
})
export class ImagePropertiesPanelComponent {
  /** Snapshot of the selected image — driven by the parent. */
  @Input() set state(value: ImageSelectionState | null) {
    this._state.set(value);
    if (value) {
      this._widthInput.set(Math.round(value.widthPx));
      this._heightInput.set(Math.round(value.heightPx));
    }
  }

  @Input() lockAspect = true;

  @Output() close = new EventEmitter<void>();
  @Output() widthChange = new EventEmitter<number>();
  @Output() heightChange = new EventEmitter<number>();
  @Output() lockAspectChange = new EventEmitter<boolean>();
  @Output() alignmentChange = new EventEmitter<'left' | 'center' | 'right' | null>();
  @Output() removeImage = new EventEmitter<void>();
  @Output() resetAspect = new EventEmitter<void>();
  @Output() positionModeChange = new EventEmitter<ImagePositionMode>();

  private readonly _state = signal<ImageSelectionState | null>(null);
  // Mirror the state into editable inputs so typing doesn't fight the snapshot.
  protected readonly _widthInput = signal<number>(0);
  protected readonly _heightInput = signal<number>(0);

  protected readonly current = computed(() => this._state());

  protected applyWidth(): void {
    const w = this.clampPx(this._widthInput());
    this._widthInput.set(w);
    this.widthChange.emit(w);
  }

  protected applyHeight(): void {
    const h = this.clampPx(this._heightInput());
    this._heightInput.set(h);
    this.heightChange.emit(h);
  }

  protected toggleLock(checked: boolean): void {
    this.lockAspectChange.emit(checked);
  }

  protected align(value: 'left' | 'center' | 'right'): void {
    this.alignmentChange.emit(value);
  }

  protected clearAlign(): void {
    this.alignmentChange.emit(null);
  }

  protected setMode(mode: ImagePositionMode): void {
    this.positionModeChange.emit(mode);
  }

  /** Min 16 px — matches the hard floor in wysiwyg-editor's resize-end logic. */
  private clampPx(value: number): number {
    if (!Number.isFinite(value) || value <= 0) return 16;
    return Math.max(16, Math.round(value));
  }
}
