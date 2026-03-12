import {
  Component,
  Input,
  Output,
  EventEmitter,
  OnChanges,
  SimpleChanges,
  signal,
  HostListener
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { PageMargins } from '../../models/document.model';

/**
 * Komponent linijki (ruler) w stylu MS Word.
 * Wyświetla JEDNĄ linijkę (poziomą LUB pionową) z przesuwanymi uchwytami marginesów.
 *
 * Użycie:
 *   <d2-ruler mode="horizontal" .../>    ← linijka nad kartką
 *   <d2-ruler mode="vertical"   .../>    ← linijka obok kartki
 *
 * Wymiary A4: portrait 21 × 29.7 cm, landscape 29.7 × 21 cm
 * 1 cm ≈ 37.795 px @ 96 DPI
 */
@Component({
  selector: 'd2-ruler',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ruler.html',
  styleUrl: './ruler.scss'
})
export class RulerComponent implements OnChanges {

  /** Tryb: 'horizontal' (nad kartką) lub 'vertical' (obok kartki) */
  @Input() mode: 'horizontal' | 'vertical' = 'horizontal';

  /** Aktualne marginesy dokumentu (cm) */
  @Input() margins: PageMargins = { top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 };

  /** Orientacja strony */
  @Input() orientation: 'portrait' | 'landscape' = 'portrait';

  /** Poziom zoomu (%) */
  @Input() zoomLevel = 100;

  /** Emituje nowe marginesy po zakończeniu przeciągania */
  @Output() marginsChange = new EventEmitter<PageMargins>();

  readonly CM_TO_PX = 37.795;

  // ────── Geometria ──────

  get pageWidthCm(): number { return this.orientation === 'portrait' ? 21 : 29.7; }
  get pageHeightCm(): number { return this.orientation === 'portrait' ? 29.7 : 21; }

  /** Długość osi linijki w cm */
  get axisCm(): number {
    return this.mode === 'horizontal' ? this.pageWidthCm : this.pageHeightCm;
  }

  /** Długość osi w px */
  get axisPx(): number {
    return Math.round(this.axisCm * this.CM_TO_PX);
  }

  /** Ticki (co 1 cm) */
  get ticks(): number[] {
    return Array.from({ length: Math.floor(this.axisCm) + 1 }, (_, i) => i);
  }

  /** Margines "bliższy" (lewy / górny) w px */
  get startMarginPx(): number {
    return (this.mode === 'horizontal' ? this.activeMargins.left : this.activeMargins.top) * this.CM_TO_PX;
  }

  /** Margines "dalszy" (prawy / dolny) w px */
  get endMarginPx(): number {
    return (this.mode === 'horizontal' ? this.activeMargins.right : this.activeMargins.bottom) * this.CM_TO_PX;
  }

  // ────── Drag state ──────

  private _dragging: 'start' | 'end' | null = null;
  private _dragStartClientPx = 0;
  private _dragStartMarginCm = 0;
  isDragging = signal(false);
  dragIndicatorPos = signal(0);
  activeSide = signal<'start' | 'end' | null>(null);

  private _tempMargins: PageMargins | null = null;

  get activeMargins(): PageMargins {
    return this._tempMargins ?? this.margins;
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['margins']) {
      this._tempMargins = null;
    }
  }

  // ────── Handles ──────

  onStartHandleDown(e: MouseEvent): void {
    e.preventDefault();
    e.stopPropagation();
    const cm = this.mode === 'horizontal' ? this.margins.left : this.margins.top;
    this._begin('start', this.mode === 'horizontal' ? e.clientX : e.clientY, cm);
  }

  onEndHandleDown(e: MouseEvent): void {
    e.preventDefault();
    e.stopPropagation();
    const cm = this.mode === 'horizontal' ? this.margins.right : this.margins.bottom;
    this._begin('end', this.mode === 'horizontal' ? e.clientX : e.clientY, cm);
  }

  // ────── Global events ──────

  @HostListener('document:mousemove', ['$event'])
  onMouseMove(e: MouseEvent): void {
    if (!this._dragging) return;
    e.preventDefault();

    const scale = this.zoomLevel / 100;
    const client = this.mode === 'horizontal' ? e.clientX : e.clientY;
    const deltaCm = ((client - this._dragStartClientPx) / scale) / this.CM_TO_PX;

    let newCm = this._dragging === 'start'
      ? this._dragStartMarginCm + deltaCm
      : this._dragStartMarginCm - deltaCm;

    // Clamp: min 0.3 cm, keep ≥1 cm content
    const opposite = this._dragging === 'start'
      ? (this.mode === 'horizontal' ? this.margins.right : this.margins.bottom)
      : (this.mode === 'horizontal' ? this.margins.left : this.margins.top);
    newCm = Math.max(0.3, Math.min(this.axisCm - opposite - 1, newCm));
    newCm = Math.round(newCm * 100) / 100;

    const key = this._dragging === 'start'
      ? (this.mode === 'horizontal' ? 'left' : 'top')
      : (this.mode === 'horizontal' ? 'right' : 'bottom');

    this._tempMargins = { ...this.margins, [key]: newCm };
    this.dragIndicatorPos.set(
      this._dragging === 'start' ? newCm * this.CM_TO_PX : this.axisPx - newCm * this.CM_TO_PX
    );
  }

  @HostListener('document:mouseup')
  onMouseUp(): void {
    if (!this._dragging) return;
    if (this._tempMargins) {
      this.marginsChange.emit({ ...this._tempMargins });
    }
    this._dragging = null;
    this.isDragging.set(false);
    this.activeSide.set(null);
    this._tempMargins = null;
  }

  // ────── Helpers ──────

  private _begin(side: 'start' | 'end', clientStart: number, cm: number): void {
    this._dragging = side;
    this._dragStartClientPx = clientStart;
    this._dragStartMarginCm = cm;
    this.isDragging.set(true);
    this.activeSide.set(side);
  }

  tickPos(cm: number): number {
    return cm * this.CM_TO_PX;
  }

  tickLabel(cm: number): string {
    return cm > 0 && cm < this.axisCm ? `${cm}` : '';
  }

  get startTooltip(): string {
    const s = this.mode === 'horizontal' ? 'Lewy' : 'Górny';
    const v = (this.mode === 'horizontal' ? this.activeMargins.left : this.activeMargins.top).toFixed(2);
    return `${s} margines: ${v} cm`;
  }

  get endTooltip(): string {
    const s = this.mode === 'horizontal' ? 'Prawy' : 'Dolny';
    const v = (this.mode === 'horizontal' ? this.activeMargins.right : this.activeMargins.bottom).toFixed(2);
    return `${s} margines: ${v} cm`;
  }
}
