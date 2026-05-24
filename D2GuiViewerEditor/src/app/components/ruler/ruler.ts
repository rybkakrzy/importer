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

  /**
   * Wcięcie paragrafu (cm) dla aktualnie zaznaczonego bloku (P/UL/OL/LI/TABLE/IMG…).
   * Gdy ustawione (mode='horizontal'), uchwyty linijki reprezentują lewą/prawą krawędź
   * paragrafu (margines strony + wcięcie), a przeciągnięcie emituje `blockIndentChange`
   * zamiast modyfikować marginesy strony — tak jak w MS Word.
   */
  @Input() blockIndent: { start: number; end: number } | null = null;

  /** Emituje nowe marginesy po zakończeniu przeciągania */
  @Output() marginsChange = new EventEmitter<PageMargins>();

  /**
   * Emituje nowe wcięcie paragrafu (cm) dla zaznaczonego bloku po zakończeniu
   * przeciągania. Tylko gdy `blockIndent` jest niepuste (mode='horizontal').
   */
  @Output() blockIndentChange = new EventEmitter<{ start?: number; end?: number }>();

  /**
   * Emituje stan linii prowadzącej (jak w MS Word) podczas przeciągania uchwytu.
   * `offsetPx` to NIEZSKALOWANA odległość od krawędzi strony (lewej dla horizontal,
   * górnej dla vertical) — rodzic renderuje kreskę nad kartką, bo wewnątrz linijki
   * (overflow:hidden, 22px) byłaby przycięta.
   */
  @Output() dragGuideChange = new EventEmitter<{ active: boolean; axis: 'horizontal' | 'vertical'; offsetPx: number }>();

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

  /** Długość osi w px ze zoomem (faktyczna szerokość/wysokość do wyrenderowania) */
  get axisPxScaled(): number {
    return Math.round(this.axisPx * (this.zoomLevel / 100));
  }

  /** Scale factor for positioning elements */
  get scale(): number {
    return this.zoomLevel / 100;
  }

  /** Ticki (co 1 cm) */
  get ticks(): number[] {
    return Array.from({ length: Math.floor(this.axisCm) + 1 }, (_, i) => i);
  }

  /** Margines "bliższy" (lewy / górny) w px - scaled.
   *  W trybie poziomym z `blockIndent` uwzględnia wcięcie paragrafu. */
  get startMarginPx(): number {
    const baseCm = this.mode === 'horizontal' ? this.activeMargins.left : this.activeMargins.top;
    const indentCm = this.mode === 'horizontal' && this.activeBlockIndent ? this.activeBlockIndent.start : 0;
    return (baseCm + indentCm) * this.CM_TO_PX * this.scale;
  }

  /** Margines "dalszy" (prawy / dolny) w px - scaled.
   *  W trybie poziomym z `blockIndent` uwzględnia wcięcie paragrafu. */
  get endMarginPx(): number {
    const baseCm = this.mode === 'horizontal' ? this.activeMargins.right : this.activeMargins.bottom;
    const indentCm = this.mode === 'horizontal' && this.activeBlockIndent ? this.activeBlockIndent.end : 0;
    return (baseCm + indentCm) * this.CM_TO_PX * this.scale;
  }

  // ────── Drag state ──────

  private _dragging: 'start' | 'end' | null = null;
  private _dragStartClientPx = 0;
  private _dragStartMarginCm = 0;
  isDragging = signal(false);
  dragIndicatorPos = signal(0);
  activeSide = signal<'start' | 'end' | null>(null);

  private _tempMargins: PageMargins | null = null;
  private _tempBlockIndent: { start: number; end: number } | null = null;

  get activeMargins(): PageMargins {
    return this._tempMargins ?? this.margins;
  }

  get activeBlockIndent(): { start: number; end: number } | null {
    return this._tempBlockIndent ?? this.blockIndent;
  }

  /** Czy w trybie wcięcia paragrafu (horizontal + blockIndent dostarczone) */
  private get isParagraphIndentMode(): boolean {
    return this.mode === 'horizontal' && this.blockIndent != null;
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['margins']) {
      this._tempMargins = null;
    }
    if (changes['blockIndent']) {
      this._tempBlockIndent = null;
    }
  }

  // ────── Handles ──────

  onStartHandleDown(e: MouseEvent): void {
    e.preventDefault();
    e.stopPropagation();
    // W trybie wcięcia paragrafu pozycja uchwytu = pageMargin + blockIndent
    const baseCm = this.mode === 'horizontal' ? this.margins.left : this.margins.top;
    const indentCm = this.isParagraphIndentMode ? (this.blockIndent?.start ?? 0) : 0;
    this._begin('start', this.mode === 'horizontal' ? e.clientX : e.clientY, baseCm + indentCm);
  }

  onEndHandleDown(e: MouseEvent): void {
    e.preventDefault();
    e.stopPropagation();
    const baseCm = this.mode === 'horizontal' ? this.margins.right : this.margins.bottom;
    const indentCm = this.isParagraphIndentMode ? (this.blockIndent?.end ?? 0) : 0;
    this._begin('end', this.mode === 'horizontal' ? e.clientX : e.clientY, baseCm + indentCm);
  }

  // ────── Global events ──────

  @HostListener('document:mousemove', ['$event'])
  onMouseMove(e: MouseEvent): void {
    if (!this._dragging) return;
    e.preventDefault();

    const client = this.mode === 'horizontal' ? e.clientX : e.clientY;
    const deltaCm = (client - this._dragStartClientPx) / (this.CM_TO_PX * this.scale);

    let newCm = this._dragging === 'start'
      ? this._dragStartMarginCm + deltaCm
      : this._dragStartMarginCm - deltaCm;

    // Pozycja uchwytu vs przeciwległa (żeby zachować ≥1 cm treści)
    const oppositePosition = this._dragging === 'start'
      ? (this.mode === 'horizontal'
          ? this.margins.right + (this.isParagraphIndentMode ? (this.blockIndent?.end ?? 0) : 0)
          : this.margins.bottom)
      : (this.mode === 'horizontal'
          ? this.margins.left + (this.isParagraphIndentMode ? (this.blockIndent?.start ?? 0) : 0)
          : this.margins.top);

    if (this.isParagraphIndentMode) {
      // Wcięcie może być ujemne (paragraf wychodzi poza margines strony), ale nie poza
      // krawędź kartki ani zbliży się do przeciwległego uchwytu na <1cm.
      newCm = Math.max(0.1, Math.min(this.axisCm - oppositePosition - 1, newCm));
    } else {
      newCm = Math.max(0.3, Math.min(this.axisCm - oppositePosition - 1, newCm));
    }
    newCm = Math.round(newCm * 100) / 100;

    if (this.isParagraphIndentMode) {
      // newCm = nowa pozycja krawędzi od strony strony → wcięcie = newCm - pageMargin
      const pageMarginCm = this._dragging === 'start' ? this.margins.left : this.margins.right;
      const indentCm = Math.max(-pageMarginCm + 0.1, Math.round((newCm - pageMarginCm) * 100) / 100);
      const base = this.activeBlockIndent ?? { start: 0, end: 0 };
      this._tempBlockIndent = this._dragging === 'start'
        ? { ...base, start: indentCm }
        : { ...base, end: indentCm };
      this.dragIndicatorPos.set(newCm * this.CM_TO_PX * this.scale * (this._dragging === 'start' ? 1 : 0) +
        (this._dragging === 'end' ? this.axisPxScaled - newCm * this.CM_TO_PX * this.scale : 0));
      this._emitGuide(true);
      return;
    }

    const key = this._dragging === 'start'
      ? (this.mode === 'horizontal' ? 'left' : 'top')
      : (this.mode === 'horizontal' ? 'right' : 'bottom');

    this._tempMargins = { ...this.margins, [key]: newCm };
    this.dragIndicatorPos.set(
      this._dragging === 'start' ? newCm * this.CM_TO_PX * this.scale : this.axisPxScaled - newCm * this.CM_TO_PX * this.scale
    );
    this._emitGuide(true);
  }

  @HostListener('document:mouseup')
  onMouseUp(): void {
    if (!this._dragging) return;
    if (this.isParagraphIndentMode) {
      if (this._tempBlockIndent) {
        const side = this._dragging === 'start' ? { start: this._tempBlockIndent.start } : { end: this._tempBlockIndent.end };
        this.blockIndentChange.emit(side);
      }
    } else if (this._tempMargins) {
      this.marginsChange.emit({ ...this._tempMargins });
    }
    this._dragging = null;
    this.isDragging.set(false);
    this.activeSide.set(null);
    this._tempMargins = null;
    this._tempBlockIndent = null;
    this._emitGuide(false);
  }

  // ────── Helpers ──────

  private _begin(side: 'start' | 'end', clientStart: number, cm: number): void {
    this._dragging = side;
    this._dragStartClientPx = clientStart;
    this._dragStartMarginCm = cm;
    this.isDragging.set(true);
    this.activeSide.set(side);
    // Początkowa pozycja kreski = aktualna krawędź uchwytu
    this.dragIndicatorPos.set(
      side === 'start' ? cm * this.CM_TO_PX * this.scale : this.axisPxScaled - cm * this.CM_TO_PX * this.scale
    );
    this._emitGuide(true);
  }

  /** Emituje stan linii prowadzącej do rodzica (offset niezskalowany od krawędzi strony). */
  private _emitGuide(active: boolean): void {
    this.dragGuideChange.emit({
      active,
      axis: this.mode,
      offsetPx: this.dragIndicatorPos() / Math.max(this.scale, 0.0001)
    });
  }

  tickPos(cm: number): number {
    return cm * this.CM_TO_PX * this.scale;
  }

  tickLabel(cm: number): string {
    return cm > 0 && cm < this.axisCm ? `${cm}` : '';
  }

  get startTooltip(): string {
    if (this.isParagraphIndentMode) {
      const v = (this.activeBlockIndent?.start ?? 0).toFixed(2);
      return `Wcięcie z lewej: ${v} cm`;
    }
    const s = this.mode === 'horizontal' ? 'Lewy' : 'Górny';
    const v = (this.mode === 'horizontal' ? this.activeMargins.left : this.activeMargins.top).toFixed(2);
    return `${s} margines: ${v} cm`;
  }

  get endTooltip(): string {
    if (this.isParagraphIndentMode) {
      const v = (this.activeBlockIndent?.end ?? 0).toFixed(2);
      return `Wcięcie z prawej: ${v} cm`;
    }
    const s = this.mode === 'horizontal' ? 'Prawy' : 'Dolny';
    const v = (this.mode === 'horizontal' ? this.activeMargins.right : this.activeMargins.bottom).toFixed(2);
    return `${s} margines: ${v} cm`;
  }
}
