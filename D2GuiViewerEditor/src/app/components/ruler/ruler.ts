import {
  Component,
  Input,
  Output,
  EventEmitter,
  OnChanges,
  SimpleChanges,
  signal,
  HostListener, ChangeDetectionStrategy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { PageMargins } from '../../models/document.model';
import { CSS_PX_PER_CM } from '../../core/utils/units.util';

/**
 * Segment tekstu (kolumna) na poziomej linijce — pozycja i szerokość w cm
 * od LEWEJ krawędzi kartki. Dla układu wielokolumnowego sekcji linijka
 * dostaje po jednym segmencie na kolumnę.
 */
export interface RulerColumnSegment {
  startCm: number;
  widthCm: number;
}

/**
 * Komponent linijki (ruler) w stylu MS Word.
 * Wyświetla JEDNĄ linijkę (poziomą LUB pionową) z przesuwanymi uchwytami marginesów.
 *
 * Użycie:
 *   <d2-ruler mode="horizontal" .../>    ← linijka nad kartką
 *   <d2-ruler mode="vertical"   .../>    ← linijka obok kartki
 *
 * Wymiary A4: portrait 21 × 29.7 cm, landscape 29.7 × 21 cm
 * 1 cm = CSS_PX_PER_CM px @ 96 DPI (shared units.util constant)
 */
@Component({
  selector: 'd2-ruler',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ruler.html',
  styleUrl: './ruler.scss',
  // OnPush (ADR-0108 r.9): pionowa linijka jest renderowana PER STRONA, więc 128-stronicowy
  // dokument = 128 instancji; w trybie Default każdy cykl CD (każdy klik/zaznaczenie → zone)
  // re-renderował wszystkie tick-i (~1.8 s/cykl — edytor „zamierał"). Wejścia to prymitywy
  // i stabilne referencje (computed w rodzicu), drag idzie przez @HostListener (oznacza widok
  // jako dirty) i sygnały — OnPush nic nie traci.
  changeDetection: ChangeDetectionStrategy.OnPush
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
   * Opcjonalna długość osi linijki w cm (tylko vertical). Gdy ustawiona, linijka wypełnia
   * podziałką FAKTYCZNĄ wysokość kartki — strona z wysokim nagłówkiem bywa wyższa niż A4,
   * więc bez tego pod podziałką zostawałby pusty obszar. Null = domyślnie A4 (29.7 cm).
   */
  @Input() axisLengthCm: number | null = null;

  /**
   * Wcięcie paragrafu (cm) dla aktualnie zaznaczonego bloku (P/UL/OL/LI/TABLE/IMG…).
   * Gdy ustawione (mode='horizontal'), uchwyty linijki reprezentują lewą/prawą krawędź
   * paragrafu (margines strony + wcięcie), a przeciągnięcie emituje `blockIndentChange`
   * zamiast modyfikować marginesy strony — tak jak w MS Word.
   */
  @Input() blockIndent: { start: number; end: number } | null = null;

  /**
   * Kolumny sekcji, w której stoi kursor (tylko mode='horizontal'). Gdy ≥2 segmenty,
   * linijka — jak MS Word — pokazuje biały obszar per kolumna (szare strefy na marginesy
   * i odstępy między kolumnami), a uchwyty wcięć działają względem AKTYWNEJ kolumny.
   * Null / 1 segment = zwykły układ jednokolumnowy.
   */
  @Input() columnSegments: RulerColumnSegment[] | null = null;

  /** Indeks kolumny, w której stoi kursor (0-based). Używane tylko z `columnSegments`. */
  @Input() activeColumnIndex = 0;

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

  readonly CM_TO_PX = CSS_PX_PER_CM;

  // ────── Geometria ──────

  get pageWidthCm(): number { return this.orientation === 'portrait' ? 21 : 29.7; }
  get pageHeightCm(): number { return this.orientation === 'portrait' ? 29.7 : 21; }

  /** Długość osi linijki w cm */
  get axisCm(): number {
    if (this.mode === 'vertical' && this.axisLengthCm != null && this.axisLengthCm > 0) {
      return this.axisLengthCm;
    }
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

  /** Ticki (co 1 cm) — memoizowane po długości osi (nowa tablica przy każdym CD = pełny diff @for). */
  get ticks(): number[] {
    const n = Math.floor(this.axisCm) + 1;
    if (this._ticksCache.length !== n) this._ticksCache = Array.from({ length: n }, (_, i) => i);
    return this._ticksCache;
  }
  private _ticksCache: number[] = [];

  /**
   * Aktywny segment kolumny (cm od lewej krawędzi kartki) albo null, gdy kursor
   * nie stoi w układzie wielokolumnowym. W trybie kolumnowym zastępuje marginesy
   * strony jako punkt odniesienia uchwytów wcięć.
   */
  get activeSegment(): { start: number; width: number } | null {
    if (this.mode !== 'horizontal' || !this.columnSegments || this.columnSegments.length < 2) {
      return null;
    }
    const i = Math.max(0, Math.min(this.columnSegments.length - 1, this.activeColumnIndex));
    const s = this.columnSegments[i];
    return { start: s.startCm, width: s.widthCm };
  }

  /** Lewa krawędź obszaru tekstu (cm od lewej krawędzi kartki) — margines strony lub start aktywnej kolumny. */
  private get segStartCm(): number {
    const seg = this.activeSegment;
    if (seg) return seg.start;
    return this.mode === 'horizontal' ? this.activeMargins.left : this.activeMargins.top;
  }

  /** Odległość prawej krawędzi obszaru tekstu od PRAWEJ krawędzi kartki (cm). */
  private get segEndCm(): number {
    const seg = this.activeSegment;
    if (seg) return this.axisCm - (seg.start + seg.width);
    return this.mode === 'horizontal' ? this.activeMargins.right : this.activeMargins.bottom;
  }

  /** Margines "bliższy" (lewy / górny) w px - scaled.
   *  W trybie poziomym z `blockIndent` uwzględnia wcięcie paragrafu (względem aktywnej kolumny). */
  get startMarginPx(): number {
    const indentCm = this.mode === 'horizontal' && this.activeBlockIndent ? this.activeBlockIndent.start : 0;
    return (this.segStartCm + indentCm) * this.CM_TO_PX * this.scale;
  }

  /** Margines "dalszy" (prawy / dolny) w px - scaled.
   *  W trybie poziomym z `blockIndent` uwzględnia wcięcie paragrafu (względem aktywnej kolumny). */
  get endMarginPx(): number {
    const indentCm = this.mode === 'horizontal' && this.activeBlockIndent ? this.activeBlockIndent.end : 0;
    return (this.segEndCm + indentCm) * this.CM_TO_PX * this.scale;
  }

  /**
   * Szare strefy poziomej linijki (px, scaled). W trybie kolumnowym: wszystko poza
   * segmentami kolumn (marginesy strony + odstępy między kolumnami) — geometria sekcji,
   * bez wcięcia bloku. W trybie zwykłym: dwie strefy zgodne z pozycjami uchwytów.
   */
  get hGrayZones(): { left: number; width: number }[] {
    const segs = this.columnSegments;
    if (this.mode === 'horizontal' && segs && segs.length >= 2) {
      const zones: { left: number; width: number }[] = [];
      const px = (cm: number) => cm * this.CM_TO_PX * this.scale;
      let cursor = 0;
      for (const s of segs) {
        const left = px(s.startCm);
        if (left > cursor + 0.5) zones.push({ left: cursor, width: left - cursor });
        cursor = px(s.startCm + s.widthCm);
      }
      if (cursor < this.axisPxScaled - 0.5) {
        zones.push({ left: cursor, width: this.axisPxScaled - cursor });
      }
      return zones;
    }
    return [
      { left: 0, width: this.startMarginPx },
      { left: this.axisPxScaled - this.endMarginPx, width: this.endMarginPx }
    ];
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
    // W trybie wcięcia paragrafu pozycja uchwytu = start obszaru tekstu (margines
    // strony lub start aktywnej kolumny) + blockIndent
    const seg = this.activeSegment;
    const baseCm = seg ? seg.start : (this.mode === 'horizontal' ? this.margins.left : this.margins.top);
    const indentCm = this.isParagraphIndentMode ? (this.blockIndent?.start ?? 0) : 0;
    this._begin('start', this.mode === 'horizontal' ? e.clientX : e.clientY, baseCm + indentCm);
  }

  onEndHandleDown(e: MouseEvent): void {
    e.preventDefault();
    e.stopPropagation();
    const seg = this.activeSegment;
    const baseCm = seg
      ? this.axisCm - (seg.start + seg.width)
      : (this.mode === 'horizontal' ? this.margins.right : this.margins.bottom);
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

    // Punkt odniesienia uchwytów: margines strony albo krawędź aktywnej kolumny
    const seg = this.activeSegment;
    const baseStartCm = seg ? seg.start : (this.mode === 'horizontal' ? this.margins.left : this.margins.top);
    const baseEndCm = seg
      ? this.axisCm - (seg.start + seg.width)
      : (this.mode === 'horizontal' ? this.margins.right : this.margins.bottom);

    // Pozycja uchwytu vs przeciwległa (żeby zachować ≥1 cm treści)
    const oppositePosition = this._dragging === 'start'
      ? baseEndCm + (this.mode === 'horizontal' && this.isParagraphIndentMode ? (this.blockIndent?.end ?? 0) : 0)
      : baseStartCm + (this.mode === 'horizontal' && this.isParagraphIndentMode ? (this.blockIndent?.start ?? 0) : 0);

    // Kolumny są wąskie — treść między uchwytami może zejść poniżej 1 cm, ale nie do zera.
    const minContentCm = seg ? Math.min(1, Math.max(0.2, seg.width - 0.2)) : 1;

    if (this.isParagraphIndentMode) {
      // Wcięcie może być ujemne (paragraf wychodzi poza margines strony), ale nie poza
      // krawędź kartki ani zbliży się do przeciwległego uchwytu na <minContentCm.
      // W trybie kolumnowym uchwyt nie wychodzi poza swoją kolumnę (wcięcie ≥ 0).
      const minCm = seg ? (this._dragging === 'start' ? baseStartCm : baseEndCm) : 0.1;
      newCm = Math.max(minCm, Math.min(this.axisCm - oppositePosition - minContentCm, newCm));
    } else {
      newCm = Math.max(0.3, Math.min(this.axisCm - oppositePosition - minContentCm, newCm));
    }
    newCm = Math.round(newCm * 100) / 100;

    if (this.isParagraphIndentMode) {
      // newCm = nowa pozycja krawędzi od krawędzi kartki → wcięcie = newCm - baza
      const pageMarginCm = this._dragging === 'start' ? baseStartCm : baseEndCm;
      const minIndentCm = seg ? 0 : -pageMarginCm + 0.1;
      const indentCm = Math.max(minIndentCm, Math.round((newCm - pageMarginCm) * 100) / 100);
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
