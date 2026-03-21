import {
  Component,
  OnInit,
  OnDestroy,
  inject,
  signal,
  ElementRef,
  ViewChild,
  HostListener,
  NgZone,
} from '@angular/core';
import { ActivatedRoute, RouterLink } from '@angular/router';
import { CommonModule } from '@angular/common';
import { getDocument, GlobalWorkerOptions, TextLayer } from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
import { DocumentStorageService } from '../../services/document-storage.service';

// Worker jest kopiowany do katalogu dist przez angular.json (assets)
GlobalWorkerOptions.workerSrc = '/pdf.worker.min.mjs';

@Component({
  selector: 'd2-pdf-viewer',
  standalone: true,
  imports: [CommonModule, RouterLink],
  templateUrl: './pdf-viewer.html',
  styleUrl: './pdf-viewer.scss',
})
export class PdfViewerComponent implements OnInit, OnDestroy {
  private route = inject(ActivatedRoute);
  private storageService = inject(DocumentStorageService);
  private ngZone = inject(NgZone);

  // ── State ───────────────────────────────────────────────────────────────────
  isLoading = signal(true);
  isRendering = signal(false);
  errorMessage = signal<string | null>(null);
  isNotFound = signal(false);
  totalPages = signal(0);
  currentPage = signal(1);
  scale = signal(1.5);
  documentName = signal('Dokument PDF');

  // ── Search ──────────────────────────────────────────────────────────────────
  showSearchBar = signal(false);
  searchQuery = signal('');
  searchResults = signal(0);
  currentMatchIndex = signal(0);

  private pdfDoc: PDFDocumentProxy | null = null;
  private intersectionObserver: IntersectionObserver | null = null;
  private textSpans: HTMLElement[] = [];
  private matchElements: HTMLElement[] = [];
  private highlightOverlays: HTMLElement[] = [];

  @ViewChild('viewerContainer') viewerContainer!: ElementRef<HTMLDivElement>;
  @ViewChild('searchInput') searchInput?: ElementRef<HTMLInputElement>;

  // ── Zoom ────────────────────────────────────────────────────────────────────
  readonly SCALE_STEPS = [0.5, 0.75, 1.0, 1.25, 1.5, 1.75, 2.0, 2.5, 3.0];

  get scalePercent(): number {
    return Math.round(this.scale() * 100);
  }

  get canZoomIn(): boolean {
    return this.scale() < this.SCALE_STEPS[this.SCALE_STEPS.length - 1];
  }

  get canZoomOut(): boolean {
    return this.scale() > this.SCALE_STEPS[0];
  }

  // ── Lifecycle ───────────────────────────────────────────────────────────────
  ngOnInit(): void {
    const masterId = this.route.snapshot.queryParamMap.get('masterId');
    if (!masterId) {
      this.errorMessage.set('Brak identyfikatora dokumentu.');
      this.isLoading.set(false);
      return;
    }

    this.storageService.getDocument(masterId).subscribe({
      next: (doc) => {
        this.documentName.set(doc.name);
        this.loadPdf(doc.content);
      },
      error: (err) => {
        if (err.status === 404) {
          this.isNotFound.set(true);
          this.errorMessage.set('Nie znaleziono dokumentu');
        } else {
          this.errorMessage.set('Nie można pobrać dokumentu z serwera.');
        }
        this.isLoading.set(false);
      },
    });
  }

  ngOnDestroy(): void {
    this.disconnectObserver();
    this.pdfDoc?.destroy();
  }

  // ── PDF loading ─────────────────────────────────────────────────────────────
  private async loadPdf(base64: string): Promise<void> {
    try {
      const binary = atob(base64);
      const data = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        data[i] = binary.charCodeAt(i);
      }

      const loadingTask = getDocument({ data });
      this.pdfDoc = await loadingTask.promise;
      this.totalPages.set(this.pdfDoc.numPages);
      this.isLoading.set(false);

      // Allow Angular to update the DOM (@if block) before accessing ViewChild
      setTimeout(() => this.renderAllPages(), 0);
    } catch {
      this.errorMessage.set('Błąd podczas parsowania pliku PDF.');
      this.isLoading.set(false);
    }
  }

  // ── Rendering ───────────────────────────────────────────────────────────────
  private async renderAllPages(): Promise<void> {
    if (!this.pdfDoc || !this.viewerContainer) return;
    this.isRendering.set(true);
    this.textSpans = [];
    this.clearHighlights();

    const container = this.viewerContainer.nativeElement;
    container.innerHTML = '';
    this.disconnectObserver();

    // Track which page is currently most visible in the viewport
    this.intersectionObserver = new IntersectionObserver(
      (entries) => {
        const visiblePages = entries
          .filter((e) => e.isIntersecting)
          .map((e) => parseInt((e.target as HTMLElement).dataset['page']!, 10))
          .filter((n) => !isNaN(n));

        if (visiblePages.length > 0) {
          const topmost = Math.min(...visiblePages);
          this.ngZone.run(() => this.currentPage.set(topmost));
        }
      },
      { root: container.closest('.pdf-scroll-area'), threshold: 0.1 }
    );

    for (let pageNum = 1; pageNum <= this.pdfDoc.numPages; pageNum++) {
      const wrapper = document.createElement('div');
      wrapper.className = 'pdf-page-wrapper';
      wrapper.dataset['page'] = String(pageNum);
      container.appendChild(wrapper);
      this.intersectionObserver.observe(wrapper);
      await this.renderPage(pageNum, wrapper);
    }

    this.isRendering.set(false);
  }

  private async renderPage(pageNum: number, wrapper: HTMLDivElement): Promise<void> {
    const page = await this.pdfDoc!.getPage(pageNum);
    const viewport = page.getViewport({ scale: this.scale() });

    // Canvas
    const canvas = document.createElement('canvas');
    canvas.height = viewport.height;
    canvas.width = viewport.width;
    wrapper.appendChild(canvas);
    await page.render({ canvas, viewport }).promise;

    // Text layer — enables selection & search highlighting
    const textLayerDiv = document.createElement('div');
    textLayerDiv.className = 'text-layer';
    // Inline styles are required because Angular's scoped CSS doesn't apply to
    // dynamically created elements (no _ngcontent attribute)
    textLayerDiv.style.position = 'absolute';
    textLayerDiv.style.top = '0';
    textLayerDiv.style.left = '0';
    textLayerDiv.style.width = `${viewport.width}px`;
    textLayerDiv.style.height = `${viewport.height}px`;
    textLayerDiv.style.overflow = 'hidden';
    textLayerDiv.style.lineHeight = '1';
    textLayerDiv.style.pointerEvents = 'none';
    wrapper.appendChild(textLayerDiv);

    const textContent = await page.getTextContent();
    const textLayer = new TextLayer({
      textContentSource: textContent,
      container: textLayerDiv,
      viewport,
    });
    await textLayer.render();

    // Collect non-empty spans for search
    const spans = Array.from(textLayerDiv.querySelectorAll('span')) as HTMLElement[];
    this.textSpans.push(...spans.filter((s) => !!s.textContent?.trim()));

    page.cleanup();
  }

  // ── Search ─────────────────────────────────────────────────────────────────────
  toggleSearch(): void {
    const next = !this.showSearchBar();
    this.showSearchBar.set(next);
    if (next) {
      setTimeout(() => this.searchInput?.nativeElement?.focus(), 50);
    } else {
      this.clearHighlights();
      this.searchQuery.set('');
      this.searchResults.set(0);
      this.currentMatchIndex.set(0);
    }
  }

  onSearchInput(e: Event): void {
    this.searchQuery.set((e.target as HTMLInputElement).value);
    this.performSearch();
  }

  performSearch(): void {
    this.clearHighlights();
    const query = this.searchQuery().trim().toLowerCase();
    if (!query) {
      this.searchResults.set(0);
      this.currentMatchIndex.set(0);
      return;
    }
    this.matchElements = this.textSpans.filter(
      (span) => span.textContent?.toLowerCase().includes(query)
    );
    this.matchElements.forEach((span) => this.createOverlay(span, query, false));
    this.searchResults.set(this.matchElements.length);
    if (this.matchElements.length > 0) {
      this.currentMatchIndex.set(1);
      this.activateMatch(0);
    } else {
      this.currentMatchIndex.set(0);
    }
  }

  nextMatch(): void {
    if (!this.matchElements.length) return;
    const nextIdx = this.currentMatchIndex() % this.matchElements.length;
    this.currentMatchIndex.set(nextIdx + 1);
    this.activateMatch(nextIdx);
  }

  prevMatch(): void {
    if (!this.matchElements.length) return;
    const prevIdx =
      (this.currentMatchIndex() - 2 + this.matchElements.length) % this.matchElements.length;
    this.currentMatchIndex.set(prevIdx + 1);
    this.activateMatch(prevIdx);
  }

  private activateMatch(index: number): void {
    this.highlightOverlays.forEach((ov, i) => {
      if (i === index) {
        ov.style.background = 'rgba(255, 98, 0, 0.6)';
        ov.style.outline = '2px solid rgba(255, 98, 0, 0.85)';
      } else {
        ov.style.background = 'rgba(255, 196, 0, 0.45)';
        ov.style.outline = '';
      }
    });
    this.scrollToOverlay(this.highlightOverlays[index]);
  }

  private createOverlay(span: HTMLElement, query: string, active: boolean): void {
    const wrapper = span.closest('.pdf-page-wrapper') as HTMLElement | null;
    if (!wrapper) return;

    // Use Range API to get the exact visual rect of only the matching substring.
    // getBoundingClientRect() on a Range accounts for PDF.js scaleX transforms and
    // returns the true visual bounds — not the full-line span bounds.
    const textNode = span.firstChild;
    if (!textNode || textNode.nodeType !== Node.TEXT_NODE) return;

    const fullText = (textNode.textContent ?? '').toLowerCase();
    const idx = fullText.indexOf(query);
    if (idx === -1) return;

    const range = document.createRange();
    range.setStart(textNode, idx);
    range.setEnd(textNode, idx + query.length);

    const rangeRect  = range.getBoundingClientRect();
    const wrapperRect = wrapper.getBoundingClientRect();

    if (rangeRect.width === 0 && rangeRect.height === 0) return;

    const overlay = document.createElement('div');
    overlay.style.position      = 'absolute';
    overlay.style.left          = `${rangeRect.left - wrapperRect.left}px`;
    overlay.style.top           = `${rangeRect.top  - wrapperRect.top}px`;
    overlay.style.width         = `${Math.max(rangeRect.width,  3)}px`;
    overlay.style.height        = `${Math.max(rangeRect.height, 3)}px`;
    overlay.style.background    = active ? 'rgba(255, 98, 0, 0.6)' : 'rgba(255, 196, 0, 0.45)';
    overlay.style.borderRadius  = '2px';
    overlay.style.pointerEvents = 'none';
    overlay.style.zIndex        = '20';
    if (active) overlay.style.outline = '2px solid rgba(255, 98, 0, 0.85)';
    wrapper.appendChild(overlay);
    this.highlightOverlays.push(overlay);
  }

  private scrollToOverlay(overlay: HTMLElement | undefined): void {
    if (!overlay) return;
    const scrollArea = this.viewerContainer?.nativeElement
      ?.closest('.pdf-scroll-area') as HTMLElement | null;
    if (!scrollArea) return;

    const ovRect   = overlay.getBoundingClientRect();
    const areaRect = scrollArea.getBoundingClientRect();
    const scrollTop = scrollArea.scrollTop
      + (ovRect.top  - areaRect.top)
      - (scrollArea.clientHeight / 2)
      + (ovRect.height / 2);
    scrollArea.scrollTo({ top: scrollTop, behavior: 'smooth' });
  }

  private clearHighlights(): void {
    this.highlightOverlays.forEach((ov) => ov.remove());
    this.highlightOverlays = [];
    this.matchElements = [];
  }

  // ── Zoom actions ─────────────────────────────────────────────────────────────
  async zoomIn(): Promise<void> {
    if (!this.canZoomIn || this.isRendering()) return;
    const idx = this.SCALE_STEPS.indexOf(this.scale());
    if (idx < this.SCALE_STEPS.length - 1) {
      this.scale.set(this.SCALE_STEPS[idx + 1]);
      await this.renderAllPages();
    }
  }

  async zoomOut(): Promise<void> {
    if (!this.canZoomOut || this.isRendering()) return;
    const idx = this.SCALE_STEPS.indexOf(this.scale());
    if (idx > 0) {
      this.scale.set(this.SCALE_STEPS[idx - 1]);
      await this.renderAllPages();
    }
  }

  // ── Keyboard shortcuts ───────────────────────────────────────────────────────
  @HostListener('window:keydown', ['$event'])
  onKeyDown(e: KeyboardEvent): void {
    if (this.isLoading() || this.errorMessage()) return;

    if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
      e.preventDefault();
      if (!this.showSearchBar()) {
        this.toggleSearch();
      } else {
        this.searchInput?.nativeElement?.focus();
      }
      return;
    }

    if (e.key === 'Escape' && this.showSearchBar()) {
      e.preventDefault();
      this.toggleSearch();
      return;
    }

    if (e.ctrlKey || e.metaKey) {
      if (e.key === '+' || e.key === '=') {
        e.preventDefault();
        this.zoomIn();
      } else if (e.key === '-') {
        e.preventDefault();
        this.zoomOut();
      }
    }
  }

  // ── Helpers ──────────────────────────────────────────────────────────────────
  private disconnectObserver(): void {
    this.intersectionObserver?.disconnect();
    this.intersectionObserver = null;
  }
}
