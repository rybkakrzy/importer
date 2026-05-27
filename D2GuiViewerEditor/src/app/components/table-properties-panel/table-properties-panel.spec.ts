import { TestBed, ComponentFixture } from '@angular/core/testing';
import { TablePropertiesPanelComponent } from './table-properties-panel';

describe('TablePropertiesPanelComponent', () => {
  let fixture: ComponentFixture<TablePropertiesPanelComponent>;
  let component: TablePropertiesPanelComponent;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [TablePropertiesPanelComponent],
    }).compileComponents();
    fixture = TestBed.createComponent(TablePropertiesPanelComponent);
    component = fixture.componentInstance;
  });

  async function render(): Promise<void> {
    await fixture.whenStable();
    fixture.detectChanges();
  }

  function el(selector: string): HTMLElement | null {
    return fixture.nativeElement.querySelector(selector);
  }

  it('shows the empty state and no action sections when no table is active', async () => {
    fixture.componentRef.setInput('hasActiveTable', false);
    await render();

    expect(el('.table-panel-empty')).not.toBeNull();
    expect(el('.table-panel-body')).toBeNull();
    expect(el('.table-panel-empty')!.textContent).toContain('Kliknij tabelę');
  });

  function titles(): (string | undefined)[] {
    return Array.from(
      fixture.nativeElement.querySelectorAll('.table-panel-section-title')
    ).map((n) => (n as HTMLElement).textContent?.trim());
  }

  it('shows layout sections on the default tab and a separated delete zone', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    await render();

    expect(el('.table-panel-empty')).toBeNull();
    expect(component.activeTab()).toBe('layout');
    expect(titles()).toEqual(['Wiersze i kolumny', 'Komórki', 'Rozmiar i rozkład', 'Wygląd']);
    expect(el('.table-panel-danger-zone .table-panel-btn-danger')).not.toBeNull();
  });

  it('shows border sections only after switching to the Obramowania tab', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    await render();
    // Border sections are not in the DOM while the layout tab is active.
    expect(titles()).not.toContain('Gdzie narysować');

    component.setTab('border');
    await render();
    expect(titles()).toEqual(['Rodzaj linii', 'Grubość linii', 'Kolor linii', 'Gdzie narysować', 'Reset']);
  });

  it('shows the auto-detected scope label instead of a manual selector', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    fixture.componentRef.setInput('borderTargetInfo', { kind: 'range', rows: 2, cols: 3 });
    component.setTab('border');
    await render();

    // No manual scope control is rendered.
    expect(el('.table-panel-segmented')).toBeNull();
    const target = el('.table-panel-target');
    expect(target).not.toBeNull();
    expect(target!.textContent).toContain('zaznaczone komórki (2×3)');
  });

  it('emits every table action wired by the editor', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    await render();

    const fired: string[] = [];
    const outputs = [
      'insertRowAbove', 'insertRowBelow', 'insertColLeft', 'insertColRight',
      'deleteRow', 'deleteCol', 'deleteTable',
      'mergeCells', 'splitCell', 'splitTable',
      'autoFitContents', 'autoFitWindow', 'fixedWidth',
      'distributeRows', 'distributeCols', 'toggleGridLines', 'clearCellColor', 'close',
    ] as const;
    for (const name of outputs) {
      (component as any)[name].subscribe(() => fired.push(name));
    }

    const buttons = Array.from(
      fixture.nativeElement.querySelectorAll('button')
    ) as HTMLButtonElement[];
    buttons.forEach((b) => b.click());

    for (const name of outputs) {
      expect(fired).toContain(name);
    }
  });

  it('emits setCellColor with the chosen swatch color', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    fixture.componentRef.setInput('shadingColors', ['#FF0000', '#00FF00']);
    await render();

    let picked: string | undefined;
    component.setCellColor.subscribe((c) => (picked = c));

    const swatch = el('.table-panel-swatch') as HTMLButtonElement;
    expect(swatch).not.toBeNull();
    swatch.click();

    expect(picked).toBe('#FF0000');
  });

  it('emits one applyBorderScope per placement icon, in model order', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    component.setTab('border');
    await render();

    const scopeBtns = Array.from(
      fixture.nativeElement.querySelectorAll('.table-panel-scope')
    ) as HTMLButtonElement[];
    expect(scopeBtns.length).toBe(component.borderScopes.length);

    const scopes: string[] = [];
    component.applyBorderScope.subscribe((s) => scopes.push(s));
    scopeBtns.forEach((b) => b.click());
    expect(scopes).toEqual(component.borderScopes.map((s) => s.scope));
  });

  it('emits line style, width and color changes from their controls', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    component.setTab('border');
    await render();

    let style: string | undefined;
    let width: number | undefined;
    let color: string | undefined;
    component.borderStyleChange.subscribe((s) => (style = s));
    component.borderWidthChange.subscribe((w) => (width = w));
    component.borderColorChange.subscribe((c) => (color = c));

    // Pick the "dashed" line style chip.
    const styleChips = Array.from(fixture.nativeElement.querySelectorAll('.table-panel-chiprow')[0].querySelectorAll('.table-panel-chip')) as HTMLButtonElement[];
    styleChips[1].click(); // index 1 = dashed
    expect(style).toBe('dashed');

    // Pick a width chip (index 3 = "Gruba" / 4px).
    const widthChips = Array.from(fixture.nativeElement.querySelectorAll('.table-panel-chiprow')[1].querySelectorAll('.table-panel-chip')) as HTMLButtonElement[];
    widthChips[3].click();
    expect(width).toBe(4);

    // Pick the first colour swatch.
    (fixture.nativeElement.querySelector('.table-panel-swatch') as HTMLButtonElement).click();
    expect(color).toBe(component.borderColors[0]);
  });

  it('emits clear / restore-default / reset-settings from the Reset section', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    component.setTab('border');
    await render();

    let cleared = false, restored = false, reset = false;
    component.clearBorders.subscribe(() => (cleared = true));
    component.restoreDefaultBorders.subscribe(() => (restored = true));
    component.resetBorderSettings.subscribe(() => (reset = true));

    const byText = (t: string) =>
      (Array.from(fixture.nativeElement.querySelectorAll('button')).find(
        (b) => (b as HTMLElement).textContent?.trim() === t
      ) as HTMLButtonElement);

    byText('Usuń obramowania').click();
    byText('Domyślne obramowanie').click();
    byText('Reset linii').click();
    expect([cleared, restored, reset]).toEqual([true, true, true]);
  });

  it('border interactions never emit close (panel must not leave the table)', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    component.setTab('border');
    await render();

    let closed = false;
    component.close.subscribe(() => (closed = true));

    // Click every control on the border tab (scopes, chips, swatches, reset buttons).
    (Array.from(fixture.nativeElement.querySelectorAll('.table-panel-body button')) as HTMLButtonElement[])
      .forEach((b) => b.click());

    expect(closed).toBe(false);
  });

  it('reflects the grid-lines state in the toggle label', async () => {
    fixture.componentRef.setInput('hasActiveTable', true);
    fixture.componentRef.setInput('gridLinesVisible', true);
    await render();
    expect(el('.table-panel-toggle-state')!.textContent).toContain('Widoczne');

    fixture.componentRef.setInput('gridLinesVisible', false);
    await render();
    expect(el('.table-panel-toggle-state')!.textContent).toContain('Ukryte');
  });

  it('prevents default on mousedown for non-input targets to preserve the editor selection', () => {
    const button = document.createElement('button');
    const evt = new MouseEvent('mousedown', { cancelable: true });
    Object.defineProperty(evt, 'target', { value: button });
    component.onPanelMouseDown(evt);
    expect(evt.defaultPrevented).toBe(true);

    const input = document.createElement('input');
    const evt2 = new MouseEvent('mousedown', { cancelable: true });
    Object.defineProperty(evt2, 'target', { value: input });
    component.onPanelMouseDown(evt2);
    expect(evt2.defaultPrevented).toBe(false);
  });
});
