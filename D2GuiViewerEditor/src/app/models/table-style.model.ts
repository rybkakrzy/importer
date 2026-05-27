/**
 * Model obramowań/linii tabeli (frontend). Obramowanie jest utrwalane jako style
 * inline na elementach `<td>` (border / border-top / -right / -bottom / -left) —
 * spójnie z dotychczasowymi akcjami tabeli — dzięki czemu przeżywa zapis/autozapis
 * (HTML → DOCX) i ponowne otwarcie. Świadomie nie używamy klas CSS (nie przetrwałyby
 * konwersji DOCX ani eksportu poza aplikację).
 */

export type TableBorderLineStyle = 'solid' | 'dashed' | 'dotted' | 'double' | 'none';

/** Gdzie zastosować linię (jak przyciski obramowań w MS Word). */
export type TableBorderScope =
  | 'all'
  | 'none'
  | 'outer'
  | 'inner'
  | 'inner-horizontal'
  | 'inner-vertical'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right';

export interface TableBorderSettings {
  color: string;
  /** Grubość w px. */
  width: number;
  style: TableBorderLineStyle;
}

export const DEFAULT_TABLE_BORDER: TableBorderSettings = {
  color: '#cccccc',
  width: 1,
  style: 'solid'
};

export const TABLE_BORDER_LINE_STYLES: { value: TableBorderLineStyle; label: string }[] = [
  { value: 'solid', label: 'Ciągła' },
  { value: 'dashed', label: 'Przerywana' },
  { value: 'dotted', label: 'Kropkowana' },
  { value: 'double', label: 'Podwójna' },
  { value: 'none', label: 'Brak' }
];

export const TABLE_BORDER_WIDTHS: { value: number; label: string }[] = [
  { value: 1, label: 'Cienka' },
  { value: 2, label: 'Standardowa' },
  { value: 3, label: 'Średnia' },
  { value: 4, label: 'Gruba' }
];

/** Bezpieczna, neutralna paleta kolorów linii. */
export const TABLE_BORDER_COLORS: string[] = [
  '#000000', '#404040', '#595959', '#808080', '#a6a6a6', '#cccccc',
  '#c00000', '#ed7d31', '#bf9000', '#548235', '#2f5597', '#7030a0'
];

/**
 * Definicje przycisków „gdzie narysować linię". `icon` to identyfikator wariantu
 * miniatury rysowanej w panelu (neutralne symbole, bez kopiowania ikon Office).
 */
export interface TableBorderScopeDef {
  scope: TableBorderScope;
  label: string;
}

export const TABLE_BORDER_SCOPES: TableBorderScopeDef[] = [
  { scope: 'all', label: 'Wszystkie' },
  { scope: 'none', label: 'Brak' },
  { scope: 'outer', label: 'Zewnętrzne' },
  { scope: 'inner', label: 'Wewnętrzne' },
  { scope: 'inner-horizontal', label: 'Poziome wewn.' },
  { scope: 'inner-vertical', label: 'Pionowe wewn.' },
  { scope: 'top', label: 'Górna' },
  { scope: 'bottom', label: 'Dolna' },
  { scope: 'left', label: 'Lewa' },
  { scope: 'right', label: 'Prawa' }
];
