/**
 * Modele dla edytora dokumentów Word
 */

/** Zawartość dokumentu z konwersji DOCX */
export interface DocumentContent {
  html: string;
  metadata: DocumentMetadata;
  images: DocumentImage[];
  styles: DocumentStyle[];
}

/** Metadane dokumentu */
export interface DocumentMetadata {
  title?: string;
  author?: string;
  subject?: string;
  created?: string;
  modified?: string;
  pageCount?: number;
  wordCount?: number;
}

/** Styl dokumentu (Nagłówek 1, Normalny, itp.) */
export interface DocumentStyle {
  id: string;
  name: string;
  type: string; // paragraph, character
  basedOn?: string;
  nextStyle?: string;
  
  // Właściwości czcionki
  fontFamily?: string;
  fontSize?: number; // w punktach
  color?: string;
  isBold?: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
  
  // Właściwości paragrafu
  alignment?: string; // left, center, right, justify
  spaceBefore?: number; // w punktach
  spaceAfter?: number; // w punktach
  lineSpacing?: number; // mnożnik
  leftIndent?: number; // w cm
  rightIndent?: number; // w cm
  firstLineIndent?: number; // w cm
  
  // Poziom outline (dla nagłówków)
  outlineLevel?: number;
}

/** Marginesy dokumentu (w cm) */
export interface PageMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

/** Ustawienia strony */
export interface PageSettings {
  margins: PageMargins;
  orientation: 'portrait' | 'landscape';
  paperSize: 'a4' | 'letter' | 'legal';
}

/** Predefiniowane ustawienia marginesów */
export const MARGIN_PRESETS: { name: string; margins: PageMargins }[] = [
  { name: 'Normalne', margins: { top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 } },
  { name: 'Wąskie', margins: { top: 1.27, bottom: 1.27, left: 1.27, right: 1.27 } },
  { name: 'Średnie', margins: { top: 2.54, bottom: 2.54, left: 1.91, right: 1.91 } },
  { name: 'Szerokie', margins: { top: 2.54, bottom: 2.54, left: 5.08, right: 5.08 } },
  { name: 'Lustrzane', margins: { top: 2.54, bottom: 2.54, left: 3.18, right: 2.54 } },
];

/** Obraz osadzony w dokumencie */
export interface DocumentImage {
  id: string;
  contentType: string;
  base64Data: string;
}

/** Request zapisu dokumentu */
export interface SaveDocumentRequest {
  html: string;
  originalFileName?: string;
  metadata?: DocumentMetadata;
}

/** Szablon dokumentu */
export interface DocumentTemplate {
  id: string;
  name: string;
  description: string;
}

/** Odpowiedź z wgrywania obrazu */
export interface ImageUploadResponse {
  base64: string;
  fileName: string;
  size: number;
}

/** Konfiguracja formatowania tekstu */
export interface TextFormatting {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  subscript: boolean;
  superscript: boolean;
}

/** Styl paragrafu */
export interface ParagraphStyle {
  fontFamily: string;
  fontSize: number;
  textColor: string;
  backgroundColor: string;
  alignment: 'left' | 'center' | 'right' | 'justify';
  lineHeight: number;
}

/** Stan edytora */
export interface EditorState {
  isModified: boolean;
  canUndo: boolean;
  canRedo: boolean;
  wordCount: number;
  characterCount: number;
  fontSize?: number;
  fontFamily?: string;
  currentFormatting: TextFormatting;
  currentStyle: Partial<ParagraphStyle>;
}

/** Opcje eksportu */
export type ExportFormat = 'docx' | 'pdf' | 'html' | 'txt';

/** Poziom nagłówka */
export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/** Typ listy */
export type ListType = 'bullet' | 'numbered';

/** Komendy edytora */
export type EditorCommand = 
  | 'bold' | 'italic' | 'underline' | 'strikethrough'
  | 'subscript' | 'superscript'
  | 'alignLeft' | 'alignCenter' | 'alignRight' | 'alignJustify'
  | 'justifyLeft' | 'justifyCenter' | 'justifyRight' | 'justifyFull'
  | 'indent' | 'outdent'
  | 'bulletList' | 'numberedList'
  | 'insertUnorderedList' | 'insertOrderedList'
  | 'insertLink' | 'insertImage' | 'insertTable'
  | 'undo' | 'redo'
  | 'selectAll'
  | 'removeFormat'
  | 'heading1' | 'heading2' | 'heading3' | 'heading4' | 'heading5' | 'heading6'
  | 'paragraph';
