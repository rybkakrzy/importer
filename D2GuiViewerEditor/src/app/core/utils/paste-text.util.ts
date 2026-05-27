/**
 * Pure helpers for the "Wklej tylko tekst" (Paste Text Only) feature.
 *
 * Intentionally framework-free and DOM-light so they are unit-testable under
 * Vitest/jsdom. The WYSIWYG editor wires these into its paste pipeline.
 */

const BLOCK_TAGS = new Set([
  'P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
  'LI', 'TR', 'BLOCKQUOTE', 'SECTION', 'ARTICLE', 'PRE',
]);

/**
 * Collapses clipboard whitespace into something predictable for a plain-text
 * paste: normalizes newlines, drops invisible characters, squeezes runs of
 * spaces, and caps consecutive blank lines at one (Word-like behavior).
 * Tabs are preserved because they carry table/column structure (Excel TSV).
 */
export function normalizeWhitespace(text: string): string {
  return text
    .replace(/\r\n?/g, '\n')
    .replace(/ /g, ' ')
    .replace(/[​-‍﻿]/g, '')
    .replace(/[ \t]*\n[ \t]*/g, '\n')
    .replace(/ {2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]+$/gm, '')
    .trim();
}

/**
 * Converts an HTML fragment to plain text while keeping the structure a user
 * expects from Word's "Keep Text Only": paragraphs/headings become line breaks,
 * list items get a marker, table cells are tab-separated. Only used as a
 * fallback when the clipboard offers no `text/plain`.
 */
export function htmlToText(html: string): string {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return normalizeWhitespace(extractText(doc.body));
}

function extractText(node: Node): string {
  let out = '';
  node.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      out += child.textContent ?? '';
      return;
    }
    if (child.nodeType === Node.ELEMENT_NODE) {
      out += renderElement(child as HTMLElement);
    }
  });
  return out;
}

function renderElement(el: HTMLElement): string {
  const tag = el.tagName;
  if (tag === 'BR') {
    return '\n';
  }
  if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'NOSCRIPT') {
    return '';
  }
  if (tag === 'LI') {
    return renderListItem(el);
  }
  if (tag === 'TD' || tag === 'TH') {
    return extractText(el) + '\t';
  }
  const inner = extractText(el);
  return BLOCK_TAGS.has(tag) ? inner + '\n' : inner;
}

function renderListItem(el: HTMLElement): string {
  const ordered = el.parentElement?.tagName === 'OL';
  const marker = ordered ? `${listItemIndex(el)}. ` : '• ';
  return marker + extractText(el) + '\n';
}

function listItemIndex(el: HTMLElement): number {
  let index = 1;
  let prev = el.previousElementSibling;
  while (prev) {
    if (prev.tagName === 'LI') {
      index++;
    }
    prev = prev.previousElementSibling;
  }
  return index;
}

/**
 * Resolves the plain-text payload for a "paste text only" action: prefer the
 * source-provided `text/plain`, fall back to converting `text/html`.
 */
export function resolvePlainText(plain: string, html: string): string {
  if (plain) {
    return normalizeWhitespace(plain);
  }
  return html ? htmlToText(html) : '';
}
