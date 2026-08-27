// page view: the document as word will print it, and the way it gets printed.
//
// nothing here lays out text. docx-preview renders the real .docx — the
// school's own styles.xml, theme, header and footer, at the template's real
// page size and margins — and the browser lays it out. the previous page view
// cloned the editor's dom into an offscreen stage, which guaranteed every
// block was skipped by `content-visibility: auto` and measured as its
// placeholder height. this reads the file instead, so there is nothing to
// virtualize away.
//
// this is also the print pipeline. a plugin cannot hand a file to word —
// cardmirror's `openExternal` accepts http(s) and mailto and nothing else — so
// the honest route to paper is to render the real package here and let chromium
// print it. electron's print dialog offers "save as pdf", which is the pdf
// export as well.

import { renderAsync } from 'docx-preview';
import { hasRenderedBreaks, injectBreaksAt } from '../docx/breaks.js';
import { readText, toBlob, unzip, writeText, zip } from '../docx/zip.js';
import { fillPages } from './fill.js';

export const PREVIEW_ID = 'laymirror-page-view';
const STYLE_ID = 'laymirror-page-style';
const DOCUMENT = 'word/document.xml';

/** how the pages were decided, which the chrome states plainly — a preview
 *  that quietly guesses is worse than one that says it guessed. */
export type Pagination = 'word' | 'document' | 'estimated';

export interface PreviewResult {
  pages: number;
  pagination: Pagination;
}

const CHROME_CSS = `
#${PREVIEW_ID} {
  position: fixed;
  inset: 0;
  z-index: 99998;
  overflow: auto;
  background: #55575c;
}
#${PREVIEW_ID} .lm-bar {
  position: sticky;
  top: 0;
  z-index: 2;
  display: flex;
  gap: 10px;
  align-items: center;
  justify-content: center;
  padding: 8px 12px;
  background: #2b2d31;
  color: #fff;
  font: 13px/1.4 system-ui, sans-serif;
}
#${PREVIEW_ID} .lm-bar button {
  font: inherit;
  padding: 3px 10px;
  cursor: pointer;
}
#${PREVIEW_ID} .lm-note { opacity: .75 }
#${PREVIEW_ID} .docx-wrapper { background: transparent; padding: 20px 0 40px }
#${PREVIEW_ID} section.docx {
  background: #fff;
  box-shadow: 0 4px 18px rgba(0,0,0,.45);
  margin: 0 auto 20px;
}
@media print {
  body > *:not(#${PREVIEW_ID}) { display: none !important }
  #${PREVIEW_ID} { position: static !important; overflow: visible !important; background: #fff !important }
  #${PREVIEW_ID} .lm-bar { display: none !important }
  #${PREVIEW_ID} .docx-wrapper { padding: 0 !important }
  #${PREVIEW_ID} section.docx { box-shadow: none !important; margin: 0 !important; break-after: page }
  #${PREVIEW_ID} section.docx:last-child { break-after: auto }
}
`;

let onKey: ((e: KeyboardEvent) => void) | null = null;

export const isPreviewOpen = (): boolean => document.getElementById(PREVIEW_ID) !== null;

export function closePreview(): void {
  if (onKey) {
    document.removeEventListener('keydown', onKey, true);
    onKey = null;
  }
  document.getElementById(PREVIEW_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();
}

function ensureStyle(): void {
  if (document.getElementById(STYLE_ID)) return;
  const sheet = document.createElement('style');
  sheet.id = STYLE_ID;
  sheet.textContent = CHROME_CSS;
  document.head.appendChild(sheet);
}

const RENDER_OPTIONS = {
  className: 'docx',
  inWrapper: true,
  breakPages: true,
  // the whole point: replay word's own pagination when the file carries it
  ignoreLastRenderedPageBreak: false,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
} as const;

async function render(bytes: Uint8Array, into: HTMLElement): Promise<number> {
  into.replaceChildren();
  await renderAsync(toBlob(bytes), into, undefined, RENDER_OPTIONS);
  return into.querySelectorAll('section.docx').length;
}

/** the content box of a rendered page, in css pixels. */
function contentHeightOf(section: HTMLElement): number {
  const style = getComputedStyle(section);
  const box = section.getBoundingClientRect();
  const padding =
    Number.parseFloat(style.paddingTop || '0') + Number.parseFloat(style.paddingBottom || '0');
  const height = Number.parseFloat(style.minHeight || '0') || box.height;
  return Math.max(0, height - padding);
}

/** measure the one long page docx-preview produced and decide where it should
 *  have ended. the dom here is the renderer's own output — plain elements with
 *  no containment — so these heights are real. */
function estimateBreaks(root: HTMLElement): number[] {
  const section = root.querySelector<HTMLElement>('section.docx');
  if (!section) return [];
  const article = section.querySelector<HTMLElement>('article');
  if (!article) return [];

  const blocks = Array.from(article.children).filter(
    (child): child is HTMLElement => child instanceof HTMLElement,
  );
  if (blocks.length === 0) return [];

  const rects = blocks.map((el) => el.getBoundingClientRect());
  const heights = rects.map((rect, i) => {
    const next = rects[i + 1];
    // the gap to the next block already contains whatever margin survived
    // collapsing, which is not worth recomputing
    return next ? Math.max(rect.height, next.top - rect.top) : rect.height;
  });

  return fillPages(heights, contentHeightOf(section));
}

export interface PreviewOptions {
  /** shown in the bar so the user knows what they are looking at. */
  label?: string;
}

/** render the file at `bytes`. resolves once the pages are on screen. */
export async function openPreview(
  bytes: Uint8Array,
  options: PreviewOptions = {},
): Promise<PreviewResult> {
  closePreview();
  ensureStyle();

  const root = document.createElement('div');
  root.id = PREVIEW_ID;
  root.setAttribute('contenteditable', 'false');

  const bar = document.createElement('div');
  bar.className = 'lm-bar';
  root.append(bar);

  const body = document.createElement('div');
  root.append(body);
  document.body.append(root);

  const documentXml = readText(unzip(bytes), DOCUMENT) ?? '';
  let pages = await render(bytes, body);

  // word's own `lastRenderedPageBreak` is exact. failing that, more than one
  // rendered page means the file broke itself — a template that starts a new
  // page before every heading 1 does this without anyone typing a break — and
  // estimating on top of that would break every page twice.
  const pagination: Pagination = hasRenderedBreaks(documentXml)
    ? 'word'
    : pages > 1
      ? 'document'
      : 'estimated';

  // nothing in the file said where the pages end, so work it out from the
  // layout the renderer just produced and render once more with real breaks
  if (pagination === 'estimated') {
    const indices = estimateBreaks(body);
    if (indices.length > 0) {
      const parts = unzip(bytes);
      writeText(parts, DOCUMENT, injectBreaksAt(documentXml, indices));
      pages = await render(zip(parts), body);
    }
  }

  bar.append(chrome(pages, pagination, options));

  onKey = (event: KeyboardEvent) => {
    if (event.key !== 'Escape') return;
    event.preventDefault();
    event.stopPropagation();
    closePreview();
  };
  // capture, because cardmirror binds escape too
  document.addEventListener('keydown', onKey, true);

  return { pages, pagination };
}

const NOTE: Record<Pagination, string> = {
  word: 'pages as word laid them out',
  document: 'pages break where the template says',
  estimated: 'approximate — save from word for exact pages',
};

function chrome(pages: number, pagination: Pagination, options: PreviewOptions): DocumentFragment {
  const frag = document.createDocumentFragment();

  const count = document.createElement('span');
  count.textContent = `${pages} page${pages === 1 ? '' : 's'}`;
  frag.append(count);

  const note = document.createElement('span');
  note.className = 'lm-note';
  note.textContent = NOTE[pagination];
  frag.append(note);

  if (options.label) {
    const label = document.createElement('span');
    label.className = 'lm-note';
    label.textContent = options.label;
    frag.append(label);
  }

  // electron's print dialog carries "save as pdf", so this is the pdf export
  // as well as the printer — laymirror has no shell to hand the file to word
  const print = document.createElement('button');
  print.type = 'button';
  print.textContent = 'print / pdf';
  print.addEventListener('click', () => window.print());
  frag.append(print);

  const close = document.createElement('button');
  close.type = 'button';
  close.textContent = 'close';
  close.addEventListener('click', closePreview);
  frag.append(close);

  return frag;
}
