// the read-only paged surface: what the judge's copy will look like, and
// what the print button prints.
//
// the content is a clone of what cardmirror has already rendered, measured at
// the width of the text column and split by `paginate`. the clones carry
// cardmirror's own classes and sit inside a container that matches the lay
// stylesheet's scope, so page view and work view are styled by one set of
// rules and cannot disagree.
//
// the header and footer drawn here are built from the document's metadata,
// not from the donor's `header1.xml` — that part is word markup, a floating
// text box in the real school template, and there is no honest way to render
// it as html. the printed docx carries the school's own header; this is a
// preview of the page, not of their letterhead.

import { EDITOR_SELECTOR } from '../host/cardmirror.js';
import { measureBlocks, isPageBreak } from './measure.js';
import { paginate, usableHeightPx, usableWidthPx } from './paginate.js';
import { printStyles, twipsToPx } from './print.js';
import type { DocMeta } from '../docx/headers.js';
import type { Profile } from '../profile/profile.js';

export const PAGE_VIEW_ID = 'laymirror-page-view';
const STYLE_ID = 'laymirror-page-style';
/** the lay stylesheet is scoped to the editor, and these clones are not in
 *  it. wearing cardmirror's pane class puts them back inside that scope. */
const SCOPE_CLASS = 'pmd-pane-editor';

const el = (tag: string, className?: string, text?: string): HTMLElement => {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
};

function editorContent(): HTMLElement | null {
  return (
    document.querySelector<HTMLElement>('.ProseMirror') ??
    document.querySelector<HTMLElement>(EDITOR_SELECTOR)
  );
}

function styles(profile: Profile): string {
  const width = twipsToPx(profile.page.widthTwips);
  const height = twipsToPx(profile.page.heightTwips);
  const m = profile.page.margin;

  return `
#${PAGE_VIEW_ID} {
  position: fixed;
  inset: 0;
  z-index: 99998;
  overflow: auto;
  background: #4a4a4a;
  padding: 24px 0 48px;
}
#${PAGE_VIEW_ID} .lm-chrome {
  position: sticky;
  top: 0;
  display: flex;
  gap: 12px;
  justify-content: center;
  align-items: center;
  font: 13px/1.5 system-ui, sans-serif;
  color: #fff;
  padding-bottom: 16px;
}
#${PAGE_VIEW_ID} .lm-page {
  position: relative;
  box-sizing: border-box;
  width: ${width}px;
  height: ${height}px;
  margin: 0 auto 24px;
  background: #fff;
  color: #000;
  box-shadow: 0 6px 24px rgba(0,0,0,.45);
  overflow: hidden;
  /* cardmirror zooms #editor; a page is a page */
  zoom: 1;
}
#${PAGE_VIEW_ID} .lm-head,
#${PAGE_VIEW_ID} .lm-foot {
  position: absolute;
  left: ${twipsToPx(m.left)}px;
  width: ${usableWidthPx(profile.page)}px;
  font: 10pt/1.3 inherit;
}
#${PAGE_VIEW_ID} .lm-head {
  top: ${twipsToPx(m.header)}px;
  border-bottom: 1px solid #000;
  padding-bottom: 2px;
  display: flex;
  justify-content: space-between;
  gap: 12px;
  font-weight: 700;
}
#${PAGE_VIEW_ID} .lm-foot {
  bottom: ${twipsToPx(m.footer)}px;
  text-align: center;
}
#${PAGE_VIEW_ID} .lm-body {
  position: absolute;
  top: ${twipsToPx(m.top)}px;
  left: ${twipsToPx(m.left)}px;
  width: ${usableWidthPx(profile.page)}px;
  height: ${usableHeightPx(profile.page)}px;
  overflow: hidden;
}
#${PAGE_VIEW_ID} .lm-slice { overflow: hidden; }
#${PAGE_VIEW_ID} .lm-stage {
  position: absolute;
  left: -10000px;
  top: 0;
  width: ${usableWidthPx(profile.page)}px;
  visibility: hidden;
}
${printStyles(profile.page, `#${PAGE_VIEW_ID} .lm-page`, PAGE_VIEW_ID)}
`;
}

function applyStyles(css: string): void {
  let sheet = document.getElementById(STYLE_ID) as HTMLStyleElement | null;
  if (!sheet) {
    sheet = document.createElement('style');
    sheet.id = STYLE_ID;
    document.head.appendChild(sheet);
  }
  sheet.textContent = css;
}

/** a block prepared for measuring or for a page: cardmirror's own markup,
 *  inert. */
function cloneBlock(source: HTMLElement): HTMLElement {
  const copy = source.cloneNode(true) as HTMLElement;
  copy.removeAttribute('contenteditable');
  copy.removeAttribute('id');
  return copy;
}

let onKey: ((e: KeyboardEvent) => void) | null = null;

export function isPageViewOpen(): boolean {
  return document.getElementById(PAGE_VIEW_ID) !== null;
}

export function closePageView(): void {
  if (onKey) {
    document.removeEventListener('keydown', onKey, true);
    onKey = null;
  }
  document.getElementById(PAGE_VIEW_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();
}

export interface PageViewResult {
  pages: number;
}

/** returns null when there is nothing to lay out. */
export function openPageView(profile: Profile, meta: DocMeta): PageViewResult | null {
  closePageView();

  const content = editorContent();
  if (!content) return null;

  const blocks = Array.from(content.children).filter(
    (child): child is HTMLElement => child instanceof HTMLElement,
  );
  if (blocks.length === 0) return null;

  applyStyles(styles(profile));

  const root = el('div');
  root.id = PAGE_VIEW_ID;
  root.setAttribute('contenteditable', 'false');

  // measured in the page's own column, at the page's own width, or every
  // line count belongs to some other layout
  const stage = el('div', `lm-stage ${SCOPE_CLASS}`);
  const staged = blocks.map(cloneBlock);
  for (const copy of staged) stage.append(copy);
  root.append(stage);
  document.body.append(root);

  const metrics = measureBlocks(staged, profile);
  const { pages } = paginate(metrics, profile.page);
  stage.remove();

  root.append(chrome(pages.length));

  pages.forEach((box, index) => {
    const sheet = el('div', 'lm-page');
    sheet.append(header(meta));

    const body = el('div', `lm-body ${SCOPE_CLASS}`);
    for (const slice of box.slices) {
      const source = blocks[slice.block];
      if (!source) continue;

      const window_ = el('div', 'lm-slice');
      window_.style.height = `${slice.toPx - slice.fromPx}px`;

      const inner = el('div');
      inner.style.marginTop = `${-slice.fromPx}px`;
      // a manual break is chrome, not content: it must not print
      inner.append(isPageBreak(source) ? el('div') : cloneBlock(source));

      window_.append(inner);
      body.append(window_);
    }

    sheet.append(body);
    sheet.append(footer(index + 1, pages.length));
    root.append(sheet);
  });

  onKey = (event: KeyboardEvent) => {
    if (event.key !== 'Escape') return;
    event.preventDefault();
    event.stopPropagation();
    closePageView();
  };
  document.addEventListener('keydown', onKey, true);

  return { pages: pages.length };
}

function chrome(pageCount: number): HTMLElement {
  const bar = el('div', 'lm-chrome');
  bar.append(el('span', undefined, `${pageCount} page${pageCount === 1 ? '' : 's'}`));

  const print = el('button', undefined, 'print');
  print.addEventListener('click', () => window.print());
  bar.append(print);

  const close = el('button', undefined, 'close');
  close.addEventListener('click', closePageView);
  bar.append(close);

  return bar;
}

function header(meta: DocMeta): HTMLElement {
  const head = el('div', 'lm-head');
  head.append(el('span', undefined, [meta.teamCode, meta.authors].filter(Boolean).join(' · ')));
  head.append(el('span', undefined, meta.title));
  return head;
}

function footer(page: number, of: number): HTMLElement {
  return el('div', 'lm-foot', `Page ${page} of ${of}`);
}
