// the read-only paged surface: what the judge's copy will look like, and
// what the print button prints.
//
// one continuous flow, N windows into it. the whole document is laid out
// once at the width of the printed column, and each page shows the slice of
// that layout belonging to it. laying each page out separately was the first
// attempt and it was wrong: margins collapse differently once a paragraph is
// on its own, so the pages disagreed with the measurement they came from.
//
// the header and footer drawn here are built from the document's metadata,
// not from the donor's `header1.xml` — that part is word markup, a floating
// text box in the real school template, and there is no honest way to render
// it as html. the printed docx carries the school's own header; this is a
// preview of the page, not of their letterhead.

import { EDITOR_SELECTOR } from '../host/cardmirror.js';
import { buildFlow, flowBlocks, measureBlocks } from './measure.js';
import { paginate, usableHeightPx, usableWidthPx } from './paginate.js';
import { printStyles, twipsToPx } from './print.js';
import type { DocMeta } from '../docx/headers.js';
import type { Profile } from '../profile/profile.js';

export const PAGE_VIEW_ID = 'laymirror-page-view';
const STYLE_ID = 'laymirror-page-style';

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
  background: #55575c;
  padding: 16px 0 48px;
}
#${PAGE_VIEW_ID} .lm-chrome {
  position: sticky;
  top: 0;
  z-index: 1;
  display: flex;
  gap: 8px;
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
  margin: 0 auto 20px;
  background: #fff;
  color: #000;
  box-shadow: 0 4px 18px rgba(0,0,0,.5);
  overflow: hidden;
  zoom: 1;
}
#${PAGE_VIEW_ID} .lm-head,
#${PAGE_VIEW_ID} .lm-foot {
  position: absolute;
  left: ${twipsToPx(m.left)}px;
  width: ${usableWidthPx(profile.page)}px;
  font: 700 10pt/1.3 var(--pmd-body-font, serif);
  color: #000;
}
#${PAGE_VIEW_ID} .lm-head {
  top: ${twipsToPx(m.header)}px;
  border-bottom: 1px solid #000;
  padding-bottom: 2px;
  display: flex;
  justify-content: space-between;
  gap: 12px;
}
#${PAGE_VIEW_ID} .lm-foot {
  bottom: ${twipsToPx(m.footer)}px;
  text-align: center;
  font-weight: 400;
}
#${PAGE_VIEW_ID} .lm-body {
  position: absolute;
  top: ${twipsToPx(m.top)}px;
  left: ${twipsToPx(m.left)}px;
  width: ${usableWidthPx(profile.page)}px;
  height: ${usableHeightPx(profile.page)}px;
  overflow: hidden;
}
/* the window onto the flow: as tall as this page's share of it */
#${PAGE_VIEW_ID} .lm-window { overflow: hidden; }
#${PAGE_VIEW_ID} .lm-stage {
  position: absolute;
  left: -10000px;
  top: 0;
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
  document.body.append(root);

  // measured in the page's own column, at the page's own width, or every
  // line count belongs to some other layout
  const stage = el('div', 'lm-stage');
  const measured = buildFlow(blocks, profile);
  stage.append(measured);
  root.append(stage);

  const metrics = measureBlocks(flowBlocks(measured), profile);
  const { pages } = paginate(metrics, profile.page);
  stage.remove();

  root.append(chrome(pages.length));

  let offset = 0;
  pages.forEach((box, index) => {
    const sheet = el('div', 'lm-page');
    sheet.append(header(meta));

    const body = el('div', 'lm-body');
    const window_ = el('div', 'lm-window');
    // this page's share of the flow, and no more: a page that ends early
    // because a block would not fit must show the space, not the block
    window_.style.height = `${box.heightPx}px`;

    const flow = buildFlow(blocks, profile);
    flow.style.marginTop = `${-offset}px`;
    window_.append(flow);
    body.append(window_);

    sheet.append(body);
    sheet.append(footer(index + 1, pages.length));
    root.append(sheet);
    offset += box.heightPx;
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
