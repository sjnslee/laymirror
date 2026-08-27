// where the page ends, drawn in the editor.
//
// cardmirror shows a document as one unbroken column, so nothing on screen
// says where a printed page stops — which for a lay file, printed and handed
// to a parent judge, is most of what you want to know while writing it.
//
// two things put a break there, and laymirror draws both:
//
//   a style     the template says `w:pageBreakBefore` on heading 1, so every
//               pocket starts a page. this is where a lay file's breaks
//               actually come from — nobody types them — and because it is a
//               property of the style it survives a cardmirror round-trip
//               untouched. drawn with a css rule, so it costs nothing and
//               never fights prosemirror for the dom.
//
//   a character a document that came from somewhere else may carry a literal
//               form feed. cardmirror imports it as text and draws nothing at
//               all, which is the complaint. it cannot be styled — a css rule
//               cannot select one character — so those are measured and drawn
//               in an overlay that sits above the editor and touches nothing.

import { CLASS } from '../host/cardmirror.js';
import type { BlockType } from '../template/styles.js';

const STYLE_ID = 'laymirror-break-rules';
const OVERLAY_ID = 'laymirror-break-overlay';

/** form feed, vertical tab, and unicode's own line and paragraph separators —
 *  every character a word-derived pipeline uses to mean "new page here". */
const BREAK_CHARS = /[\f\v\u2028\u2029]/g;

const SELECTOR: Partial<Record<BlockType, string>> = {
  pocket: `h1.${CLASS.pocket}`,
  hat: `h2.${CLASS.hat}`,
  block: `h3.${CLASS.block}`,
  tag: `h4.${CLASS.tag}`,
  analytic: `p.${CLASS.analytic}`,
  undertag: `p.${CLASS.undertag}`,
};

const RULE = `
.ProseMirror %SEL% {
  position: relative;
}
.ProseMirror %SEL%:not(:first-child)::before {
  content: 'page';
  position: absolute;
  left: 0;
  right: 0;
  top: -1.05em;
  border-top: 2px dashed currentColor;
  opacity: .38;
  font: 600 9px/1.2 system-ui, sans-serif;
  letter-spacing: .12em;
  text-transform: uppercase;
  text-align: right;
  pointer-events: none;
}
.ProseMirror %SEL%:not(:first-child) {
  margin-top: 2.1em;
}
`;

const OVERLAY_CSS = `
#${OVERLAY_ID} {
  position: fixed;
  inset: 0;
  z-index: 40;
  pointer-events: none;
}
#${OVERLAY_ID} i {
  position: absolute;
  border-top: 2px dashed currentColor;
  opacity: .38;
}
`;

export const shown = (): boolean => document.getElementById(STYLE_ID) !== null;

export function clear(): void {
  document.getElementById(STYLE_ID)?.remove();
  document.getElementById(OVERLAY_ID)?.remove();
  stopTracking();
}

/** draw a rule above every block the template breaks a page before, plus one
 *  wherever the text itself carries a page-break character. returns how many
 *  of each, so the panel can say what it found. */
export function draw(
  editor: HTMLElement,
  breaks: readonly BlockType[],
): { styled: number; literal: number } {
  const selectors = breaks.map((type) => SELECTOR[type]).filter((sel): sel is string => !!sel);

  const sheet = document.getElementById(STYLE_ID) ?? document.createElement('style');
  sheet.id = STYLE_ID;
  sheet.textContent =
    selectors.map((sel) => RULE.replaceAll('%SEL%', sel)).join('') + OVERLAY_CSS;
  if (!sheet.isConnected) document.head.append(sheet);

  const styled = selectors.length
    ? editor.querySelectorAll(selectors.map((sel) => `${sel}:not(:first-child)`).join(',')).length
    : 0;

  startTracking(editor);
  return { styled, literal: paint(editor) };
}

// ── the overlay ───────────────────────────────────────────────────────

let tracked: HTMLElement | null = null;
let pending = 0;
let observer: MutationObserver | null = null;

const repaint = (): void => {
  if (pending) return;
  pending = requestAnimationFrame(() => {
    pending = 0;
    if (tracked?.isConnected) paint(tracked);
    else clear();
  });
};

function startTracking(editor: HTMLElement): void {
  if (tracked === editor) return;
  stopTracking();
  tracked = editor;
  observer = new MutationObserver(repaint);
  observer.observe(editor, { childList: true, subtree: true, characterData: true });
  window.addEventListener('scroll', repaint, true);
  window.addEventListener('resize', repaint);
}

function stopTracking(): void {
  observer?.disconnect();
  observer = null;
  tracked = null;
  window.removeEventListener('scroll', repaint, true);
  window.removeEventListener('resize', repaint);
}

/** every page-break character in the editor's text, as a dom range each. */
function literalBreaks(editor: HTMLElement): Range[] {
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  const out: Range[] = [];

  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.nodeValue ?? '';
    BREAK_CHARS.lastIndex = 0;
    for (let hit = BREAK_CHARS.exec(text); hit; hit = BREAK_CHARS.exec(text)) {
      const range = document.createRange();
      range.setStart(node, hit.index);
      range.setEnd(node, hit.index + 1);
      out.push(range);
    }
  }
  return out;
}

/** a break character has no width, so its own rect is a zero-wide sliver.
 *  the line is drawn across the editor's text column at that height instead,
 *  which is what a page break looks like. */
function paint(editor: HTMLElement): number {
  const layer = document.getElementById(OVERLAY_ID) ?? document.createElement('div');
  layer.id = OVERLAY_ID;
  if (!layer.isConnected) document.body.append(layer);
  layer.replaceChildren();

  const frame = editor.getBoundingClientRect();
  let drawn = 0;

  for (const range of literalBreaks(editor)) {
    const rect = range.getBoundingClientRect();
    // zero everywhere means the block is off screen and chromium never laid
    // it out — skipped rather than guessed at, and painted when it scrolls in
    if (rect.top === 0 && rect.bottom === 0) continue;
    if (rect.bottom < frame.top || rect.top > frame.bottom) continue;

    const line = document.createElement('i');
    line.style.left = `${frame.left}px`;
    line.style.width = `${frame.width}px`;
    line.style.top = `${rect.top + rect.height / 2}px`;
    layer.append(line);
    drawn++;
  }

  return drawn;
}
