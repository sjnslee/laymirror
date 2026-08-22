// what word's draft view draws for a manual page break: a dotted rule
// labelled "Page Break", across the column, above the block it precedes.
//
// drawn as an overlay rather than written into the document. the marks are
// positioned over the editor, never inserted into `.ProseMirror` — prosemirror
// owns its children and reconciles anything it did not put there out again.

import { blockForAnchor, editorBlocks } from '../host/anchors.js';
import type { PageBreak } from '../docx/breaks.js';

const LAYER_ID = 'laymirror-break-marks';
const STYLE_ID = 'laymirror-break-style';

const CSS = `
#${LAYER_ID} {
  position: absolute;
  inset: 0;
  pointer-events: none;
  z-index: 5;
}
#${LAYER_ID} .lm-break {
  position: absolute;
  display: flex;
  align-items: center;
  gap: 6px;
  font: 10px/1 system-ui, sans-serif;
  letter-spacing: .06em;
  text-transform: uppercase;
  color: var(--pmd-c-muted, #8a8f98);
  white-space: nowrap;
}
#${LAYER_ID} .lm-break::before,
#${LAYER_ID} .lm-break::after {
  content: '';
  flex: 1;
  border-top: 1px dotted currentColor;
}
`;

function ensureStyle(): void {
  if (document.getElementById(STYLE_ID)) return;
  const sheet = document.createElement('style');
  sheet.id = STYLE_ID;
  sheet.textContent = CSS;
  document.head.appendChild(sheet);
}

export const breakMarksShown = (): boolean => document.getElementById(LAYER_ID) !== null;

export function clearBreakMarks(): void {
  document.getElementById(LAYER_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();
}

/** draw a rule above each break's block. returns how many landed — a break
 *  whose anchor has gone is simply not drawn. */
export function showBreakMarks(editor: HTMLElement, breaks: readonly PageBreak[]): number {
  clearBreakMarks();
  ensureStyle();

  const host = editor.offsetParent instanceof HTMLElement ? editor.offsetParent : editor;
  if (getComputedStyle(host).position === 'static') host.style.position = 'relative';

  const layer = document.createElement('div');
  layer.id = LAYER_ID;

  const blocks = editorBlocks(editor);
  const hostBox = host.getBoundingClientRect();
  let drawn = 0;

  for (const mark of breaks) {
    const block = blockForAnchor(blocks, mark);
    if (!block) continue;
    const box = block.getBoundingClientRect();

    const rule = document.createElement('div');
    rule.className = 'lm-break';
    rule.textContent = 'Page Break';
    rule.style.top = `${box.top - hostBox.top - 7}px`;
    rule.style.left = `${box.left - hostBox.left}px`;
    rule.style.width = `${box.width}px`;
    layer.append(rule);
    drawn++;
  }

  host.append(layer);
  return drawn;
}
