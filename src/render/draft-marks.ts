// word's dotted page-break rule, drawn over work view.
//
// the same `paginate` computation page view uses, so the two always agree
// about where page 2 starts. the marks are an overlay beside the editor
// content, never nodes inside it: prosemirror owns those children and would
// either reconcile ours away or mistake them for document content.

import { EDITOR_SELECTOR } from '../host/cardmirror.js';
import { buildFlow, flowBlocks, measureBlocks } from './measure.js';
import { paginate } from './paginate.js';
import type { BreakMark } from './paginate.js';
import type { Profile } from '../profile/profile.js';

const LAYER_ID = 'laymirror-draft-marks';
const STYLE_ID = 'laymirror-draft-style';
/** an edit relaid out on every keystroke would measure more than it draws. */
const SETTLE_MS = 400;

let observer: MutationObserver | null = null;
let timer: ReturnType<typeof setTimeout> | null = null;
let profileInForce: Profile | null = null;

const CSS = `
#${LAYER_ID} {
  position: absolute;
  left: 0;
  right: 0;
  top: 0;
  height: 0;
  pointer-events: none;
  z-index: 5;
}
#${LAYER_ID} .lm-mark {
  position: absolute;
  left: 0;
  right: 0;
  border-top: 1px dotted #888;
  font: 10px/1 system-ui, sans-serif;
  color: #888;
}
#${LAYER_ID} .lm-mark span {
  position: absolute;
  right: 4px;
  top: -6px;
  padding: 0 4px;
  background: var(--pmd-c-surface, #fff);
}
#${LAYER_ID} .lm-mark.lm-manual { border-top-style: dashed; border-top-color: #4a7; }
#${LAYER_ID} .lm-mark.lm-manual span { color: #4a7; }
`;

function content(): HTMLElement | null {
  return (
    document.querySelector<HTMLElement>('.ProseMirror') ??
    document.querySelector<HTMLElement>(EDITOR_SELECTOR)
  );
}

function ensureStyle(): void {
  if (document.getElementById(STYLE_ID)) return;
  const sheet = document.createElement('style');
  sheet.id = STYLE_ID;
  sheet.textContent = CSS;
  document.head.appendChild(sheet);
}

/** measured at the width of the printed column, not the editor's, or the
 *  line counts belong to a layout nobody is going to print. */
function measureOffscreen(blocks: readonly HTMLElement[], profile: Profile) {
  const stage = document.createElement('div');
  Object.assign(stage.style, {
    position: 'absolute',
    left: '-10000px',
    top: '0',
    visibility: 'hidden',
  });

  const flow = buildFlow(blocks, profile);
  stage.append(flow);
  document.body.append(stage);
  try {
    return measureBlocks(flowBlocks(flow), profile);
  } finally {
    stage.remove();
  }
}

/** where a break lands in the live editor. a break inside a block is placed
 *  proportionally: the editor is a different width, so its lines are not the
 *  page's lines and there is no exact answer to give. */
function offsetOf(mark: BreakMark, blocks: readonly HTMLElement[], stagedHeight: number, origin: number, scroll: number): number | null {
  const block = blocks[mark.block];
  if (!block) return null;

  const rect = block.getBoundingClientRect();
  const top = rect.top - origin + scroll;
  if (mark.offsetPx <= 0 || stagedHeight <= 0) return top;
  return top + (mark.offsetPx / stagedHeight) * rect.height;
}

export function clearDraftMarks(): void {
  if (timer !== null) clearTimeout(timer);
  timer = null;
  observer?.disconnect();
  observer = null;
  profileInForce = null;
  document.getElementById(LAYER_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();
}

export function draftMarksShown(): boolean {
  return document.getElementById(LAYER_ID) !== null;
}

/** returns how many breaks were drawn, or null when there is nothing to
 *  measure. */
export function drawDraftMarks(profile: Profile): number | null {
  const editor = content();
  const host = editor?.parentElement ?? editor;
  if (!editor || !host) return null;

  const blocks = Array.from(editor.children).filter(
    (child): child is HTMLElement => child instanceof HTMLElement,
  );
  if (blocks.length === 0) return null;

  ensureStyle();
  if (getComputedStyle(host).position === 'static') host.style.position = 'relative';

  const metrics = measureOffscreen(blocks, profile);
  const { breaks } = paginate(metrics, profile.page);

  const layer = document.getElementById(LAYER_ID) ?? document.createElement('div');
  layer.id = LAYER_ID;
  layer.replaceChildren();
  host.append(layer);

  const origin = host.getBoundingClientRect().top;
  breaks.forEach((mark, index) => {
    const y = offsetOf(mark, blocks, metrics[mark.block]?.heightPx ?? 0, origin, host.scrollTop);
    if (y === null) return;

    const rule = document.createElement('div');
    rule.className = mark.manual ? 'lm-mark lm-manual' : 'lm-mark';
    rule.style.top = `${y}px`;

    const label = document.createElement('span');
    label.textContent = mark.manual ? 'page break' : `page ${index + 2}`;
    rule.append(label);
    layer.append(rule);
  });

  return breaks.length;
}

/** draw, and keep drawing as the document changes. */
export function showDraftMarks(profile: Profile): number | null {
  profileInForce = profile;
  const drawn = drawDraftMarks(profile);
  if (drawn === null) return null;

  const editor = content();
  if (editor && !observer) {
    observer = new MutationObserver(() => {
      if (timer !== null) clearTimeout(timer);
      timer = setTimeout(() => {
        if (profileInForce) drawDraftMarks(profileInForce);
      }, SETTLE_MS);
    });
    observer.observe(editor, { childList: true, subtree: true, characterData: true });
  }

  return drawn;
}
