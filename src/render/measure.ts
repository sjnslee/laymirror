// the dom half of pagination: turn rendered blocks into the numbers
// `paginate` works on. kept apart from the paginator so the rules are
// testable without a layout engine, and so this file stays the only place
// that touches real geometry.

import { blockTypeOf } from './css.js';
import { PAGE_BREAK_TEXT } from '../profile/mapping.js';
import type { BlockMetrics } from './paginate.js';
import type { Profile } from '../profile/profile.js';

export function isPageBreak(el: Element): boolean {
  return (el.textContent ?? '').trim() === PAGE_BREAK_TEXT;
}

/** the top of each line box within an element, relative to the element.
 *
 *  a range over the element's contents reports one rect per line, which is
 *  the browser's own line breaking rather than a guess at it. rects that
 *  share a top are fragments of the same line. */
function lineTopsPx(el: HTMLElement): number[] {
  const range = el.ownerDocument.createRange();
  range.selectNodeContents(el);

  const origin = el.getBoundingClientRect().top;
  const tops: number[] = [];
  for (const rect of Array.from(range.getClientRects())) {
    if (rect.height === 0) continue;
    const top = rect.top - origin;
    const last = tops[tops.length - 1];
    if (last === undefined || top - last > 1) tops.push(top);
  }
  return tops;
}

const numeric = (value: string): number => {
  const n = Number.parseFloat(value);
  return Number.isFinite(n) ? n : 0;
};

/** blocks must already be laid out at the width they will be paginated at,
 *  or every line count is wrong. */
export function measureBlocks(
  blocks: readonly HTMLElement[],
  profile: Profile,
): BlockMetrics[] {
  const rects = blocks.map((el) => el.getBoundingClientRect());

  return blocks.map((el, i) => {
    const spec = profile.types[blockTypeOf(el)];
    const next = rects[i + 1];
    // the distance to the next block already contains whatever margin
    // survived collapsing, which is not something worth recomputing
    const heightPx = next
      ? next.top - rects[i]!.top
      : rects[i]!.height + numeric(getComputedStyle(el).marginBottom);

    return {
      heightPx,
      lineTopsPx: lineTopsPx(el),
      breakBefore: spec.pageBreakBefore === true,
      keepNext: spec.keepNext === true,
      keepLines: spec.keepLines === true,
      manualBreak: isPageBreak(el),
    };
  });
}
