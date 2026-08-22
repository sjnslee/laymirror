// where to break when word has not already told us.
//
// this only runs when the file carries neither word's own
// `w:lastRenderedPageBreak` markers nor a manual break — that is, when
// cardmirror has just saved and stripped them. it observes heights measured
// from real rendered layout rather than predicting them from a model, which
// is the difference between this and the paginator it replaces.

/** a block that will not fit on the current page starts the next one.
 *
 *  a block taller than a whole page gets a page to itself and is allowed to
 *  run over. docx-preview sizes a page with `min-height`, so an overlong page
 *  grows rather than clipping — the page looks wrong, but no text is ever
 *  lost, and losing text is the one outcome worth ruling out. */
export function fillPages(heights: readonly number[], pageHeight: number): number[] {
  if (pageHeight <= 0) return [];

  const starts: number[] = [];
  let used = 0;

  for (let i = 0; i < heights.length; i++) {
    const height = Math.max(0, heights[i] ?? 0);

    if (used > 0 && used + height > pageHeight) {
      starts.push(i);
      used = height;
      continue;
    }
    used += height;
  }

  return starts;
}
