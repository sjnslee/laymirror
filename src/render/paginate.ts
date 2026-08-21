// measure and split. the pure half of page view, and the same computation
// draft marks read.
//
// written from the ooxml spec and word's documented pagination behaviour.
// deliberately not derived from any proprietary layout engine.
//
// the splitting is by line, not by block: a card body runs to half a page
// and refusing to break it would push whitespace onto every page word fills.
// so a page holds slices — a vertical window into a block — and widow and
// orphan control is the rule about where those windows may fall.

import type { PageSetup } from '../profile/profile.js';

/** css pixels per inch. fixed in every browser regardless of the display. */
export const CSS_DPI = 96;

export interface BlockMetrics {
  /** the block's extent in the flow, from its own top to the next block's,
   *  so collapsed margins are already accounted for by the measurement
   *  rather than re-derived here. */
  heightPx: number;
  /** the top of each line within the block. one entry means it is a single
   *  line and cannot be split; empty means it must not be. */
  lineTopsPx: readonly number[];
  breakBefore: boolean;
  keepNext: boolean;
  keepLines: boolean;
  /** our own page break, which cardmirror's model cannot carry. */
  manualBreak: boolean;
}

/** a vertical window into one block. `fromPx` is 0 for a block that starts
 *  on this page; a non-zero value is a continuation. */
export interface Slice {
  block: number;
  fromPx: number;
  toPx: number;
}

export interface PageBox {
  slices: Slice[];
  heightPx: number;
}

/** where a page break falls, for the dotted rule work view draws.
 *  `offsetPx` is 0 when the break is between blocks. */
export interface BreakMark {
  block: number;
  offsetPx: number;
  manual: boolean;
}

export interface PaginateResult {
  pages: PageBox[];
  breaks: BreakMark[];
}

export interface PaginateOptions {
  /** lines of a split block that may not be left alone at the top of a page. */
  widowLines?: number;
  /** lines that may not be left alone at the bottom of one. */
  orphanLines?: number;
  dpi?: number;
}

const twipsToPx = (twips: number, dpi: number): number => (twips / 1440) * dpi;

/** the text column: word's body area is the page less its top and bottom
 *  margins, with the header and footer living inside those margins. */
export function usableHeightPx(page: PageSetup, dpi = CSS_DPI): number {
  return twipsToPx(page.heightTwips - page.margin.top - page.margin.bottom, dpi);
}

export function usableWidthPx(page: PageSetup, dpi = CSS_DPI): number {
  return twipsToPx(page.widthTwips - page.margin.left - page.margin.right, dpi);
}

/** blocks that must land on one page together: a run joined by keepNext. */
function atomsOf(blocks: readonly BlockMetrics[]): number[][] {
  const atoms: number[][] = [];
  let current: number[] = [];

  blocks.forEach((block, i) => {
    const starts = block.breakBefore || block.manualBreak;
    if (current.length > 0 && starts) {
      atoms.push(current);
      current = [];
    }
    current.push(i);
    if (!block.keepNext) {
      atoms.push(current);
      current = [];
    }
  });

  if (current.length > 0) atoms.push(current);
  return atoms;
}

/** the last line boundary at or before `limitPx`, as a line count. */
function linesThatFit(tops: readonly number[], limitPx: number): number {
  let fit = 0;
  for (let i = 1; i < tops.length; i++) {
    if (tops[i]! <= limitPx) fit = i;
    else break;
  }
  return fit;
}

export function paginate(
  blocks: readonly BlockMetrics[],
  page: PageSetup,
  opts: PaginateOptions = {},
): PaginateResult {
  const widowLines = Math.max(1, opts.widowLines ?? 2);
  const orphanLines = Math.max(1, opts.orphanLines ?? 2);
  const height = usableHeightPx(page, opts.dpi ?? CSS_DPI);

  const pages: PageBox[] = [];
  const breaks: BreakMark[] = [];
  let current: PageBox = { slices: [], heightPx: 0 };

  const remaining = (): number => height - current.heightPx;

  const flush = (at: BreakMark | null): void => {
    if (current.slices.length === 0) return;
    pages.push(current);
    current = { slices: [], heightPx: 0 };
    if (at) breaks.push(at);
  };

  const place = (block: number, fromPx: number, toPx: number): void => {
    current.slices.push({ block, fromPx, toPx });
    current.heightPx += toPx - fromPx;
  };

  for (const atom of atomsOf(blocks)) {
    const first = atom[0]!;
    if (blocks[first]!.breakBefore || blocks[first]!.manualBreak) {
      flush({ block: first, offsetPx: 0, manual: blocks[first]!.manualBreak });
    }

    const atomHeight = atom.reduce((sum, i) => sum + blocks[i]!.heightPx, 0);
    if (atomHeight <= remaining()) {
      for (const i of atom) place(i, 0, blocks[i]!.heightPx);
      continue;
    }

    // a group held together by keepNext is moved whole rather than split,
    // unless it could never fit on a page of its own
    if (atom.length > 1 && atomHeight <= height) {
      flush({ block: first, offsetPx: 0, manual: false });
      for (const i of atom) place(i, 0, blocks[i]!.heightPx);
      continue;
    }

    for (const i of atom) {
      const block = blocks[i]!;
      let from = 0;

      for (;;) {
        const left = remaining();
        const rest = block.heightPx - from;

        if (rest <= left) {
          place(i, from, block.heightPx);
          break;
        }

        const cut = splitAt(block, from, left, widowLines, orphanLines);
        if (cut !== null) {
          place(i, from, cut);
          flush({ block: i, offsetPx: cut, manual: false });
          from = cut;
          continue;
        }

        // no legal cut in the space left, so the remainder moves on whole
        if (current.slices.length > 0) {
          flush({ block: i, offsetPx: from, manual: false });
          continue;
        }

        // and it does not fit on a page of its own either — a taller-than-a-
        // page image, or a block held together by keepLines. take what the
        // page holds rather than looping on a cut that will never come.
        const take = left > 0 ? left : rest;
        place(i, from, from + take);
        from += take;
        if (from >= block.heightPx) break;
        flush({ block: i, offsetPx: from, manual: false });
      }
    }
  }

  flush(null);
  if (pages.length === 0) pages.push({ slices: [], heightPx: 0 });
  return { pages, breaks };
}

/** the offset a block may be cut at, or null when no legal cut exists in the
 *  space left. */
function splitAt(
  block: BlockMetrics,
  fromPx: number,
  availablePx: number,
  widowLines: number,
  orphanLines: number,
): number | null {
  if (block.keepLines) return null;

  const tops = block.lineTopsPx;
  if (tops.length < 2) return null;

  const startLine = tops.findIndex((top) => top >= fromPx);
  if (startLine === -1) return null;

  let line = linesThatFit(tops, fromPx + availablePx);
  // orphan control: leave at least this many lines behind
  while (line - startLine >= orphanLines) {
    // widow control: carry at least this many forward
    if (tops.length - line >= widowLines) return tops[line]!;
    line -= 1;
  }
  return null;
}
