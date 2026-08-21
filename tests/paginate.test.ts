import { describe, it, expect } from 'vitest';
import {
  paginate,
  usableHeightPx,
  usableWidthPx,
  type BlockMetrics,
  type PaginateResult,
} from '../src/render/paginate.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';

// letter, one inch all round: a 9 x 6.5 inch text column, 864 x 624 css px
const page = DEFAULT_LAY.page;
const PAGE_PX = 864;

const LINE = 24;

/** a block of n lines, splittable unless told otherwise. */
const block = (lines: number, over: Partial<BlockMetrics> = {}): BlockMetrics => ({
  heightPx: lines * LINE,
  lineTopsPx: Array.from({ length: lines }, (_, i) => i * LINE),
  breakBefore: false,
  keepNext: false,
  keepLines: false,
  manualBreak: false,
  ...over,
});

const heights = (result: PaginateResult): number[] => result.pages.map((p) => p.heightPx);

/** every block covered exactly once, in order, with nothing lost or repeated. */
function expectWholeDocument(result: PaginateResult, blocks: readonly BlockMetrics[]): void {
  const covered = new Map<number, number>();
  let previous = -1;

  for (const page of result.pages) {
    for (const slice of page.slices) {
      expect(slice.block).toBeGreaterThanOrEqual(previous);
      previous = slice.block;
      const at = covered.get(slice.block) ?? 0;
      // a continuation must start exactly where the last slice stopped
      expect(slice.fromPx).toBe(at);
      expect(slice.toPx).toBeGreaterThanOrEqual(slice.fromPx);
      covered.set(slice.block, slice.toPx);
    }
  }

  blocks.forEach((b, i) => {
    expect(covered.get(i)).toBe(b.heightPx);
  });
}

describe('the text column', () => {
  it('is the page less its margins', () => {
    expect(usableHeightPx(page)).toBe(PAGE_PX);
    expect(usableWidthPx(page)).toBe(624);
  });
});

describe('paginate', () => {
  it('keeps a short document on one page', () => {
    const blocks = [block(4), block(6), block(2)];
    const result = paginate(blocks, page);

    expect(result.pages).toHaveLength(1);
    expect(result.breaks).toHaveLength(0);
    expectWholeDocument(result, blocks);
  });

  it('starts a second page when the first is full', () => {
    // 36 lines of 24px is 864px exactly, so the next block cannot fit
    const blocks = [block(36, { keepLines: true }), block(4, { keepLines: true })];
    const result = paginate(blocks, page);

    expect(heights(result)).toEqual([864, 96]);
    expectWholeDocument(result, blocks);
  });

  it('breaks before a block that asks to start a page', () => {
    const blocks = [block(4), block(4, { breakBefore: true })];
    const result = paginate(blocks, page);

    expect(result.pages).toHaveLength(2);
    expect(result.breaks).toEqual([{ block: 1, offsetPx: 0, manual: false }]);
  });

  it('does not open a page break on the very first block', () => {
    const result = paginate([block(4, { breakBefore: true }), block(4)], page);
    expect(result.pages).toHaveLength(1);
    expect(result.breaks).toHaveLength(0);
  });

  it('reports a manual break as manual', () => {
    const blocks = [block(4), block(1, { manualBreak: true })];
    const result = paginate(blocks, page);

    expect(result.breaks).toEqual([{ block: 1, offsetPx: 0, manual: true }]);
    expectWholeDocument(result, blocks);
  });

  it('moves a keepNext run whole rather than splitting it', () => {
    // a tag that would land at the very bottom takes its card with it
    const blocks = [block(34), block(2, { keepNext: true }), block(6)];
    const result = paginate(blocks, page);

    expect(result.pages[0]!.slices.map((s) => s.block)).toEqual([0]);
    expect(result.pages[1]!.slices.map((s) => s.block)).toEqual([1, 2]);
    expectWholeDocument(result, blocks);
  });

  it('moves a keepLines block whole rather than splitting it', () => {
    const blocks = [block(30), block(10, { keepLines: true })];
    const result = paginate(blocks, page);

    expect(result.pages[0]!.slices).toHaveLength(1);
    expect(result.pages[1]!.slices[0]).toEqual({ block: 1, fromPx: 0, toPx: 240 });
  });

  it('splits a long block at a line boundary', () => {
    // 30 lines then a 20 line card: 6 lines fit, 14 carry over
    const blocks = [block(30), block(20)];
    const result = paginate(blocks, page);

    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]!.slices[1]).toEqual({ block: 1, fromPx: 0, toPx: 6 * LINE });
    expect(result.pages[1]!.slices[0]).toEqual({ block: 1, fromPx: 6 * LINE, toPx: 20 * LINE });
    expect(result.breaks).toEqual([{ block: 1, offsetPx: 6 * LINE, manual: false }]);
    expectWholeDocument(result, blocks);
  });

  it('will not leave a widow alone at the top of a page', () => {
    // 35 lines then 2: cutting after the first would strand one line
    const blocks = [block(35), block(2)];
    const result = paginate(blocks, page, { widowLines: 2, orphanLines: 1 });

    expect(result.pages[0]!.slices.map((s) => s.block)).toEqual([0]);
    expect(result.pages[1]!.slices[0]).toEqual({ block: 1, fromPx: 0, toPx: 2 * LINE });
  });

  it('will not leave an orphan alone at the bottom of one', () => {
    // exactly one line of the card fits, which is fewer than the two required
    const blocks = [block(35), block(8)];
    const result = paginate(blocks, page, { widowLines: 1, orphanLines: 2 });

    expect(result.pages[0]!.slices.map((s) => s.block)).toEqual([0]);
    expect(result.pages[1]!.slices[0]!.fromPx).toBe(0);
    expectWholeDocument(result, blocks);
  });

  it('gets a block taller than a page onto pages anyway', () => {
    const blocks = [block(100, { keepLines: true })];
    const result = paginate(blocks, page);

    expect(result.pages.length).toBeGreaterThan(2);
    expectWholeDocument(result, blocks);
  });

  it('always returns a page, even for nothing at all', () => {
    expect(paginate([], page).pages).toHaveLength(1);
  });

  it('loses nothing across a long mixed document', () => {
    const blocks: BlockMetrics[] = [];
    for (let i = 0; i < 60; i++) {
      blocks.push(
        block((i % 7) + 1, {
          keepNext: i % 5 === 0,
          keepLines: i % 11 === 0,
          breakBefore: i % 17 === 0,
          manualBreak: i % 23 === 0,
        }),
      );
    }
    expectWholeDocument(paginate(blocks, page), blocks);
  });
});
