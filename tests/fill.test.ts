import { describe, expect, it } from 'vitest';
import { fillPages } from '../src/render/fill.js';

describe('fillPages', () => {
  it('keeps everything on one page when it fits', () => {
    expect(fillPages([100, 100, 100], 500)).toEqual([]);
  });

  it('starts a new page at the block that would not fit', () => {
    expect(fillPages([100, 100, 100], 250)).toEqual([2]);
  });

  it('breaks repeatedly through a long document', () => {
    expect(fillPages(Array(9).fill(100), 300)).toEqual([3, 6]);
  });

  // clipping is the one outcome worth ruling out: docx-preview sizes a page
  // with min-height, so an overlong block makes the page grow, not hide text
  it('gives an oversized block its own page rather than dropping it', () => {
    expect(fillPages([100, 900, 100], 300)).toEqual([1, 2]);
  });

  it('never breaks before the first block', () => {
    expect(fillPages([900], 300)).toEqual([]);
  });

  it('is empty for an empty document', () => {
    expect(fillPages([], 300)).toEqual([]);
  });

  it('refuses a nonsense page height instead of looping', () => {
    expect(fillPages([100, 100], 0)).toEqual([]);
  });

  it('treats a missing or negative height as zero', () => {
    expect(fillPages([100, -50, 100], 250)).toEqual([]);
  });
});
