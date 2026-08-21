// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import {
  isFontInstalled,
  missingFonts,
  stackFor,
  substituteFont,
  SUBSTITUTES,
} from '../src/ui/fonts.js';
import { toCss } from '../src/render/css.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { stubCanvas } from './dom.js';

afterEach(() => vi.restoreAllMocks());

describe('font detection', () => {
  it('never cries wolf when it cannot measure', () => {
    // jsdom has no canvas, so measurement is unavailable — the honest answer
    // is "assume present" rather than warn about every font
    expect(isFontInstalled('Palatino Linotype')).toBe(true);
    expect(missingFonts(DEFAULT_LAY)).toEqual([]);
  });

  it('names the families the machine does not have', () => {
    stubCanvas(['Times New Roman']);
    expect(isFontInstalled('Times New Roman')).toBe(false);
    expect(isFontInstalled('Georgia')).toBe(true);
    expect(missingFonts(DEFAULT_LAY)).toEqual(['Times New Roman']);
  });
});

describe('substitution', () => {
  it('falls back to the family itself when nothing is chosen', () => {
    expect(stackFor({ ...DEFAULT_LAY, fontFallbacks: {} }, 'Garamond')).toBe('"Garamond", serif');
  });

  it('redirects one family and leaves the others alone', () => {
    const next = substituteFont(DEFAULT_LAY, 'Garamond', 'Tinos, serif');

    expect(stackFor(next, 'Garamond')).toBe('Tinos, serif');
    expect(stackFor(next, 'Cambria')).toBe(stackFor(DEFAULT_LAY, 'Cambria'));
    // the original is untouched
    expect(stackFor(DEFAULT_LAY, 'Garamond')).not.toBe('Tinos, serif');
  });

  it('reaches the stylesheet, because both read the same profile', () => {
    const substitute = SUBSTITUTES[0]!.stack;
    const next = substituteFont(DEFAULT_LAY, 'Times New Roman', substitute);
    expect(toCss(next)).toContain(substitute);
  });
});
