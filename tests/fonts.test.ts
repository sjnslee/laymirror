// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { isFontInstalled, missingFonts } from '../src/ui/fonts.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';

describe('font detection', () => {
  it('never cries wolf when it cannot measure', () => {
    // jsdom has no canvas, so measurement is unavailable — the honest answer
    // is "assume present" rather than warn about every font
    expect(isFontInstalled('Palatino Linotype')).toBe(true);
    expect(missingFonts(DEFAULT_LAY)).toEqual([]);
  });
});
