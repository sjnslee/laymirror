// metric probe: render a string in the candidate family with a known fallback
// and see whether the width moves. the same trick cardmirror's font-detect
// uses. cardmirror bundles substitutes for cambria, calibri and times new
// roman, but not palatino linotype or garamond — those are where our
// pagination drifts from word's.

import { fontStackFor } from '../render/css.js';
import type { Profile } from '../profile/profile.js';

const PROBE = 'mmmmmmmmmmlliWWWWWWWWWW';

function widthIn(family: string, fallback: string): number | null {
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  ctx.font = `72px "${family}", ${fallback}`;
  return ctx.measureText(PROBE).width;
}

/** installed when the family measures differently from at least one generic
 *  fallback — if it is absent, every fallback yields that fallback's width. */
export function isFontInstalled(family: string): boolean {
  const widths = ['monospace', 'serif', 'sans-serif'].map((f) => widthIn(family, f));
  if (widths.some((w) => w === null)) return true; // can't tell; don't cry wolf
  return new Set(widths).size === 1;
}

export function missingFonts(profile: Profile): string[] {
  const families = new Set<string>();
  for (const spec of Object.values(profile.types)) if (spec.font) families.add(spec.font);
  return [...families].filter((f) => !isFontInstalled(f));
}

export interface Substitute {
  label: string;
  stack: string;
}

/** the faces cardmirror bundles, so they are there whatever the machine has.
 *  the first three are metric-compatible with the family they name — same
 *  advance widths, so a page breaks where word breaks it. the rest are not,
 *  and are offered because a legible near-match beats a browser default. */
export const SUBSTITUTES: readonly Substitute[] = [
  { label: 'EB Garamond (carried by laymirror)', stack: '"EB Garamond", Garamond, serif' },
  { label: 'Tinos (matches Times New Roman)', stack: 'Tinos, "Times New Roman", serif' },
  { label: 'Caladea (matches Cambria)', stack: 'Caladea, Cambria, serif' },
  { label: 'Carlito (matches Calibri)', stack: 'Carlito, Calibri, sans-serif' },
  { label: 'Arimo (matches Arial)', stack: 'Arimo, Arial, sans-serif' },
  { label: 'Gelasio (matches Georgia)', stack: 'Gelasio, Georgia, serif' },
  { label: 'Noto Serif', stack: '"Noto Serif", serif' },
  { label: 'DejaVu Serif', stack: '"DejaVu Serif", serif' },
  { label: 'Noto Sans', stack: '"Noto Sans", sans-serif' },
  { label: 'Atkinson Hyperlegible', stack: '"Atkinson Hyperlegible", sans-serif' },
];

/** what a family will actually be drawn in, as the stylesheet decides it. */
export const stackFor = fontStackFor;

/** a copy of the profile with one family redirected. the substitution lives
 *  on the profile because `toCss` reads it there — the screen and the file
 *  stay one object. */
export function substituteFont(profile: Profile, family: string, stack: string): Profile {
  return { ...profile, fontFallbacks: { ...profile.fontFallbacks, [family]: stack } };
}
