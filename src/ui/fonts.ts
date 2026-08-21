// metric probe: render a string in the candidate family with a known fallback
// and see whether the width moves. the same trick cardmirror's font-detect
// uses. cardmirror bundles substitutes for cambria, calibri and times new
// roman, but not palatino linotype or garamond — those are where our
// pagination drifts from word's.

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
