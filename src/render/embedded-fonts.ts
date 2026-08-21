// fonts the plugin carries itself.
//
// cardmirror bundles metric-compatible faces for arial, calibri, cambria,
// times new roman and georgia, so a template asking for those is covered.
// garamond it does not, and neither does macos — a lay template calling for
// garamond fell back to whatever the browser called a serif, which is what
// "the fonts are wrong" looked like.
//
// eb garamond is a genuine garamond revival under the open font license. it
// is not metric-compatible with monotype's garamond (nothing free is), so
// line breaks can still differ from word's — but it is the right typeface
// rather than a stand-in for one.

import regular from '@fontsource/eb-garamond/files/eb-garamond-latin-400-normal.woff2';
import bold from '@fontsource/eb-garamond/files/eb-garamond-latin-700-normal.woff2';
import italic from '@fontsource/eb-garamond/files/eb-garamond-latin-400-italic.woff2';

const STYLE_ID = 'laymirror-fonts';

const face = (weight: number, style: 'normal' | 'italic', src: string): string =>
  `@font-face {
  font-family: "EB Garamond";
  font-style: ${style};
  font-weight: ${weight};
  font-display: swap;
  src: url(${src}) format("woff2");
}`;

export const EMBEDDED_FACES = [
  face(400, 'normal', regular),
  face(700, 'normal', bold),
  face(400, 'italic', italic),
].join('\n');

/** the families the plugin can supply whatever the machine has. */
export const EMBEDDED_FAMILIES = ['EB Garamond'];

export function ensureEmbeddedFonts(): void {
  if (document.getElementById(STYLE_ID)) return;
  const sheet = document.createElement('style');
  sheet.id = STYLE_ID;
  sheet.textContent = EMBEDDED_FACES;
  document.head.appendChild(sheet);
}

export function removeEmbeddedFonts(): void {
  document.getElementById(STYLE_ID)?.remove();
}

export function embeddedFontsPresent(): boolean {
  return document.getElementById(STYLE_ID) !== null;
}
