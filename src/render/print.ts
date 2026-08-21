// printing is printing the page view. the boxes are already the size of the
// target page, so `@page` only has to agree with them and take its own
// margin to zero — the browser then re-paginates nothing, which is the whole
// reason page view lays the content out itself.

import { CSS_DPI } from './paginate.js';
import type { PageSetup } from '../profile/profile.js';

const twipsToIn = (twips: number): string => `${(twips / 1440).toFixed(4)}in`;

/** the print half of the page-view stylesheet. */
export function printStyles(page: PageSetup, pageSelector: string, rootId: string): string {
  return [
    `@page {`,
    `  size: ${twipsToIn(page.widthTwips)} ${twipsToIn(page.heightTwips)};`,
    // the page boxes carry the margins themselves
    `  margin: 0;`,
    `}`,
    `@media print {`,
    `  body > *:not(#${rootId}) { display: none !important; }`,
    `  #${rootId} {`,
    '    position: static !important;',
    '    overflow: visible !important;',
    '    background: #fff !important;',
    '    padding: 0 !important;',
    '  }',
    `  #${rootId} .lm-chrome { display: none !important; }`,
    `  ${pageSelector} {`,
    '    box-shadow: none !important;',
    '    margin: 0 !important;',
    // one box per sheet, and never a stray blank last sheet
    '    break-after: page;',
    '  }',
    `  ${pageSelector}:last-child { break-after: auto; }`,
    '}',
  ].join('\n');
}

export function printPageView(): void {
  window.print();
}

/** css pixels for a twip measurement, at the fixed browser dpi. */
export const twipsToPx = (twips: number): number => (twips / 1440) * CSS_DPI;
