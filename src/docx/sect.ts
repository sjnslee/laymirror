// the section properties: page size, margins, and the header/footer
// references. cardmirror emits one hardcoded sectPr — letter with 1" margins
// and no header or footer at all — so this is entirely ours.

import type { PageSetup } from '../profile/profile.js';

export interface SectRefs {
  headerRelId?: string | null;
  footerRelId?: string | null;
}

export function buildSectPr(page: PageSetup, refs: SectRefs = {}): string {
  const m = page.margin;
  const parts: string[] = [];

  if (refs.headerRelId) {
    parts.push(`<w:headerReference w:type="default" r:id="${refs.headerRelId}"/>`);
  }
  if (refs.footerRelId) {
    parts.push(`<w:footerReference w:type="default" r:id="${refs.footerRelId}"/>`);
  }

  parts.push(`<w:pgSz w:w="${page.widthTwips}" w:h="${page.heightTwips}"/>`);
  parts.push(
    `<w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}"` +
      ` w:left="${m.left}" w:header="${m.header}" w:footer="${m.footer}" w:gutter="0"/>`,
  );
  parts.push('<w:cols w:space="720"/>');
  parts.push('<w:docGrid w:linePitch="360"/>');

  return `<w:sectPr>${parts.join('')}</w:sectPr>`;
}

/** the body's final sectPr is the document's; replace it wholesale rather
 *  than patching, since cardmirror's is always the same hardcoded block. */
export function replaceSectPr(documentXml: string, sectPr: string): string {
  if (/<w:sectPr\b/.test(documentXml)) {
    return documentXml.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/, sectPr);
  }
  return documentXml.replace('</w:body>', `${sectPr}</w:body>`);
}
