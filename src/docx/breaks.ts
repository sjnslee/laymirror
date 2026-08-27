// page breaks, as far as a rendered preview is concerned.
//
// laymirror does not put breaks into a document any more. the template's own
// styles carry `w:pageBreakBefore` — a lay template breaks before every
// heading 1 — and that is a property of the style, so it survives a cardmirror
// round-trip untouched and word and docx-preview both honour it.
//
// what is left here is only what page view needs: reading whether a file
// already says where its pages end, and inserting breaks into a throwaway copy
// when it does not.

/** what word's ctrl+enter produces, and the only form docx-preview honours in
 *  the body — it reads `pageBreakBefore` from a paragraph's style and ignores
 *  a direct one in `pPr`. */
export const BREAK_PARAGRAPH = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

export interface Span {
  start: number;
  end: number;
  xml: string;
}

/** every `<w:p>` in document order. `w:p` cannot nest, so a flat scan is the
 *  whole story — paragraphs inside table cells are included, in place. */
export function paragraphSpans(documentXml: string): Span[] {
  const spans: Span[] = [];
  for (const match of documentXml.matchAll(/<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)) {
    spans.push({ start: match.index, end: match.index + match[0].length, xml: match[0] });
  }
  return spans;
}

/** insert breaks before paragraphs chosen by index — the fallback path, where
 *  a fill pass over real rendered layout has decided where the pages end. */
export function injectBreaksAt(documentXml: string, indices: readonly number[]): string {
  if (indices.length === 0) return documentXml;
  const spans = paragraphSpans(documentXml);
  const targets = [...new Set(indices)].filter((i) => i > 0 && i < spans.length);

  let xml = documentXml;
  for (const index of targets.sort((a, b) => b - a)) {
    const at = spans[index]!.start;
    xml = xml.slice(0, at) + BREAK_PARAGRAPH + xml.slice(at);
  }
  return xml;
}

/** does the document already say where its pages end? word writes
 *  `lastRenderedPageBreak` on every save, so a file it touched carries its own
 *  pagination and needs no guessing at all. */
export function hasRenderedBreaks(documentXml: string): boolean {
  return documentXml.includes('lastRenderedPageBreak');
}

export function hasManualBreaks(documentXml: string): boolean {
  return /<w:br\b[^>]*w:type="page"/.test(documentXml);
}
