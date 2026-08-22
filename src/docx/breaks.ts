// manual page breaks, held outside the document and injected on save.
//
// cardmirror cannot carry a page break. its importer turns `<w:br w:type=
// "page"/>` into a plain newline ("no hard-page-break support",
// src/import/importer.ts:771) and its exporter writes it back as a line break
// ("the doc model keeps the break but not its type", src/export/exporter.ts:611).
// `<w:pageBreakBefore/>` is dropped by both. so nothing put *in* the document
// survives a round-trip, and the old `[page break]` text sentinel was a
// visible, corruptible stand-in for a thing the model simply cannot hold.
//
// what cardmirror does keep is a stable heading id: pocket / hat / block /
// tag / analytic each carry a uuid that survives edits and round-trips as a
// `pmd-heading-<uuid>` bookmark inside the heading's own paragraph
// (src/schema/ids.ts). that is the anchor.

const BOOKMARK_PREFIX = 'pmd-heading-';

/** before the block `offset` positions after the heading owning `headingId`.
 *  `offset: 0` is before that heading's own paragraph. */
export interface PageBreak {
  headingId: string;
  offset: number;
}

/** what word's ctrl+enter produces, and the only form docx-preview honours —
 *  it reads `pageBreakBefore` from a paragraph's style and ignores a direct
 *  one in `pPr`. */
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
    spans.push({
      start: match.index,
      end: match.index + match[0].length,
      xml: match[0],
    });
  }
  return spans;
}

const anchorPattern = (headingId: string): RegExp =>
  new RegExp(`w:name="${BOOKMARK_PREFIX}${headingId.replace(/[^\w-]/g, '')}"`);

/** the paragraph index a break lands before, or null when its anchor is gone.
 *  a vanished anchor is dropped rather than guessed at — a page break in the
 *  wrong place is worse than one that quietly stopped applying. */
export function resolveBreak(spans: Span[], mark: PageBreak): number | null {
  if (!Number.isInteger(mark.offset) || mark.offset < 0) return null;
  const pattern = anchorPattern(mark.headingId);
  const anchor = spans.findIndex((span) => pattern.test(span.xml));
  if (anchor === -1) return null;
  const target = anchor + mark.offset;
  return target < spans.length ? target : null;
}

/** insert a real page break before each resolved paragraph.
 *
 *  insertions run back to front so earlier offsets stay valid, and a target
 *  reached twice takes one break, not two. */
export function injectBreaks(documentXml: string, breaks: readonly PageBreak[]): string {
  if (breaks.length === 0) return documentXml;

  const spans = paragraphSpans(documentXml);
  const targets = new Set<number>();
  for (const mark of breaks) {
    const index = resolveBreak(spans, mark);
    // a break before the first paragraph only buys an empty leading page
    if (index !== null && index > 0) targets.add(index);
  }

  let xml = documentXml;
  for (const index of [...targets].sort((a, b) => b - a)) {
    const at = spans[index]!.start;
    xml = xml.slice(0, at) + BREAK_PARAGRAPH + xml.slice(at);
  }
  return xml;
}

/** the ids the document actually carries, so stale marks can be pruned. */
export function headingIdsIn(documentXml: string): Set<string> {
  const ids = new Set<string>();
  for (const match of documentXml.matchAll(
    new RegExp(`w:name="${BOOKMARK_PREFIX}([\\w-]+)"`, 'g'),
  )) {
    ids.add(match[1]!);
  }
  return ids;
}

export function pruneBreaks(
  breaks: readonly PageBreak[],
  documentXml: string,
): PageBreak[] {
  const live = headingIdsIn(documentXml);
  return breaks.filter((mark) => live.has(mark.headingId));
}

/** insert breaks before paragraphs chosen by index rather than by anchor —
 *  the fallback path, where a fill pass over real rendered layout has decided
 *  where the pages end. */
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
