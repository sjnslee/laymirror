import { describe, expect, it } from 'vitest';
import {
  BREAK_PARAGRAPH,
  headingIdsIn,
  injectBreaks,
  paragraphSpans,
  pruneBreaks,
  resolveBreak,
} from '../src/docx/breaks.js';

/** a heading paragraph as cardmirror's exporter writes one: the bookmark is
 *  INSIDE the `w:p`, wrapping the inlines (src/export/exporter.ts:443). */
const heading = (id: string, text: string) =>
  '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>' +
  `<w:bookmarkStart w:id="0" w:name="pmd-heading-${id}"/>` +
  `<w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`;

const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const doc = (...body: string[]) =>
  `<w:document><w:body>${body.join('')}<w:sectPr/></w:body></w:document>`;

const sample = doc(
  para('intro'),
  heading('aaa', 'first hat'),
  para('one'),
  para('two'),
  heading('bbb', 'second hat'),
  para('three'),
);

const countBreaks = (xml: string) => [...xml.matchAll(/w:type="page"/g)].length;

describe('paragraphSpans', () => {
  it('finds every paragraph in order', () => {
    expect(paragraphSpans(sample)).toHaveLength(6);
  });

  it('handles a self-closing empty paragraph', () => {
    expect(paragraphSpans(doc(para('a'), '<w:p/>', para('b')))).toHaveLength(3);
  });
});

describe('resolveBreak', () => {
  const spans = paragraphSpans(sample);

  it('offset 0 is the heading itself', () => {
    expect(resolveBreak(spans, { headingId: 'aaa', offset: 0 })).toBe(1);
  });

  it('counts blocks forward from the anchor', () => {
    expect(resolveBreak(spans, { headingId: 'aaa', offset: 2 })).toBe(3);
  });

  it('is null when the anchor is gone', () => {
    expect(resolveBreak(spans, { headingId: 'zzz', offset: 0 })).toBeNull();
  });

  it('is null when the offset runs past the end', () => {
    expect(resolveBreak(spans, { headingId: 'bbb', offset: 99 })).toBeNull();
  });

  it('rejects a nonsense offset rather than throwing', () => {
    expect(resolveBreak(spans, { headingId: 'aaa', offset: -1 })).toBeNull();
  });
});

describe('injectBreaks', () => {
  it('puts a real page break before the anchored paragraph', () => {
    const out = injectBreaks(sample, [{ headingId: 'bbb', offset: 0 }]);
    expect(countBreaks(out)).toBe(1);
    // it lands immediately before the heading it anchors to
    expect(out).toContain(`${BREAK_PARAGRAPH}<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>` +
      '<w:bookmarkStart w:id="0" w:name="pmd-heading-bbb"/>');
  });

  it('keeps several breaks in the right places', () => {
    const out = injectBreaks(sample, [
      { headingId: 'aaa', offset: 2 },
      { headingId: 'bbb', offset: 0 },
    ]);
    expect(countBreaks(out)).toBe(2);
    expect(out.indexOf('page')).toBeLessThan(out.lastIndexOf('page'));
  });

  it('collapses two marks that resolve to the same paragraph', () => {
    const out = injectBreaks(sample, [
      { headingId: 'aaa', offset: 3 },
      { headingId: 'bbb', offset: 0 },
    ]);
    expect(countBreaks(out)).toBe(1);
  });

  it('drops a mark whose anchor no longer exists', () => {
    expect(countBreaks(injectBreaks(sample, [{ headingId: 'gone', offset: 1 }]))).toBe(0);
  });

  // a break before paragraph zero only buys an empty leading page
  it('refuses to break before the first paragraph', () => {
    const first = doc(heading('aaa', 'hat'), para('body'));
    expect(countBreaks(injectBreaks(first, [{ headingId: 'aaa', offset: 0 }]))).toBe(0);
  });

  it('leaves the document alone when there are no breaks', () => {
    expect(injectBreaks(sample, [])).toBe(sample);
  });

  it('does not disturb the section properties', () => {
    const out = injectBreaks(sample, [{ headingId: 'bbb', offset: 0 }]);
    expect(out).toContain('<w:sectPr/></w:body>');
  });
});

describe('pruneBreaks', () => {
  it('keeps live anchors and forgets dead ones', () => {
    const kept = pruneBreaks(
      [
        { headingId: 'aaa', offset: 1 },
        { headingId: 'dead', offset: 1 },
      ],
      sample,
    );
    expect(kept).toEqual([{ headingId: 'aaa', offset: 1 }]);
  });
});

describe('headingIdsIn', () => {
  it('reads the ids the document carries', () => {
    expect(headingIdsIn(sample)).toEqual(new Set(['aaa', 'bbb']));
  });
});
