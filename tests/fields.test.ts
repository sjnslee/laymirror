// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { fillFields, findFields } from '../src/docx/fields.js';

const HDR = 'word/header1.xml';
const hdr = (body: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  body +
  '</w:hdr>';

const text = (xml: string): string =>
  [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('|');

/** what a real school header looks like: a code word split across runs by
 *  word's revision ids, a right tab, a title, and a live page field. */
const REAL = hdr(
  '<w:p><w:r><w:t>BCP </w:t></w:r><w:r><w:t>26</w:t></w:r><w:r><w:t>-</w:t></w:r>' +
    '<w:r><w:t>27</w:t></w:r>' +
    '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>' +
    '<w:r><w:t>File Title</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Name</w:t></w:r>' +
    '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>' +
    '<w:r><w:t xml:space="preserve"> Page </w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>1</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
);

describe('findFields', () => {
  it('reads a run of text split across runs as one field', () => {
    expect(findFields({ [HDR]: REAL }).map((f) => f.label)).toEqual([
      'BCP 26-27',
      'File Title',
      'Name',
    ]);
  });

  // " page " and " of " are a word field's own decoration, not the user's to
  // edit, and the result inside the field is recomputed on every open
  it('offers nothing that belongs to a word field', () => {
    const labels = findFields({ [HDR]: REAL }).map((f) => f.label);
    expect(labels).not.toContain('Page');
    expect(labels).not.toContain('1');
  });

  it('prefers a zero-width mark where the template author left one', () => {
    const marked = hdr(
      '<w:p><w:r><w:t>Property of ⁠Somebody⁠, do not copy</w:t></w:r></w:p>',
    );
    expect(findFields({ [HDR]: marked }).map((f) => f.label)).toEqual(['Somebody']);
  });

  it('ignores anything that is not a header or footer', () => {
    expect(findFields({ 'word/document.xml': REAL })).toEqual([]);
  });

  it('keys a field by where it is, so its value can change', () => {
    const [first] = findFields({ [HDR]: REAL });
    expect(first!.key).toBe(`${HDR}#0.0`);
  });
});

describe('fillFields', () => {
  const keys = () => findFields({ [HDR]: REAL }).map((f) => f.key);

  it('replaces a field spread over several runs with one value', () => {
    const [code] = keys();
    const out = fillFields({ [HDR]: REAL }, { [code!]: 'WDL 27-28' });
    expect(text(out[HDR]!)).toBe('WDL 27-28|File Title|Name| Page |1');
  });

  it('leaves the page field alone', () => {
    const out = fillFields({ [HDR]: REAL }, { [keys()[2]!]: 'Shane' });
    expect(out[HDR]).toContain('PAGE');
    expect(text(out[HDR]!)).toContain('Shane| Page |1');
  });

  // an unfilled field prints the school's own placeholder rather than a blank
  it('keeps the template text where there is no value', () => {
    expect(fillFields({ [HDR]: REAL }, {})[HDR]).toBe(REAL);
  });

  it('accepts an empty value as a deliberate blank', () => {
    const out = fillFields({ [HDR]: REAL }, { [keys()[1]!]: '' });
    expect(text(out[HDR]!)).toBe('BCP |26|-|27|Name| Page |1');
  });

  // the run carries the formatting — smallCaps, the border, the size — so an
  // emptied run has to stay in the document
  it('empties a run rather than removing it', () => {
    const out = fillFields({ [HDR]: REAL }, { [keys()[0]!]: 'X' });
    expect([...out[HDR]!.matchAll(/<w:r>/g)]).toHaveLength(
      [...REAL.matchAll(/<w:r>/g)].length,
    );
  });
});
