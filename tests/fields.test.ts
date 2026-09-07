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

/** a real school header: a code split across runs by word's revision ids, a
 *  right tab, a title, and a live page field. */
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

/** the zero-width space a template author wraps editable text in, spelled out
 *  because it is invisible in a source file. */
const ZW = '\u200b';

/** the real BCP header's shape: `<zw>School<zw> <zw>26-27<zw>` — the year split
 *  across two runs — a right tab, then `<zw>File Title<zw>`. */
const MARKED = hdr(
  `<w:p><w:r><w:t>${ZW}School${ZW}</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>` +
    `<w:r><w:t>${ZW}2</w:t></w:r><w:r><w:t>6-27${ZW}</w:t></w:r>` +
    '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>' +
    `<w:r><w:t>${ZW}File Title${ZW}</w:t></w:r></w:p>` +
    '<w:p><w:r><w:t xml:space="preserve"> Page </w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
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

  // " page " and " of " are a word field's decoration, not the user's to edit
  it('offers nothing that belongs to a word field', () => {
    const labels = findFields({ [HDR]: REAL }).map((f) => f.label);
    expect(labels).not.toContain('Page');
    expect(labels).not.toContain('1');
  });

  it('prefers a zero-width mark where the template author left one', () => {
    const marked = hdr(
      '<w:p><w:r><w:t>Property of \u2060Somebody\u2060, do not copy</w:t></w:r></w:p>',
    );
    expect(findFields({ [HDR]: marked }).map((f) => f.label)).toEqual(['Somebody']);
  });

  // four marked spans, one split across runs, with the school's own text
  // between them
  it('reads the marks a real school template carries', () => {
    expect(findFields({ [HDR]: MARKED }).map((f) => f.label)).toEqual([
      'School',
      '26-27',
      'File Title',
    ]);
  });

  // a marked template says what is editable, so the guesswork must not run
  // alongside it and offer the rest of the header too
  it('offers only what is marked, once anything is', () => {
    expect(findFields({ [HDR]: MARKED }).map((f) => f.label)).not.toContain('Page');
  });

  // half a pair loses that field rather than shifting the rest onto the wrong
  // text
  it('drops a mark left without its partner', () => {
    const odd = hdr(`<w:p><w:r><w:t>${ZW}Kept${ZW} tail ${ZW}orphan</w:t></w:r></w:p>`);
    expect(findFields({ [HDR]: odd }).map((f) => f.label)).toEqual(['Kept']);
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

  // the run carries the formatting, so an emptied one has to stay
  it('empties a run rather than removing it', () => {
    const out = fillFields({ [HDR]: REAL }, { [keys()[0]!]: 'X' });
    expect([...out[HDR]!.matchAll(/<w:r>/g)]).toHaveLength(
      [...REAL.matchAll(/<w:r>/g)].length,
    );
  });

  // the value goes between the marks: eat one and the field stops existing
  it('writes inside the marks and leaves them there', () => {
    const marked = findFields({ [HDR]: MARKED });
    const out = fillFields({ [HDR]: MARKED }, {
      [marked[0]!.key]: 'WDL',
      [marked[1]!.key]: '27-28',
    })[HDR]!;
    expect([...out].filter((c) => c === ZW)).toHaveLength(6);
    expect(findFields({ [HDR]: out }).map((f) => f.label)).toEqual([
      'WDL',
      '27-28',
      'File Title',
    ]);
  });
});
