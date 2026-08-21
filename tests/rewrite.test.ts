// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { rewriteDocx } from '../src/docx/rewrite.js';
import { readMarker } from '../src/docx/marker.js';
import { CONTENT_TYPES, readText, unzip, zip } from '../src/docx/zip.js';
import { readTemplate } from '../src/profile/read-template.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { makeDocx, makeExport, makeTemplate } from './fixture.js';

const meta = { title: '1AC', authors: 'A. Debater', teamCode: 'BCP 26-27' };
const profile = { ...readTemplate(makeTemplate(), DEFAULT_LAY).profile, id: 'template:donor' };

const source = makeExport();
const out = unzip(rewriteDocx(source, profile, meta));
const documentXml = readText(out, 'word/document.xml')!;

/** the nth <w:p> of the body, as raw xml. */
const paragraph = (n: number): string => {
  const all = documentXml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [];
  return all[n] ?? '';
};

describe('rewriteDocx', () => {
  it('produces a package word will open, content types first', () => {
    expect(Object.keys(out)[0]).toBe(CONTENT_TYPES);
    expect(out['word/document.xml']).toBeDefined();
    expect(out['word/styles.xml']).toBeDefined();
  });

  it('refuses a partial read rather than writing a broken file', () => {
    // a save caught mid-write unzips to something that is not a docx
    expect(() => rewriteDocx(zip({}), profile, meta)).toThrow(/not a complete docx/);
    const truncated = makeDocx();
    delete truncated['word/document.xml'];
    expect(() => rewriteDocx(zip(truncated), profile, meta)).toThrow();
  });

  it('remaps the tag from cardmirror Heading4 to the donor style', () => {
    expect(paragraph(0)).toContain(`<w:pStyle w:val="${profile.types.tag.styleId}"/>`);
    expect(documentXml).not.toContain('w:val="Heading4"');
  });

  it('gives the cite paragraph a style cardmirror never wrote', () => {
    expect(paragraph(1)).toContain(`<w:pStyle w:val="${profile.types.cite_paragraph.styleId}"/>`);
  });

  it('gives the card body one too, and carries it one paragraph past the marks', () => {
    const body = `<w:pStyle w:val="${profile.types.card_body.styleId}"/>`;
    expect(paragraph(2)).toContain(body);
    expect(paragraph(3)).toContain(body);
  });

  it('turns a carried page break back into a real one', () => {
    // cardmirror's model has nowhere to keep a page break, so it rides
    // through as text and is restored here
    expect(paragraph(4)).toContain('<w:br w:type="page"/>');
    expect(documentXml).not.toContain('[page break]');
  });

  it('leaves ordinary prose after a heading alone', () => {
    // over-reaching here would indent a paragraph that is not evidence
    expect(paragraph(6)).not.toContain('<w:pStyle');
  });

  it('remaps the marks inside the runs', () => {
    expect(documentXml).toContain(`<w:rStyle w:val="${profile.types.underline_mark.styleId}"/>`);
    expect(documentXml).toContain(`<w:rStyle w:val="${profile.types.cite_mark.styleId}"/>`);
  });

  it('replaces the hardcoded section with the donor page and the new parts', () => {
    expect(documentXml).toContain('<w:pgSz w:w="12240" w:h="15840"/>');
    expect(documentXml).toContain('w:top="720"');
    expect(documentXml).toContain('w:bottom="1008"');
    expect(documentXml).toContain('<w:headerReference w:type="default"');
    expect(documentXml).toContain('<w:footerReference w:type="default"');
    expect(documentXml.match(/<w:sectPr\b/g)).toHaveLength(1);
    // the section references its parts by r:id, so the prefix must be bound
    expect(documentXml).toContain('xmlns:r=');
  });

  it('adds the header and footer parts and their wiring', () => {
    expect(out['word/header1.xml']).toBeDefined();
    expect(out['word/footer1.xml']).toBeDefined();
    expect(readText(out, CONTENT_TYPES)).toContain('wordprocessingml.header+xml');
    expect(readText(out, 'word/_rels/document.xml.rels')).toContain('Target="header1.xml"');
  });

  it('keeps the styles that hold cardmirror on its native import path', () => {
    const styles = readText(out, 'word/styles.xml')!;
    expect(styles).toContain('Style13ptBold');
    expect(styles).toContain('StyleUnderline');
    expect(styles).toContain('Emphasis');
  });

  it('points the attached template at a basename, never a path', () => {
    const rels = readText(out, 'word/_rels/settings.xml.rels')!;
    expect(rels).toContain('Target="Lay Cut Cards.dotx"');
    expect(rels).not.toContain('Debate.dotm');
    expect(rels).not.toContain('file:///');
  });

  it('marks the file so it opens lay next time', () => {
    expect(readMarker(out)).toBe('template:donor');
  });

  it('leaves parts it does not model exactly as it found them', () => {
    const before = unzip(source);
    for (const name of ['word/numbering.xml', '_rels/.rels']) {
      expect(Buffer.from(out[name]!)).toEqual(Buffer.from(before[name]!));
    }
  });

  it('is stable when the same file is saved twice', () => {
    const twice = unzip(rewriteDocx(rewriteDocx(source, profile, meta), profile, meta));
    expect(readText(twice, 'word/_rels/document.xml.rels')!.match(/Target="header1\.xml"/g))
      .toHaveLength(1);
    expect(readText(twice, 'word/document.xml')!.match(/<w:sectPr\b/g)).toHaveLength(1);
    expect(readMarker(twice)).toBe('template:donor');
  });
});
