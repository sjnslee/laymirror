// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import {
  buildHeaderFooter,
  writeHeaderFooter,
  HEADER_PART,
  FOOTER_PART,
  HEADER_REL_ID,
  FOOTER_REL_ID,
} from '../src/docx/headers.js';
import { buildSectPr } from '../src/docx/sect.js';
import { readText } from '../src/docx/zip.js';
import { readTemplate } from '../src/profile/read-template.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { makeDocx, makeTemplate } from './fixture.js';

const meta = { title: '1AC', authors: 'A. Debater & B. Partner', teamCode: 'BCP 26-27' };
const donor = readTemplate(makeTemplate(), DEFAULT_LAY).profile;

const WML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const hdr = (body: string) => `<w:hdr xmlns:w="${WML}">${body}</w:hdr>`;
const withHeader = (headerXml: string) => ({ ...DEFAULT_LAY, headerXml });

describe('a donor header', () => {
  it('comes back untouched when it carries no tokens', () => {
    // the strongest promise available to a school: their bytes, unedited
    expect(buildHeaderFooter(donor, meta).headerXml).toBe(donor.headerXml);
    expect(buildHeaderFooter(donor, meta).footerXml).toBe(donor.footerXml);
  });

  it('fills a token word split across runs', () => {
    // word splits a run wherever a revision id changes, so this is the normal
    // case rather than a contrived one
    const profile = withHeader(
      hdr('<w:p><w:r><w:t>{{ti</w:t></w:r><w:r><w:t>tle}}</w:t></w:r></w:p>'),
    );
    const out = buildHeaderFooter(profile, meta).headerXml;
    expect(out).toContain('<w:t>1AC</w:t>');
    expect(out).not.toContain('{{');
  });

  it('keeps tabs and ptabs where they were', () => {
    const profile = withHeader(
      hdr(
        '<w:p><w:r><w:t>{{team}}</w:t></w:r><w:r><w:tab/></w:r>' +
          '<w:r><w:t>{{title}}</w:t></w:r></w:p>',
      ),
    );
    const out = buildHeaderFooter(profile, meta).headerXml;
    expect(out).toContain('<w:tab/>');
    expect(out.indexOf('BCP 26-27')).toBeLessThan(out.indexOf('<w:tab/>'));
    expect(out.indexOf('<w:tab/>')).toBeLessThan(out.indexOf('1AC'));
  });

  it('leaves the page field alone while filling around it', () => {
    const profile = withHeader(
      hdr(
        '<w:p><w:r><w:t>{{team}}</w:t></w:r>' +
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          '<w:r><w:t>7</w:t></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
      ),
    );
    const out = buildHeaderFooter(profile, meta).headerXml;
    expect(out).toContain('PAGE');
    expect(out).toContain('fldCharType="begin"');
    expect(out).toContain('fldCharType="end"');
    expect(out).toContain('<w:t>7</w:t>'); // the cached result word recomputes
  });

  it('escapes what it substitutes', () => {
    const profile = withHeader(hdr('<w:p><w:r><w:t>{{authors}}</w:t></w:r></w:p>'));
    const out = buildHeaderFooter(profile, meta).headerXml;
    expect(out).toContain('A. Debater &amp; B. Partner');
    expect(out).not.toContain('Debater & B');
  });

  it('protects a substituted value that ends in a space', () => {
    const profile = withHeader(hdr('<w:p><w:r><w:t>{{title}} </w:t></w:r></w:p>'));
    expect(buildHeaderFooter(profile, meta).headerXml).toContain('xml:space="preserve"');
  });
});

describe('a synthesised header and footer', () => {
  const { headerXml, footerXml } = buildHeaderFooter(DEFAULT_LAY, meta);

  it('carries the team code, the title and the authors', () => {
    expect(headerXml).toContain('BCP 26-27');
    expect(headerXml).toContain('1AC');
    expect(headerXml).toContain('A. Debater &amp; B. Partner');
  });

  it('sets its tab stops from the profile page, not a guess', () => {
    // letter with one inch margins leaves 9360 twips of text column
    expect(headerXml).toContain('w:val="right" w:pos="9360"');
    expect(headerXml).toContain('w:val="center" w:pos="4680"');
  });

  it('numbers pages with live fields, never a literal', () => {
    // our paginator can drift; PAGE and NUMPAGES are what keep the judge's
    // printed copy right anyway
    expect(footerXml).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>');
    expect(footerXml).toContain('<w:instrText xml:space="preserve"> NUMPAGES </w:instrText>');
    expect(footerXml).toContain('fldCharType="begin"');
  });

  it('declares itself as xml word will accept', () => {
    expect(headerXml.startsWith('<?xml version="1.0"')).toBe(true);
    expect(headerXml).toContain(`xmlns:w="${WML}"`);
  });
});

describe('wiring the parts into the package', () => {
  it('writes both parts, their content types and their relationships', () => {
    const parts = makeDocx();
    const refs = writeHeaderFooter(parts, donor, meta);

    expect(refs).toEqual({ headerRelId: HEADER_REL_ID, footerRelId: FOOTER_REL_ID });
    expect(parts[HEADER_PART]).toBeDefined();
    expect(parts[FOOTER_PART]).toBeDefined();

    const types = readText(parts, '[Content_Types].xml')!;
    expect(types).toContain('PartName="/word/header1.xml"');
    expect(types).toContain('wordprocessingml.footer+xml');

    // cardmirror's export has no document.xml.rels of its own
    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    expect(rels).toContain(`Id="${HEADER_REL_ID}"`);
    expect(rels).toContain('Target="footer1.xml"');
  });

  it('does not stack a second relationship when the file is saved again', () => {
    const parts = makeDocx();
    writeHeaderFooter(parts, donor, meta);
    writeHeaderFooter(parts, donor, meta);

    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    expect(rels.match(new RegExp(`Id="${HEADER_REL_ID}"`, 'g'))).toHaveLength(1);
    expect(readText(parts, '[Content_Types].xml')!.match(/header1\.xml/g)).toHaveLength(1);
  });

  it('hands back ids the section can reference', () => {
    const parts = makeDocx();
    const sect = buildSectPr(donor.page, writeHeaderFooter(parts, donor, meta));
    expect(sect).toContain(`<w:headerReference w:type="default" r:id="${HEADER_REL_ID}"/>`);
    expect(sect).toContain(`<w:footerReference w:type="default" r:id="${FOOTER_REL_ID}"/>`);
  });
});
