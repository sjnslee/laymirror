// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { readTemplate, readAttachedTemplate } from '../src/profile/read-template.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { unzip } from '../src/docx/zip.js';

// jsdom gives import.meta.url an http scheme, so resolve from the root
const bytes = new Uint8Array(readFileSync('tests/fixtures/donor.docx'));
const { profile, missing } = readTemplate(bytes, DEFAULT_LAY);

describe('readTemplate', () => {
  it('reads Normal as palatino 10pt', () => {
    expect(profile.types.paragraph.font).toBe('Palatino Linotype');
    expect(profile.types.paragraph.sizePt).toBe(10);
  });

  it('reads the pocket as a page-breaking centred small-caps heading', () => {
    const pocket = profile.types.pocket;
    expect(pocket.styleId).toBe('Heading1');
    expect(pocket.font).toBe('Times New Roman');
    expect(pocket.sizePt).toBe(20);
    expect(pocket.bold).toBe(true);
    expect(pocket.smallCaps).toBe(true);
    expect(pocket.align).toBe('center');
    expect(pocket.pageBreakBefore).toBe(true);
    expect(pocket.outlineLevel).toBe(0);
  });

  it('does not give hat or block a page break — only the pocket has one', () => {
    expect(profile.types.hat.pageBreakBefore).toBeUndefined();
    expect(profile.types.block.pageBreakBefore).toBeUndefined();
    expect(profile.types.hat.outlineLevel).toBe(1);
    expect(profile.types.block.outlineLevel).toBe(2);
  });

  it('inherits palatino into Tag through basedOn Normal', () => {
    const tag = profile.types.tag;
    expect(tag.styleName).toBe('Tag');
    expect(tag.bold).toBe(true);
    expect(tag.font).toBe('Palatino Linotype');
    expect(tag.sizePt).toBe(10);
  });

  it('inherits Cite from Tag and adds the thick rule', () => {
    const cite = profile.types.cite_paragraph;
    expect(cite.styleName).toBe('Cite');
    expect(cite.underline).toBe('thick');
    expect(cite.bold).toBe(true); // from Tag
    expect(cite.font).toBe('Palatino Linotype'); // from Normal
  });

  it('resolves the card body to the theme minor font, not palatino', () => {
    const card = profile.types.card_body;
    expect(card.styleName).toBe('card');
    expect(card.font).toBe('Cambria');
    expect(card.indentLeftDxa).toBe(288);
    expect(card.indentRightDxa).toBe(288);
    expect(card.underline).toBe('single');
    expect(card.spaceAfterPt).toBe(8);
    expect(card.lineSpacing).toEqual({ rule: 'auto', value: 259 });
  });

  it('turns bold off for the underline mark, as the donor does', () => {
    expect(profile.types.underline_mark.bold).toBe(false);
    expect(profile.types.underline_mark.underline).toBe('single');
  });

  it('reads the donor page setup — letter, half-inch margins', () => {
    expect(profile.page.widthTwips).toBe(12240);
    expect(profile.page.heightTwips).toBe(15840);
    expect(profile.page.margin.top).toBe(720);
    expect(profile.page.margin.bottom).toBe(1008);
  });

  it('lifts the header and footer', () => {
    expect(profile.headerXml).toContain('<w:hdr');
    expect(profile.footerXml).toContain('<w:ftr');
  });

  it('reduces attachedTemplate to a basename and decodes it', () => {
    expect(profile.attachedTemplate).toBe('Lay Cut Cards.dotx');
  });

  it('reports the types the donor has no style for', () => {
    expect(missing).toContain('analytic');
    expect(missing).toContain('undertag');
    expect(missing).not.toContain('tag');
    expect(missing).not.toContain('card_body');
  });

  it('keeps the fallback spec for anything missing', () => {
    expect(profile.types.analytic).toEqual(DEFAULT_LAY.types.analytic);
  });
});

describe('readAttachedTemplate', () => {
  it('strips a full path down to the basename', () => {
    const parts = unzip(bytes);
    expect(readAttachedTemplate(parts)).toBe('Lay Cut Cards.dotx');
    expect(readAttachedTemplate(parts)).not.toContain('/');
  });
});
