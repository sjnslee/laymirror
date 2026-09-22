import { describe, expect, it } from 'vitest';
import { mergeNumbering, remapNumIds } from '../src/docx/numbering.js';

const template =
  '<w:numbering>' +
  '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>' +
  '</w:numbering>';

const exported =
  '<w:numbering>' +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>' +
  '</w:numbering>';

describe('mergeNumbering', () => {
  it('keeps the template definitions untouched', () => {
    const { xml } = mergeNumbering(exported, template);
    expect(xml).toContain('<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>');
    expect(xml).toContain('w:abstractNumId="3"');
  });

  it('moves cardmirror past the last template id', () => {
    const { xml, remap } = mergeNumbering(exported, template);
    expect(remap).toEqual({ 1: 8, 2: 9 });
    expect(xml).toContain('<w:num w:numId="8">');
    expect(xml).toContain('<w:num w:numId="9">');
    expect(xml).toContain('<w:abstractNum w:abstractNumId="4">');
    // the moved nums point at the moved abstractNum, not the template's
    expect(xml).toContain('<w:num w:numId="8"><w:abstractNumId w:val="4"/></w:num>');
  });

  // the schema orders abstractNum* before num*
  it('puts the abstract definitions before every num', () => {
    const xml = mergeNumbering(exported, template).xml!;
    expect(xml.lastIndexOf('<w:abstractNum ')).toBeLessThan(xml.indexOf('<w:num '));
  });

  it('carries the decimal and bullet formats through side by side', () => {
    const xml = mergeNumbering(exported, template).xml!;
    expect(xml).toContain('w:val="decimal"');
    expect(xml).toContain('w:val="bullet"');
  });

  it('leaves the part alone when cardmirror numbered nothing', () => {
    expect(mergeNumbering('<w:numbering/>', template)).toEqual({ xml: null, remap: {} });
    expect(mergeNumbering(null, template)).toEqual({ xml: null, remap: {} });
  });

  it('leaves the part alone when the template has no numbering at all', () => {
    expect(mergeNumbering(exported, null)).toEqual({ xml: null, remap: {} });
  });

  it('keeps cardmirror whole against a template with no lists of its own', () => {
    expect(mergeNumbering(exported, '<w:numbering/>')).toEqual({ xml: exported, remap: {} });
  });

  // applying twice must land in the same place rather than shift again
  it('recognises a part it has already merged', () => {
    const once = mergeNumbering(exported, template).xml!;
    expect(mergeNumbering(once, template)).toEqual({ xml: once, remap: {} });
  });
});

describe('remapNumIds', () => {
  const paragraph = (numId: number) =>
    '<w:p><w:pPr><w:pStyle w:val="Heading4"/>' +
    `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr></w:p>`;

  it('points a list reference at the merged id', () => {
    expect(remapNumIds(paragraph(1), { 1: 8 })).toContain('<w:numId w:val="8"/>');
  });

  it('leaves the level alone', () => {
    expect(remapNumIds(paragraph(1), { 1: 8 })).toContain('<w:ilvl w:val="0"/>');
  });

  it('leaves a reference the merge did not move', () => {
    expect(remapNumIds(paragraph(7), { 1: 8 })).toContain('<w:numId w:val="7"/>');
  });

  // w:val="1" is everywhere in a document; only a numPr's numId may move
  it('touches nothing outside a numPr', () => {
    const doc = `<w:body><w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:p>${paragraph(1)}</w:body>`;
    expect(remapNumIds(doc, { 1: 8 })).toContain('<w:outlineLvl w:val="1"/>');
  });
});
