// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { makeExport, makeTemplate } from './fixture.js';
import { applyTemplate } from '../src/docx/apply.js';
import { findFields, type Values } from '../src/docx/fields.js';
import { readMarker } from '../src/docx/marker.js';
import { readText, unzip, writeText, zip, type Parts } from '../src/docx/zip.js';
import { headerParts, read } from '../src/template/template.js';

const blueprint = () => {
  const result = read(makeTemplate(), 'lay.docx');
  if (!result.ok) throw new Error(result.error);
  return result.blueprint;
};

const applied = (values: Values = {}) =>
  unzip(applyTemplate(makeExport(), blueprint(), values, 'template:lay.docx'));

const documentOf = (parts: Parts) => readText(parts, 'word/document.xml')!;

describe('applyTemplate', () => {
  it('puts the school header and footer back', () => {
    const parts = applied();
    expect(readText(parts, 'word/header1.xml')).toContain('PAGE');
    expect(readText(parts, 'word/footer1.xml')).toContain('lay');
  });

  it('replaces cardmirror 1in margins with the school section', () => {
    expect(documentOf(applied())).toContain('w:bottom="1008"');
  });

  it('carries the template styles and theme verbatim', () => {
    const parts = applied();
    expect(readText(parts, 'word/styles.xml')).toContain('Palatino Linotype');
    expect(readText(parts, 'word/theme/theme1.xml')).toContain('Calibri');
  });

  // cardmirror numbers cards through native word numbering, so replacing
  // numbering.xml wholesale left its numIds resolving to the template's lists
  it("keeps cardmirror's card numbering alongside the template's lists", () => {
    const numbering = readText(applied(), 'word/numbering.xml')!;
    expect(numbering).toContain('w:numId="7"');
    expect(numbering).toContain('w:val="bullet"');
    expect(numbering).toContain('w:val="decimal"');
    expect(numbering).toContain('<w:num w:numId="8">');
    expect(numbering).toContain('<w:num w:numId="9">');
  });

  it('points the numbered tags at the ids the merge minted', () => {
    const doc = documentOf(applied());
    expect(doc).toContain('<w:numId w:val="8"/>');
    expect(doc).toContain('<w:numId w:val="9"/>');
    expect(doc).not.toContain('<w:numId w:val="1"/>');
  });

  // cardmirror's styles.xml defines Analytic and Undertag; the template's
  // lands on top of it, so a template without them left the text unstyled
  it("keeps cardmirror's definition for a style the template does not have", () => {
    const parts = unzip(makeExport());
    writeText(
      parts,
      'word/document.xml',
      documentOf(parts).replace(
        '</w:body>',
        '<w:p><w:pPr><w:pStyle w:val="Analytic"/></w:pPr>' +
          '<w:r><w:t>so the aff loses</w:t></w:r></w:p></w:body>',
      ),
    );
    writeText(
      parts,
      'word/styles.xml',
      readText(parts, 'word/styles.xml')!.replace(
        '</w:styles>',
        '<w:style w:type="paragraph" w:styleId="Analytic"><w:name w:val="Analytic"/>' +
          '<w:pPr><w:ind w:left="720"/></w:pPr></w:style></w:styles>',
      ),
    );
    const out = unzip(applyTemplate(zip(parts), blueprint(), {}, 'template:lay.docx'));
    expect(documentOf(out)).toContain('w:val="Analytic"');
    expect(readText(out, 'word/styles.xml')).toContain('w:styleId="Analytic"');
    // and the template still wins where it says something
    expect(readText(out, 'word/styles.xml')).toContain('Palatino Linotype');
  });

  it('marks the document so activation survives the file', () => {
    expect(readMarker(applied())).toBe('template:lay.docx');
  });

  // verbatim's Document_Open runs UpdateStyles on any document attached to
  // Debate.dotm, which copies its styles over the template's on every open
  it("takes verbatim's template off the document", () => {
    const rels = readText(applied(), 'word/_rels/settings.xml.rels')!;
    expect(rels).not.toContain('Debate.dotm');
    expect(rels).toContain('Lay%20Cut%20Cards.dotx');
  });

  it('detaches verbatim even when the template attaches nothing', () => {
    const template = unzip(makeTemplate());
    delete template['word/_rels/settings.xml.rels'];
    const result = read(zip(template), 'lay.docx');
    if (!result.ok) throw new Error(result.error);

    const parts = unzip(applyTemplate(makeExport(), result.blueprint, {}, 'template:lay.docx'));
    expect(readText(parts, 'word/_rels/settings.xml.rels') ?? '').not.toContain('attachedTemplate');
    expect(readText(parts, 'word/settings.xml')).not.toContain('w:attachedTemplate');
  });

  // an apply lands on a file laymirror has already applied to, so it has to
  // land in the same place rather than stack
  it('is the same file applied twice', () => {
    const once = applyTemplate(makeExport(), blueprint(), {}, 'template:lay.docx');
    const twice = applyTemplate(once, blueprint(), {}, 'template:lay.docx');
    expect(documentOf(unzip(twice))).toBe(documentOf(unzip(once)));
  });

  it('throws on a partial read rather than writing half a file', () => {
    expect(() =>
      applyTemplate(new Uint8Array([80, 75, 3, 4]), blueprint(), {}, 'x'),
    ).toThrow();
  });
});

describe('applyTemplate — the header the user typed', () => {
  const teamCode = () => findFields(headerParts(blueprint().snapshot))[0]!.key;

  it('writes a value into the school header', () => {
    const parts = applied({ [teamCode()]: 'WDL 27-28' });
    expect(readText(parts, 'word/header1.xml')).toContain('WDL 27-28');
  });

  // the template is the source, never the last save, or a value typed once
  // could never be typed over
  it('starts from the template, so a value can be replaced', () => {
    const key = teamCode();
    const once = applyTemplate(makeExport(), blueprint(), { [key]: 'first' }, 'id');
    const twice = applyTemplate(once, blueprint(), { [key]: 'second' }, 'id');
    const header = readText(unzip(twice), 'word/header1.xml')!;
    expect(header).toContain('second');
    expect(header).not.toContain('first');
  });
});

describe('applyTemplate — style mapping', () => {
  // cardmirror exports a tag as Heading4; the school's tag style is Tag
  it('sends a tag to the school Tag style', () => {
    const doc = documentOf(applied());
    expect(doc).toContain('w:val="Tag"');
    expect(doc).not.toContain('w:val="Heading4"');
  });

  it('styles a cite paragraph that arrived with no style at all', () => {
    expect(documentOf(applied())).toContain('w:val="Cite"');
  });

  it('styles a card body that arrived with no style at all', () => {
    expect(documentOf(applied())).toContain('w:val="card"');
  });

  it('leaves ordinary prose after a heading alone', () => {
    // plain text following a hat must not be indented as evidence
    const doc = documentOf(applied());
    const paragraph = [...doc.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)]
      .map((m) => m[0])
      .find((p) => p.includes('just a paragraph'))!;
    expect(paragraph).not.toContain('w:pStyle');
  });
});
