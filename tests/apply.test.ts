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

  it('carries the template styles, theme and numbering verbatim', () => {
    const parts = applied();
    expect(readText(parts, 'word/styles.xml')).toContain('Palatino Linotype');
    expect(readText(parts, 'word/theme/theme1.xml')).toContain('Calibri');
    expect(readText(parts, 'word/numbering.xml')).toContain('w:numId="7"');
  });

  it('marks the document so activation survives the file', () => {
    expect(readMarker(applied())).toBe('template:lay.docx');
  });

  // cardmirror's attached template is what lights up verbatim's ribbon
  it("keeps cardmirror's attached template", () => {
    const rels = readText(applied(), 'word/_rels/settings.xml.rels')!;
    expect(rels).toContain('Debate.dotm');
    expect(rels).not.toContain('Lay%20Cut%20Cards.dotx');
  });

  it('attaches the template by basename where the export has none', () => {
    const parts = unzip(makeExport());
    delete parts['word/_rels/settings.xml.rels'];
    writeText(
      parts,
      'word/settings.xml',
      readText(parts, 'word/settings.xml')!.replace(/<w:attachedTemplate[^>]*\/>/, ''),
    );
    const bytes = applyTemplate(zip(parts), blueprint(), {}, 'template:lay.docx');
    expect(readText(unzip(bytes), 'word/_rels/settings.xml.rels')).toContain(
      'Lay%20Cut%20Cards.dotx',
    );
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
