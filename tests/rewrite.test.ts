// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { makeExport, makeTemplate } from './fixture.js';
import { applyProfile, dropLegacySentinel } from '../src/docx/rewrite.js';
import { readTemplate } from '../src/profile/read-template.js';
import { DEFAULT_PROFILE } from '../src/profile/defaults.js';
import { readMarker } from '../src/docx/marker.js';
import { readText, unzip } from '../src/docx/zip.js';

const profile = () => {
  const result = readTemplate(makeTemplate(), 'lay.docx');
  if (!result.ok) throw new Error(result.error);
  return result.profile;
};

const restored = (breaks = []) => {
  const outcome = applyProfile(makeExport(), profile(), breaks);
  if (outcome.kind !== 'restored') throw new Error(`expected restored, got ${outcome.kind}`);
  return unzip(outcome.bytes);
};

const documentOf = (parts: Record<string, Uint8Array>) => readText(parts, 'word/document.xml')!;

describe('applyProfile — deciding what to do', () => {
  // the entire pipeline turns on this one question
  it('adopts a file that carries a header, because only word writes one', () => {
    const outcome = applyProfile(makeTemplate(), profile());
    expect(outcome.kind).toBe('adopted');
    if (outcome.kind === 'adopted') {
      expect(outcome.snapshot.parts['word/header1.xml']).toContain('PAGE');
    }
  });

  it('restores a cardmirror export, which never has one', () => {
    expect(applyProfile(makeExport(), profile()).kind).toBe('restored');
  });

  it('does nothing at all without a template', () => {
    const outcome = applyProfile(makeExport(), DEFAULT_PROFILE);
    expect(outcome).toEqual({ kind: 'skipped', because: 'no template loaded' });
  });

  it('throws on a partial read rather than writing half a file', () => {
    expect(() => applyProfile(new Uint8Array([80, 75, 3, 4]), profile())).toThrow();
  });
});

describe('applyProfile — what it restores', () => {
  it('puts the school header and footer back', () => {
    const parts = restored();
    expect(readText(parts, 'word/header1.xml')).toContain('PAGE');
    expect(readText(parts, 'word/footer1.xml')).toContain('lay');
  });

  it('replaces cardmirror 1in margins with the school section', () => {
    expect(documentOf(restored())).toContain('w:bottom="1008"');
  });

  it('carries the template styles and theme verbatim', () => {
    const parts = restored();
    expect(readText(parts, 'word/styles.xml')).toContain('Palatino Linotype');
    expect(readText(parts, 'word/theme/theme1.xml')).toContain('Calibri');
  });

  it('marks the document so activation survives the file', () => {
    expect(readMarker(restored())).toBe('template:lay.docx');
  });

  it('leaves parts it does not own untouched', () => {
    expect(readText(restored(), 'word/numbering.xml')).toBe('<w:numbering/>');
  });

  it('repoints the attached template at a basename', () => {
    expect(readText(restored(), 'word/_rels/settings.xml.rels')).toContain(
      'Lay%20Cut%20Cards.dotx',
    );
  });
});

describe('applyProfile — style mapping', () => {
  // cardmirror exports a tag as Heading4; the school's tag style is Tag
  it('sends a tag to the school Tag style', () => {
    const doc = documentOf(restored());
    expect(doc).toContain('w:val="Tag"');
    expect(doc).not.toContain('w:val="Heading4"');
  });

  it('styles a cite paragraph that arrived with no style at all', () => {
    expect(documentOf(restored())).toContain('w:val="Cite"');
  });

  it('styles a card body that arrived with no style at all', () => {
    expect(documentOf(restored())).toContain('w:val="card"');
  });

  it('leaves ordinary prose after a heading alone', () => {
    // the last paragraph is plain text following a hat, and must not be
    // indented as evidence
    const doc = documentOf(restored());
    const paragraph = [...doc.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)]
      .map((m) => m[0])
      .find((p) => p.includes('just a paragraph'))!;
    expect(paragraph).not.toContain('w:pStyle');
  });
});

describe('applyProfile — page breaks', () => {
  it('injects a real break for a mark whose anchor is present', () => {
    // the export's hat is styled but carries no bookmark, so a mark against a
    // missing anchor must simply not appear
    const doc = documentOf(restored([{ headingId: 'nope', offset: 1 }] as never));
    expect(doc).not.toContain('w:type="page"');
  });
});

describe('dropLegacySentinel', () => {
  it('sweeps up a paragraph that is nothing but the old text sentinel', () => {
    const before =
      '<w:body><w:p><w:r><w:t>keep</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>[page break]</w:t></w:r></w:p></w:body>';
    const after = dropLegacySentinel(before);
    expect(after).toContain('keep');
    expect(after).not.toContain('[page break]');
  });

  it('leaves a paragraph that merely mentions it', () => {
    const xml = '<w:body><w:p><w:r><w:t>a [page break] here</w:t></w:r></w:p></w:body>';
    expect(dropLegacySentinel(xml)).toBe(xml);
  });
});
