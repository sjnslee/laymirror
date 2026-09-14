// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { makeTemplate } from './fixture.js';
import { read } from '../src/template/template.js';
import { readText, unzip, zip } from '../src/docx/zip.js';

const ok = () => {
  const result = read(makeTemplate(), 'lay.docx');
  if (!result.ok) throw new Error(result.error);
  return result.blueprint;
};

describe('read', () => {
  it("keeps the template's styles and theme verbatim", () => {
    const parts = ok().snapshot.parts;
    expect(readText(parts, 'word/styles.xml')).toContain('Palatino Linotype');
    expect(readText(parts, 'word/theme/theme1.xml')).toContain('Calibri');
  });

  it('keeps the header and the real margins', () => {
    const blueprint = ok();
    expect(readText(blueprint.snapshot.parts, 'word/header1.xml')).toContain('PAGE');
    expect(blueprint.snapshot.sectPr).toContain('w:bottom="1008"');
  });

  it('maps cardmirror vocabulary onto the template', () => {
    expect(ok().styleMap['Heading4']).toBe('Tag');
    expect(ok().bareStyles.card_body).toBe('card');
  });

  it('finds the header fields', () => {
    expect(ok().fields.map((field) => field.label)).toEqual(['Team Code', 'lay']);
  });

  it('reports a file it cannot read instead of throwing', () => {
    const result = read(new Uint8Array([1, 2, 3]), 'notes.txt');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('notes.txt');
  });
});

// the kept copy sits in localStorage cardmirror shares, so it holds only what a
// blueprint is read from
describe('kept', () => {
  /** a template as a school ships it: macros, and sample cards in its body. */
  const shipped = () => {
    const parts = unzip(makeTemplate());
    parts['word/vbaProject.bin'] = new Uint8Array(200_000).fill(7);
    parts['docProps/thumbnail.jpeg'] = new Uint8Array(50_000).fill(9);
    return zip(parts);
  };

  const kept = () => {
    const result = read(shipped(), 'lay.docm');
    if (!result.ok) throw new Error(result.error);
    return result;
  };

  it('leaves out what never reaches the document', () => {
    const parts = unzip(kept().kept);
    expect(parts['word/vbaProject.bin']).toBeUndefined();
    expect(parts['docProps/thumbnail.jpeg']).toBeUndefined();
    expect(readText(parts, 'word/document.xml')).not.toContain('<w:p>');
  });

  it('reads back to the same blueprint', () => {
    const again = read(kept().kept, 'lay.docm');
    if (!again.ok) throw new Error(again.error);
    expect(again.blueprint).toEqual(kept().blueprint);
  });
});
