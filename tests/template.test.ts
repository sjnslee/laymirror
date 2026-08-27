// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { makeTemplate } from './fixture.js';
import { read } from '../src/template/template.js';
import { readText } from '../src/docx/zip.js';

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

  // the fixture breaks before heading 1 and nothing else, exactly as a lay
  // template does
  it('reports where the template breaks its pages', () => {
    expect(ok().breaks).toEqual(['pocket']);
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
