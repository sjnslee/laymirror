import { describe, expect, it } from 'vitest';
import { makeTemplate } from './fixture.js';
import {
  deriveBareStyles,
  deriveStyleMap,
  readStyles,
  readTemplate,
} from '../src/profile/read-template.js';
import type { StyleInfo } from '../src/profile/profile.js';

const ok = () => {
  const result = readTemplate(makeTemplate(), 'lay.docx');
  if (!result.ok) throw new Error(result.error);
  return result.profile;
};

describe('readStyles', () => {
  it('reads id, name and kind', () => {
    const styles = readStyles(
      '<w:styles>' +
        '<w:style w:type="paragraph" w:styleId="Tag"><w:name w:val="Tag"/></w:style>' +
        '<w:style w:type="character" w:styleId="Underline"><w:name w:val="Underline"/></w:style>' +
        '</w:styles>',
    );
    expect(styles).toEqual([
      { id: 'Tag', name: 'Tag', kind: 'paragraph' },
      { id: 'Underline', name: 'Underline', kind: 'character' },
    ]);
  });

  it('falls back to the id when a style has no name', () => {
    expect(readStyles('<w:styles><w:style w:styleId="Bare"></w:style></w:styles>')[0]!.name).toBe(
      'Bare',
    );
  });
});

describe('deriveStyleMap', () => {
  const styles: StyleInfo[] = [
    { id: 'Normal', name: 'Normal', kind: 'paragraph' },
    { id: 'Heading1', name: 'heading 1', kind: 'paragraph' },
    { id: 'Heading2', name: 'heading 2', kind: 'paragraph' },
    { id: 'Tag', name: 'Tag', kind: 'paragraph' },
    { id: 'Cite', name: 'Cite', kind: 'paragraph' },
    { id: 'Underline', name: 'Underline', kind: 'character' },
    { id: 'OldCite', name: 'Author-Date', kind: 'character' },
  ];

  // cardmirror exports a tag as Heading4; a lay template's tag style is `Tag`,
  // and identity would render every tag in word's stock italic blue
  it("sends a tag to the template's own Tag style", () => {
    expect(deriveStyleMap(styles)['Heading4']).toBe('Tag');
  });

  // the cite MARK is a run style, so it must never land on the paragraph
  // style called Cite — word cannot resolve a paragraph style as an rStyle
  it('sends the cite mark to a character style, never the Cite paragraph', () => {
    expect(deriveStyleMap(styles)['Style13ptBold']).toBe('OldCite');
  });

  it("sends the underline mark to the template's underline style", () => {
    expect(deriveStyleMap(styles)['StyleUnderline']).toBe('Underline');
  });

  it('leaves headings alone where the template defines them', () => {
    const map = deriveStyleMap(styles);
    expect(map['Heading1']).toBe('Heading1');
    expect(map['Heading2']).toBe('Heading2');
  });

  // mapping onto a style the template does not define would render the text
  // unstyled, which is worse than leaving cardmirror's own id in place
  it('never maps onto a style the template does not define', () => {
    const map = deriveStyleMap([{ id: 'Normal', name: 'Normal', kind: 'paragraph' }]);
    expect(Object.values(map)).toEqual([]);
  });
});

describe('deriveBareStyles', () => {
  it('finds the styles for the two types cardmirror exports bare', () => {
    expect(
      deriveBareStyles([
        { id: 'Cite', name: 'Cite', kind: 'paragraph' },
        { id: 'card', name: 'card', kind: 'paragraph' },
      ]),
    ).toEqual({ cite_paragraph: 'Cite', card_body: 'card' });
  });

  it('is null where the template says nothing', () => {
    expect(deriveBareStyles([])).toEqual({ cite_paragraph: null, card_body: null });
  });
});

describe('readTemplate', () => {
  it("keeps the template's styles and theme verbatim", () => {
    const profile = ok();
    expect(profile.snapshot!.parts['word/styles.xml']).toContain('Palatino Linotype');
    expect(profile.snapshot!.parts['word/theme/theme1.xml']).toContain('Calibri');
  });

  it('keeps the header and the real margins', () => {
    const profile = ok();
    expect(profile.snapshot!.parts['word/header1.xml']).toContain('PAGE');
    expect(profile.snapshot!.sectPr).toContain('w:bottom="1008"');
  });

  it('ids the profile by template, so two schools cannot collide', () => {
    expect(ok().id).toBe('template:lay.docx');
  });

  it('reduces the attached template to a basename', () => {
    expect(ok().snapshot!.attachedTemplate).toBe('Lay%20Cut%20Cards.dotx');
  });

  it('reports a file it cannot read instead of throwing', () => {
    const result = readTemplate(new Uint8Array([1, 2, 3]), 'notes.txt');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('notes.txt');
  });
});
