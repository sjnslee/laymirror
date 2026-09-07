import { describe, expect, it } from 'vitest';
import {
  deriveBareStyles,
  deriveStyleMap,
  readStyles,
  type StyleInfo,
} from '../src/template/styles.js';

const style = (over: Partial<StyleInfo> & { id: string }): StyleInfo => ({
  name: over.id,
  kind: 'paragraph',
  ...over,
});

describe('readStyles', () => {
  it('reads id, name and kind', () => {
    const styles = readStyles(
      '<w:styles>' +
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
        '<w:pPr><w:pageBreakBefore/></w:pPr></w:style>' +
        '<w:style w:type="character" w:styleId="Underline"><w:name w:val="Underline"/>' +
        '</w:style>' +
        '</w:styles>',
    );
    expect(styles).toEqual([
      { id: 'Heading1', name: 'heading 1', kind: 'paragraph' },
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
  const styles = [
    style({ id: 'Normal' }),
    style({ id: 'Heading1', name: 'heading 1' }),
    style({ id: 'Heading2', name: 'heading 2' }),
    style({ id: 'Tag' }),
    style({ id: 'Cite' }),
    style({ id: 'Underline', kind: 'character' }),
    style({ id: 'OldCite', name: 'Author-Date', kind: 'character' }),
  ];

  // cardmirror exports a tag as Heading4; a lay template's tag style is `Tag`,
  // and identity would render every tag in word's stock italic blue
  it("sends a tag to the template's own Tag style", () => {
    expect(deriveStyleMap(styles)['Heading4']).toBe('Tag');
  });

  // the cite MARK is a run style, so it must never land on the paragraph style
  // called Cite — word cannot resolve a paragraph style as an rStyle
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
    expect(Object.values(deriveStyleMap([style({ id: 'Normal' })]))).toEqual([]);
  });
});

describe('deriveBareStyles', () => {
  it('finds the styles for the two types cardmirror exports bare', () => {
    expect(deriveBareStyles([style({ id: 'Cite' }), style({ id: 'card' })])).toEqual({
      cite_paragraph: 'Cite',
      card_body: 'card',
    });
  });

  it('is null where the template says nothing', () => {
    expect(deriveBareStyles([])).toEqual({ cite_paragraph: null, card_body: null });
  });
});
