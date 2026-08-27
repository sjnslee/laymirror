import { describe, expect, it } from 'vitest';
import {
  breakingTypes,
  breaksBefore,
  deriveBareStyles,
  deriveStyleMap,
  HEADING_LEVEL_TO_TYPE,
  LEGACY_BY_NAME,
  NATIVE_MARK_BY_ID,
  readStyles,
  takesNativePath,
  validateMapping,
  type StyleInfo,
} from '../src/template/styles.js';

const style = (over: Partial<StyleInfo> & { id: string }): StyleInfo => ({
  name: over.id,
  kind: 'paragraph',
  basedOn: null,
  breakBefore: null,
  ...over,
});

describe('round-trip vocabulary', () => {
  it('recognises the lay template paragraph names', () => {
    expect(LEGACY_BY_NAME['tag']).toBe('tag');
    expect(LEGACY_BY_NAME['cite']).toBe('cite');
    expect(LEGACY_BY_NAME['card']).toBe('body');
    expect(LEGACY_BY_NAME['underline']).toBe('char-underline');
  });

  it('resolves headings by outline level, not by name', () => {
    expect(HEADING_LEVEL_TO_TYPE[1]).toBe('pocket');
    expect(HEADING_LEVEL_TO_TYPE[2]).toBe('hat');
    expect(HEADING_LEVEL_TO_TYPE[4]).toBe('tag');
  });

  it('keeps cite and underline marks only on the native path', () => {
    expect(NATIVE_MARK_BY_ID['Style13ptBold']).toBe('cite_mark');
    expect(NATIVE_MARK_BY_ID['Underline']).toBe('underline_mark');
  });
});

describe('native-path detection', () => {
  it('rejects a template that lacks the sentinel styles', () => {
    expect(takesNativePath(['Style13ptBold', 'Underline', 'Tag', 'Cite'], [])).toBe(false);
  });

  it('accepts once all three are present', () => {
    expect(takesNativePath(['Style13ptBold', 'StyleUnderline', 'Emphasis'], [])).toBe(true);
  });

  it('matches by name as well as by id', () => {
    expect(takesNativePath([], ['Style 13 pt Bold', 'Style Underline', 'Emphasis'])).toBe(true);
  });
});

describe('readStyles', () => {
  it('reads id, name, kind, inheritance and the page break', () => {
    const styles = readStyles(
      '<w:styles>' +
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
        '<w:pPr><w:pageBreakBefore/></w:pPr></w:style>' +
        '<w:style w:type="character" w:styleId="Underline"><w:name w:val="Underline"/>' +
        '<w:basedOn w:val="Normal"/></w:style>' +
        '</w:styles>',
    );
    expect(styles).toEqual([
      {
        id: 'Heading1',
        name: 'heading 1',
        kind: 'paragraph',
        basedOn: null,
        breakBefore: true,
      },
      {
        id: 'Underline',
        name: 'Underline',
        kind: 'character',
        basedOn: 'Normal',
        breakBefore: null,
      },
    ]);
  });

  it('falls back to the id when a style has no name', () => {
    expect(readStyles('<w:styles><w:style w:styleId="Bare"></w:style></w:styles>')[0]!.name).toBe(
      'Bare',
    );
  });
});

describe('breaksBefore', () => {
  it('is false for a style that says nothing', () => {
    expect(breaksBefore([style({ id: 'Normal' })], 'Normal')).toBe(false);
  });

  it('inherits down a basedOn chain', () => {
    const styles = [
      style({ id: 'Pocket', breakBefore: true }),
      style({ id: 'BigPocket', basedOn: 'Pocket' }),
    ];
    expect(breaksBefore(styles, 'BigPocket')).toBe(true);
  });

  // `w:val="0"` turns the toggle off, and a style that switched its parent's
  // break off must not be reported as breaking
  it('honours a style that turns its parent off', () => {
    const styles = [
      style({ id: 'Pocket', breakBefore: true }),
      style({ id: 'Quiet', basedOn: 'Pocket', breakBefore: false }),
    ];
    expect(breaksBefore(styles, 'Quiet')).toBe(false);
  });

  it('survives a chain that points at itself', () => {
    expect(breaksBefore([style({ id: 'Loop', basedOn: 'Loop' })], 'Loop')).toBe(false);
  });
});

describe('breakingTypes', () => {
  // this is what the editor draws a rule above
  it('names the block types the template starts a page before', () => {
    const styles = [
      style({ id: 'Heading1', name: 'heading 1', breakBefore: true }),
      style({ id: 'Heading2', name: 'heading 2' }),
    ];
    expect(breakingTypes(styles, deriveStyleMap(styles))).toEqual(['pocket']);
  });

  it('follows the mapping, not the export id', () => {
    const styles = [style({ id: 'Tag', name: 'Tag', breakBefore: true })];
    expect(breakingTypes(styles, { Heading4: 'Tag' })).toEqual(['tag']);
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

describe('validateMapping', () => {
  const bare = { cite_paragraph: null, card_body: null };

  it('says nothing about a template with no mapping', () => {
    expect(validateMapping([], {}, bare)).toEqual([]);
  });

  it('accepts styles cardmirror knows by name', () => {
    expect(
      validateMapping(
        [style({ id: 'Tag' }), style({ id: 'Underline', kind: 'character' })],
        { Heading4: 'Tag', StyleUnderline: 'Underline' },
        bare,
      ),
    ).toEqual([]);
  });

  it('accepts headings, which resolve by outline level', () => {
    expect(validateMapping([], { Heading1: 'Heading1', Heading2: 'Heading2' }, bare)).toEqual([]);
  });

  // this is the failure worth catching before someone cuts a whole file: it
  // exports into word perfectly and comes back as plain paragraphs
  it('flags a style name outside cardmirror vocabulary', () => {
    const warnings = validateMapping([style({ id: 'Zonk' })], { Heading4: 'Zonk' }, bare);
    expect(warnings).toHaveLength(1);
    expect(warnings[0]!.styleId).toBe('Zonk');
    expect(warnings[0]!.message).toContain('ordinary paragraph');
  });

  it('checks the bare styles too', () => {
    const warnings = validateMapping([style({ id: 'Nope' })], {}, {
      cite_paragraph: 'Nope',
      card_body: null,
    });
    expect(warnings.map((w) => w.styleId)).toEqual(['Nope']);
  });

  // on the native path cardmirror matches by id, so a style its own exporter
  // emits round-trips whether or not the legacy tables have heard of it
  it('accepts cardmirror ids when the template takes the native path', () => {
    const styles = [
      style({ id: 'Style13ptBold', name: 'Style 13 pt Bold', kind: 'character' }),
      style({ id: 'StyleUnderline', name: 'Style Underline', kind: 'character' }),
      style({ id: 'Emphasis', kind: 'character' }),
      style({ id: 'Analytic' }),
    ];
    expect(validateMapping(styles, { Analytic: 'Analytic' }, bare)).toEqual([]);
  });

  it('still flags them when it does not', () => {
    expect(
      validateMapping([style({ id: 'Analytic' })], { Analytic: 'Analytic' }, bare),
    ).toHaveLength(1);
  });

  it('reports each style once, however many types point at it', () => {
    const warnings = validateMapping(
      [style({ id: 'Zonk' })],
      { Heading4: 'Zonk', Analytic: 'Zonk' },
      bare,
    );
    expect(warnings).toHaveLength(1);
  });
});
