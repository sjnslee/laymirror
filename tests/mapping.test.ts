import { describe, expect, it } from 'vitest';
import {
  HEADING_LEVEL_TO_TYPE,
  LEGACY_BY_NAME,
  NATIVE_MARK_BY_ID,
  takesNativePath,
  validateMapping,
} from '../src/profile/mapping.js';
import { DEFAULT_PROFILE } from '../src/profile/defaults.js';
import type { Profile } from '../src/profile/profile.js';

const profile = (over: Partial<Profile> = {}): Profile => ({
  ...DEFAULT_PROFILE,
  id: 'template:lay.docx',
  ...over,
});

describe('round-trip vocabulary', () => {
  it('recognises the lay donor paragraph names', () => {
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
  it('rejects a donor that lacks the sentinel styles', () => {
    expect(takesNativePath(['Style13ptBold', 'Underline', 'Tag', 'Cite'], [])).toBe(false);
  });

  it('accepts once all three are present', () => {
    expect(takesNativePath(['Style13ptBold', 'StyleUnderline', 'Emphasis'], [])).toBe(true);
  });

  it('matches by name as well as by id', () => {
    expect(takesNativePath([], ['Style 13 pt Bold', 'Style Underline', 'Emphasis'])).toBe(true);
  });
});

describe('validateMapping', () => {
  it('says nothing about a profile with no template', () => {
    expect(validateMapping(DEFAULT_PROFILE)).toEqual([]);
  });

  it('accepts styles cardmirror knows by name', () => {
    expect(
      validateMapping(
        profile({
          styleMap: { Heading4: 'Tag', StyleUnderline: 'Underline' },
          styles: [
            { id: 'Tag', name: 'Tag', kind: 'paragraph' },
            { id: 'Underline', name: 'Underline', kind: 'character' },
          ],
        }),
      ),
    ).toEqual([]);
  });

  it('accepts headings, which resolve by outline level', () => {
    expect(
      validateMapping(profile({ styleMap: { Heading1: 'Heading1', Heading2: 'Heading2' } })),
    ).toEqual([]);
  });

  // this is the failure worth catching before someone cuts a whole file: it
  // exports into word perfectly and comes back as plain paragraphs
  it('flags a style name outside cardmirror vocabulary', () => {
    const warnings = validateMapping(
      profile({
        styleMap: { Heading4: 'Zonk' },
        styles: [{ id: 'Zonk', name: 'Zonk', kind: 'paragraph' }],
      }),
    );
    expect(warnings).toHaveLength(1);
    expect(warnings[0]!.styleId).toBe('Zonk');
    expect(warnings[0]!.message).toContain('ordinary paragraph');
  });

  it('checks the bare styles too', () => {
    const warnings = validateMapping(
      profile({
        bareStyles: { cite_paragraph: 'Nope', card_body: null },
        styles: [{ id: 'Nope', name: 'Nope', kind: 'paragraph' }],
      }),
    );
    expect(warnings.map((w) => w.styleId)).toEqual(['Nope']);
  });

  it('reports each style once, however many types point at it', () => {
    const warnings = validateMapping(
      profile({
        styleMap: { Heading4: 'Zonk', Analytic: 'Zonk' },
        styles: [{ id: 'Zonk', name: 'Zonk', kind: 'paragraph' }],
      }),
    );
    expect(warnings).toHaveLength(1);
  });
});
