import { describe, it, expect } from 'vitest';
import {
  LEGACY_BY_NAME,
  NATIVE_MARK_BY_ID,
  HEADING_LEVEL_TO_TYPE,
  takesNativePath,
  validateMapping,
} from '../src/profile/mapping.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';

describe('round-trip vocabulary', () => {
  it('recognises the lay donor paragraph names', () => {
    expect(LEGACY_BY_NAME['tag']).toBe('tag');
    expect(LEGACY_BY_NAME['cite']).toBe('cite');
    expect(LEGACY_BY_NAME['card']).toBe('body');
    expect(LEGACY_BY_NAME['underline']).toBe('char-underline');
  });

  it('maps the donor style names the profile actually uses', () => {
    const t = DEFAULT_LAY.types;
    expect(LEGACY_BY_NAME[t.tag.styleName.toLowerCase()]).toBe('tag');
    expect(LEGACY_BY_NAME[t.cite_paragraph.styleName.toLowerCase()]).toBe('cite');
    expect(LEGACY_BY_NAME[t.card_body.styleName.toLowerCase()]).toBe('body');
    expect(LEGACY_BY_NAME[t.underline_mark.styleName.toLowerCase()]).toBe('char-underline');
  });

  it('resolves headings by outline level, not by name', () => {
    const t = DEFAULT_LAY.types;
    expect(HEADING_LEVEL_TO_TYPE[(t.pocket.outlineLevel ?? 0) + 1]).toBe('pocket');
    expect(HEADING_LEVEL_TO_TYPE[(t.hat.outlineLevel ?? 0) + 1]).toBe('hat');
    expect(HEADING_LEVEL_TO_TYPE[(t.block.outlineLevel ?? 0) + 1]).toBe('block');
  });

  it('keeps cite and underline marks only on the native path', () => {
    expect(NATIVE_MARK_BY_ID['Style13ptBold']).toBe('cite_mark');
    expect(NATIVE_MARK_BY_ID['Underline']).toBe('underline_mark');
  });
});

describe('native-path detection', () => {
  it('rejects the donor as shipped — it lacks StyleUnderline and Emphasis', () => {
    expect(takesNativePath(['Style13ptBold', 'Underline', 'Tag', 'Cite', 'card'], [])).toBe(false);
  });

  it('accepts once the three sentinel styles are present', () => {
    expect(
      takesNativePath(['Style13ptBold', 'StyleUnderline', 'Emphasis', 'Tag'], []),
    ).toBe(true);
  });

  it('matches by name as well as by id', () => {
    expect(takesNativePath([], ['Style 13 pt Bold', 'Style Underline', 'Emphasis'])).toBe(true);
  });
});

describe('validateMapping', () => {
  it('passes every block type in the default profile', () => {
    const blockWarnings = validateMapping(DEFAULT_LAY).filter(
      (w) => w.type !== 'cite_mark' && w.type !== 'underline_mark',
    );
    expect(blockWarnings).toEqual([]);
  });

  it('warns that the cite mark needs the native path', () => {
    const warning = validateMapping(DEFAULT_LAY).find((w) => w.type === 'cite_mark');
    expect(warning?.message).toContain('native path');
  });

  it('flags a style name outside the vocabulary', () => {
    const bad = {
      ...DEFAULT_LAY,
      types: { ...DEFAULT_LAY.types, tag: { ...DEFAULT_LAY.types.tag, styleName: 'Zonk' } },
    };
    expect(validateMapping(bad).some((w) => w.type === 'tag')).toBe(true);
  });
});
