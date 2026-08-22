// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import {
  applyStylesheet,
  hasStylesheet,
  parseRules,
  probeDocument,
  removeStylesheet,
  retarget,
  STYLE_ID,
  targetsFor,
} from '../src/render/css.js';
import { DEFAULT_PROFILE } from '../src/profile/defaults.js';
import type { Profile } from '../src/profile/profile.js';

const profile = (over: Partial<Profile> = {}): Profile => ({
  ...DEFAULT_PROFILE,
  id: 'template:lay.docx',
  ...over,
});

describe('probeDocument', () => {
  it('names every style so the renderer must emit a rule for it', () => {
    const xml = probeDocument(['Tag', 'Cite']);
    expect(xml).toContain('w:val="Tag"');
    expect(xml).toContain('w:val="Cite"');
  });

  it('is still a valid body when nothing is asked for', () => {
    expect(probeDocument([])).toContain('<w:body>');
  });
});

describe('targetsFor', () => {
  it('points a mapped style at cardmirror class', () => {
    const targets = targetsFor(
      profile({ styleMap: { Heading4: 'Tag', Heading2: 'Heading2' } }),
    );
    expect(targets.get('Tag')).toBe('.pmd-tag');
    expect(targets.get('Heading2')).toBe('.pmd-hat');
  });

  it('includes the two bare types', () => {
    const targets = targetsFor(
      profile({ bareStyles: { cite_paragraph: 'Cite', card_body: 'card' } }),
    );
    expect(targets.get('Cite')).toBe('.pmd-cite-para');
    expect(targets.get('card')).toBe('.pmd-card-body');
  });

  it('is empty without a template', () => {
    expect(targetsFor(DEFAULT_PROFILE).size).toBe(0);
  });
});

describe('parseRules', () => {
  it('splits a flat sheet into selectors and bodies', () => {
    expect(parseRules('a { color: red } b { font-size: 2pt }')).toEqual([
      { selector: 'a', body: 'color: red' },
      { selector: 'b', body: 'font-size: 2pt' },
    ]);
  });

  it('drops empty rules', () => {
    expect(parseRules('a { }')).toEqual([]);
  });
});

describe('retarget', () => {
  // docx-preview names a style class from its lowercased id
  const css =
    '.docx { --docx-majorHAnsi-font: Calibri; --docx-accent1-color: #4F81BD; }\n' +
    'p.docx_tag span { font-weight: bold; font-size: 10.00pt }\n' +
    'p.docx_heading2 span { font-family: var(--docx-majorHAnsi-font); color: #4F81BD }\n' +
    'p.docx_unrelated span { color: green }';

  const targets = new Map([
    ['Tag', '.pmd-tag'],
    ['Heading2', '.pmd-hat'],
  ]);

  it('re-points a style rule at cardmirror class', () => {
    const out = retarget(css, targets);
    expect(out).toContain('.pmd-tag');
    expect(out).toContain('font-weight: bold');
  });

  // without these every var(--docx-majorHAnsi-font) resolves to nothing
  it('carries the theme variables across', () => {
    expect(retarget(css, targets)).toContain('--docx-majorHAnsi-font: Calibri');
  });

  it('scopes everything to the editor so nothing leaks', () => {
    for (const line of retarget(css, targets).split('\n')) {
      expect(line).toContain('#editor');
    }
  });

  it('ignores styles nothing maps to', () => {
    expect(retarget(css, targets)).not.toContain('green');
  });

  it('covers the block itself as well as its spans', () => {
    // a docx paragraph carries run properties on spans; cardmirror puts the
    // text straight in the block
    const out = retarget('p.docx_tag span { font-weight: bold }', targets);
    expect(out).toMatch(/\.pmd-tag,[^{]*\.pmd-tag span/);
  });
});

describe('the stylesheet element', () => {
  it('installs, replaces and removes cleanly', () => {
    applyStylesheet('.pmd-tag { color: red }');
    expect(hasStylesheet()).toBe(true);
    applyStylesheet('.pmd-tag { color: blue }');
    expect(document.querySelectorAll(`#${STYLE_ID}`)).toHaveLength(1);
    expect(document.getElementById(STYLE_ID)!.textContent).toContain('blue');
    removeStylesheet();
    expect(hasStylesheet()).toBe(false);
  });
});
