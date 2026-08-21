// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { toCss, applyStylesheet, removeStylesheet, STYLE_ID } from '../src/render/css.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';

const css = toCss(DEFAULT_LAY);

/** the block of declarations for one selector. */
function ruleFor(selector: string): string {
  const idx = css.indexOf(`) ${selector} {`);
  if (idx === -1) throw new Error(`no rule emitted for ${selector}`);
  return css.slice(idx, css.indexOf('}', idx));
}

describe('toCss', () => {
  it('gives each type the donor font', () => {
    expect(ruleFor('.pmd-tag')).toContain('Palatino Linotype');
    expect(ruleFor('.pmd-pocket')).toContain('Tinos'); // times new roman substitute
    expect(ruleFor('.pmd-card-body')).toContain('Caladea'); // cambria substitute
  });

  it('renders the pocket as a centred small-caps heading', () => {
    const rule = ruleFor('.pmd-pocket');
    expect(rule).toContain('font-size: 20pt');
    expect(rule).toContain('font-variant-caps: small-caps');
    expect(rule).toContain('text-align: center');
    expect(rule).toContain('font-weight: 700');
  });

  it('gives the cite a thick rule rather than a double one', () => {
    const rule = ruleFor('.pmd-cite-para');
    expect(rule).toContain('text-decoration: underline');
    expect(rule).toContain('text-decoration-thickness: 2px');
    expect(rule).not.toContain('double');
  });

  it('indents the card body by the donor 288 dxa', () => {
    const rule = ruleFor('.pmd-card-body');
    expect(rule).toContain('margin-left: 0.2000in');
    expect(rule).toContain('margin-right: 0.2000in');
    expect(rule).toContain('line-height: 1.079');
  });

  it('marks every declaration important so appearance settings cannot win', () => {
    const decls = css.match(/^\s{2}[a-z-]+:[^;]+;$/gm) ?? [];
    const notImportant = decls.filter((d) => !d.includes('!important') && !d.includes('--pmd-'));
    expect(notImportant).toEqual([]);
  });

  it('keeps cardmirror own variables in step with the rules', () => {
    expect(css).toContain('--pmd-size-tag: 10pt');
    expect(css).toContain('--pmd-size-pocket: 20pt');
  });

  it('scopes everything to the editor', () => {
    const selectors = css.match(/^[^\s@/].*\{$/gm) ?? [];
    expect(selectors.every((s) => s.includes('#editor') || s.includes('.pmd-pane-editor'))).toBe(true);
  });
});

describe('stylesheet injection', () => {
  it('applies, updates in place, and leaves nothing behind', () => {
    applyStylesheet(css);
    expect(document.getElementById(STYLE_ID)?.textContent).toContain('laymirror');

    applyStylesheet('/* second */');
    expect(document.querySelectorAll(`#${STYLE_ID}`)).toHaveLength(1);
    expect(document.getElementById(STYLE_ID)?.textContent).toBe('/* second */');

    removeStylesheet();
    expect(document.getElementById(STYLE_ID)).toBeNull();
    removeStylesheet(); // idempotent
  });
});
