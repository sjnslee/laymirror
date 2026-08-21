// profile -> editor stylesheet. the other half of `toOoxml`; both read the
// same profile so the screen cannot disagree with the file.
//
// every declaration is !important because cardmirror writes its display
// settings as inline custom properties on :root and #editor, and rewrites
// them whenever the user touches appearance settings.

import { EDITOR_SELECTOR, CSS_VAR } from '../host/cardmirror.js';
import type { BlockType, Profile, RunType, TypeSpec } from '../profile/profile.js';

export const STYLE_ID = 'laymirror-style';

/** class cardmirror puts on each block/mark. `paragraph` has no class of its
 *  own — it is a bare <p> inside the editor. */
const SELECTOR: Record<BlockType | RunType, string> = {
  paragraph: 'p',
  pocket: '.pmd-pocket',
  hat: '.pmd-hat',
  block: '.pmd-block',
  tag: '.pmd-tag',
  analytic: '.pmd-analytic',
  undertag: '.pmd-undertag',
  cite_paragraph: '.pmd-cite-para',
  card_body: '.pmd-card-body',
  underline_mark: '.pmd-underline',
  emphasis_mark: '.pmd-emphasis',
  cite_mark: '.pmd-cite',
  analytic_mark: '.pmd-analytic-run',
  undertag_mark: '.pmd-undertag-run',
};

const BLOCK_TYPES: readonly BlockType[] = [
  'pocket',
  'hat',
  'block',
  'tag',
  'analytic',
  'undertag',
  'cite_paragraph',
  'card_body',
];

/** which text type an element in the editor is, by the class cardmirror puts
 *  on it. read off the same table the stylesheet is built from, so the two
 *  cannot disagree about what a card body is. */
export function blockTypeOf(el: Element): BlockType {
  for (const type of BLOCK_TYPES) {
    if (el.matches(SELECTOR[type])) return type;
  }
  return 'paragraph';
}

const dxaToIn = (dxa: number) => `${(dxa / 1440).toFixed(4)}in`;

/** what a family is actually drawn in: the profile's substitute for it, or
 *  the family itself. the single answer to that question — the settings ui
 *  asks it too. */
export function fontStackFor(profile: Profile, family: string): string {
  return profile.fontFallbacks[family] ?? `"${family}", serif`;
}

function fontStack(spec: TypeSpec, profile: Profile): string | null {
  return spec.font ? fontStackFor(profile, spec.font) : null;
}

function declarations(spec: TypeSpec, profile: Profile): string[] {
  const out: string[] = [];
  const push = (prop: string, value: string) => out.push(`  ${prop}: ${value} !important;`);

  const font = fontStack(spec, profile);
  if (font) push('font-family', font);
  if (spec.sizePt !== undefined) push('font-size', `${spec.sizePt}pt`);
  if (spec.bold !== undefined) push('font-weight', spec.bold ? '700' : '400');
  if (spec.italic !== undefined) push('font-style', spec.italic ? 'italic' : 'normal');
  if (spec.smallCaps) push('font-variant-caps', 'small-caps');
  if (spec.color) push('color', `#${spec.color}`);
  if (spec.align) push('text-align', spec.align);

  if (spec.underline !== undefined) {
    if (spec.underline === 'none') {
      push('text-decoration', 'none');
    } else {
      push('text-decoration', spec.underline === 'double' ? 'underline double' : 'underline');
      // word's "thick" is a heavier rule, not a second line
      if (spec.underline === 'thick') push('text-decoration-thickness', '2px');
    }
  }

  if (spec.indentLeftDxa !== undefined) push('margin-left', dxaToIn(spec.indentLeftDxa));
  if (spec.indentRightDxa !== undefined) push('margin-right', dxaToIn(spec.indentRightDxa));
  if (spec.spaceBeforePt !== undefined) push('margin-top', `${spec.spaceBeforePt}pt`);
  if (spec.spaceAfterPt !== undefined) push('margin-bottom', `${spec.spaceAfterPt}pt`);

  if (spec.lineSpacing) {
    // w:line is 240ths of a line for the auto rule, otherwise twips
    const { rule, value } = spec.lineSpacing;
    push('line-height', rule === 'auto' ? (value / 240).toFixed(3) : `${value / 20}pt`);
  }

  return out;
}

/** cardmirror keys parts of its own chrome off these, so they are kept in
 *  step with the rules rather than left saying something else. */
function variables(profile: Profile): string {
  const t = profile.types;
  const pairs: [string, string | undefined][] = [
    [CSS_VAR.sizeNormal, t.paragraph.sizePt ? `${t.paragraph.sizePt}pt` : undefined],
    [CSS_VAR.sizePocket, t.pocket.sizePt ? `${t.pocket.sizePt}pt` : undefined],
    [CSS_VAR.sizeHat, t.hat.sizePt ? `${t.hat.sizePt}pt` : undefined],
    [CSS_VAR.sizeBlock, t.block.sizePt ? `${t.block.sizePt}pt` : undefined],
    [CSS_VAR.sizeTag, t.tag.sizePt ? `${t.tag.sizePt}pt` : undefined],
    [CSS_VAR.sizeAnalytic, t.analytic.sizePt ? `${t.analytic.sizePt}pt` : undefined],
    [CSS_VAR.sizeCite, t.cite_paragraph.sizePt ? `${t.cite_paragraph.sizePt}pt` : undefined],
    [CSS_VAR.sizeUnderline, t.underline_mark.sizePt ? `${t.underline_mark.sizePt}pt` : undefined],
    [CSS_VAR.sizeUndertag, t.undertag.sizePt ? `${t.undertag.sizePt}pt` : undefined],
    [CSS_VAR.bodyFont, fontStack(t.card_body, profile) ?? undefined],
  ];
  const body = pairs
    .filter((p): p is [string, string] => p[1] !== undefined)
    .map(([name, value]) => `  ${name}: ${value};`)
    .join('\n');
  return `${EDITOR_SELECTOR} {\n${body}\n}`;
}

export function toCss(profile: Profile): string {
  const blocks: string[] = [
    `/* laymirror — ${profile.name} */`,
    variables(profile),
  ];

  for (const [type, selector] of Object.entries(SELECTOR) as [BlockType | RunType, string][]) {
    const decls = declarations(profile.types[type], profile);
    if (decls.length === 0) continue;
    // :is() keeps the editor scope without inflating specificity per selector
    blocks.push(`:is(${EDITOR_SELECTOR}) ${selector} {\n${decls.join('\n')}\n}`);
  }

  return blocks.join('\n\n') + '\n';
}

export function applyStylesheet(css: string): void {
  let el = document.getElementById(STYLE_ID) as HTMLStyleElement | null;
  if (!el) {
    el = document.createElement('style');
    el.id = STYLE_ID;
    document.head.appendChild(el);
  }
  el.textContent = css;
}

export function hasStylesheet(): boolean {
  return document.getElementById(STYLE_ID) !== null;
}

/** the off-state must leave nothing behind. every match, not the first —
 *  a reloaded plugin can find one it did not put there. */
export function removeStylesheet(): void {
  for (const el of document.querySelectorAll(`#${STYLE_ID}`)) el.remove();
}
