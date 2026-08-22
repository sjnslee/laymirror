// work view typography.
//
// the editor gets the school's look by borrowing the same style resolver page
// view uses. a probe document is built from the template's own styles.xml and
// theme, rendered through docx-preview, and the stylesheet it emits is
// harvested and re-pointed at cardmirror's classes.
//
// the point is that there is exactly one resolver. the previous build parsed
// the donor into its own model for the editor and again for the file, so the
// two could — and did — disagree: `basedOn` chains were never followed and
// `majorHAnsi` never resolved, which is why the hat came out at the wrong
// size with the wrong face.

import { renderAsync } from 'docx-preview';
import { TYPE_BY_EXPORT_STYLE, type BlockType, type RunType } from '../profile/mapping.js';
import type { Profile } from '../profile/profile.js';
import { toBlob, zip, writeText, type Parts } from '../docx/zip.js';

export const STYLE_ID = 'laymirror-style';

/** cardmirror's editor container, and the class page view wears so the same
 *  rules reach both. */
export const EDITOR_SCOPE = '#editor, .pmd-pane-editor';

/** cardmirror's class for each of its own block types. */
export const CLASS_BY_TYPE: Partial<Record<BlockType | RunType, string>> = {
  pocket: '.pmd-pocket',
  hat: '.pmd-hat',
  block: '.pmd-block',
  tag: '.pmd-tag',
  analytic: '.pmd-analytic',
  undertag: '.pmd-undertag',
  cite_paragraph: '.pmd-cite-para',
  card_body: '.pmd-card-body',
};

const WML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** docx-preview names a style's class from its lowercased id. */
const classFor = (styleId: string): string => `docx_${styleId.toLowerCase()}`;

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
  '</Types>';

const ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

const DOC_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
  '</Relationships>';

/** one paragraph per style we care about, so the renderer is obliged to emit
 *  a rule for each. */
export function probeDocument(styleIds: readonly string[]): string {
  const body = styleIds
    .map(
      (id) =>
        `<w:p><w:pPr><w:pStyle w:val="${id}"/></w:pPr><w:r><w:t>probe</w:t></w:r></w:p>`,
    )
    .join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WML}"><w:body>${body}<w:sectPr/></w:body></w:document>`
  );
}

function probePackage(profile: Profile, styleIds: readonly string[]): Uint8Array {
  const parts: Parts = {};
  writeText(parts, '[Content_Types].xml', CONTENT_TYPES);
  writeText(parts, '_rels/.rels', ROOT_RELS);
  writeText(parts, 'word/_rels/document.xml.rels', DOC_RELS);
  writeText(parts, 'word/document.xml', probeDocument(styleIds));

  const snapshot = profile.snapshot;
  writeText(parts, 'word/styles.xml', snapshot?.parts['word/styles.xml'] ?? '<w:styles/>');
  const theme = snapshot?.parts['word/theme/theme1.xml'];
  if (theme) writeText(parts, 'word/theme/theme1.xml', theme);
  return zip(parts);
}

/** template style id -> the cardmirror class its rules should land on. */
export function targetsFor(profile: Profile): Map<string, string> {
  const targets = new Map<string, string>();
  for (const [exportId, templateId] of Object.entries(profile.styleMap)) {
    const type = TYPE_BY_EXPORT_STYLE[exportId];
    const className = type ? CLASS_BY_TYPE[type] : undefined;
    if (className) targets.set(templateId, className);
  }
  for (const [type, styleId] of Object.entries(profile.bareStyles)) {
    const className = CLASS_BY_TYPE[type as BlockType];
    if (styleId && className) targets.set(styleId, className);
  }
  return targets;
}

interface Rule {
  selector: string;
  body: string;
}

export function parseRules(css: string): Rule[] {
  const rules: Rule[] = [];
  // the emitted sheet is flat — no at-rules wrap the style rules we want
  for (const match of css.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
    const selector = match[1]!.replace(/\s+/g, ' ').trim();
    const body = match[2]!.replace(/\s+/g, ' ').trim();
    if (selector && body && !selector.startsWith('@')) rules.push({ selector, body });
  }
  return rules;
}

/** re-point docx-preview's selectors at cardmirror's classes, and scope the
 *  result to the editor so nothing leaks into the rest of the app. */
export function retarget(css: string, targets: Map<string, string>): string {
  const out: string[] = [];

  // the theme variables the rules reference have to come along, or every
  // `var(--docx-majorHAnsi-font)` resolves to nothing
  const variables = /\.docx\s*\{([^}]*--docx-[^}]*)\}/.exec(css)?.[1];
  if (variables) out.push(`${EDITOR_SCOPE} { ${variables.replace(/\s+/g, ' ').trim()} }`);

  for (const rule of parseRules(css)) {
    for (const [styleId, className] of targets) {
      const token = classFor(styleId);
      if (!new RegExp(`\\.${token}\\b`).test(rule.selector)) continue;
      // a docx paragraph's run properties live on its spans; cardmirror puts
      // the text directly in the block, so both have to be covered
      const scoped = rule.selector.includes('span')
        ? `:is(${EDITOR_SCOPE}) ${className}, :is(${EDITOR_SCOPE}) ${className} span`
        : `:is(${EDITOR_SCOPE}) ${className}`;
      out.push(`${scoped} { ${rule.body} }`);
      break;
    }
  }

  return out.join('\n');
}

/** build the editor stylesheet for a profile. resolves to '' when the profile
 *  has no template, which is the honest answer — there is nothing to show. */
export async function toCss(profile: Profile): Promise<string> {
  const targets = targetsFor(profile);
  if (targets.size === 0 || !profile.snapshot) return '';

  const container = document.createElement('div');
  const styleContainer = document.createElement('div');
  try {
    await renderAsync(
      toBlob(probePackage(profile, [...targets.keys()])),
      container,
      styleContainer,
      { inWrapper: false, breakPages: false, renderHeaders: false, renderFooters: false },
    );
  } catch {
    // a template we cannot render is a template we cannot style from; work
    // view simply stays as cardmirror draws it
    return '';
  }

  return retarget(styleContainer.textContent ?? '', targets);
}

export function applyStylesheet(css: string): void {
  let sheet = document.getElementById(STYLE_ID) as HTMLStyleElement | null;
  if (!sheet) {
    sheet = document.createElement('style');
    sheet.id = STYLE_ID;
    document.head.appendChild(sheet);
  }
  sheet.textContent = css;
}

export const hasStylesheet = (): boolean => document.getElementById(STYLE_ID) !== null;

export function removeStylesheet(): void {
  document.getElementById(STYLE_ID)?.remove();
}
