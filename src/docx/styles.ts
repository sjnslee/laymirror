// profile -> word/styles.xml. the other half of `toCss`.
//
// the donor's own styles.xml is the base, so anything laymirror doesn't model
// (table styles, latent style config, list definitions) survives untouched.
// only the mapped styles are overwritten.

import { REQUIRED_FOR_NATIVE_PATH } from '../profile/mapping.js';
import type { BlockType, Profile, RunType, TypeSpec } from '../profile/profile.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const EMPTY_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w:styles xmlns:w="${W}"></w:styles>`;

const CHARACTER_TYPES = new Set<BlockType | RunType>([
  'underline_mark',
  'emphasis_mark',
  'cite_mark',
  'analytic_mark',
  'undertag_mark',
]);

function child(doc: Document, parent: Element, tag: string): Element {
  const existing = parent.getElementsByTagName(tag).item(0);
  if (existing && existing.parentNode === parent) return existing;
  const el = doc.createElementNS(W, tag);
  parent.appendChild(el);
  return el;
}

function setFlag(doc: Document, parent: Element, tag: string, on: boolean): void {
  const el = child(doc, parent, tag);
  if (on) el.removeAttribute('w:val');
  else el.setAttribute('w:val', '0');
}

function drop(parent: Element, tag: string): void {
  const el = parent.getElementsByTagName(tag).item(0);
  if (el && el.parentNode === parent) parent.removeChild(el);
}

/** write a TypeSpec onto a <w:style>, leaving properties the profile doesn't
 *  model alone. */
function applySpec(doc: Document, style: Element, spec: TypeSpec, isCharacter: boolean): void {
  const name = child(doc, style, 'w:name');
  name.setAttribute('w:val', spec.styleName);

  const rPr = child(doc, style, 'w:rPr');
  if (spec.font) {
    const fonts = child(doc, rPr, 'w:rFonts');
    for (const a of ['w:ascii', 'w:hAnsi']) fonts.setAttribute(a, spec.font);
    // an explicit face must win over any inherited theme reference
    for (const a of ['w:asciiTheme', 'w:hAnsiTheme']) fonts.removeAttribute(a);
  }
  if (spec.sizePt !== undefined) {
    const half = String(Math.round(spec.sizePt * 2));
    child(doc, rPr, 'w:sz').setAttribute('w:val', half);
    child(doc, rPr, 'w:szCs').setAttribute('w:val', half);
  }
  if (spec.bold !== undefined) setFlag(doc, rPr, 'w:b', spec.bold);
  if (spec.italic !== undefined) setFlag(doc, rPr, 'w:i', spec.italic);
  if (spec.smallCaps !== undefined) setFlag(doc, rPr, 'w:smallCaps', spec.smallCaps);
  if (spec.color) child(doc, rPr, 'w:color').setAttribute('w:val', spec.color);
  if (spec.underline !== undefined) {
    child(doc, rPr, 'w:u').setAttribute('w:val', spec.underline);
  }

  if (isCharacter) {
    drop(style, 'w:pPr');
    return;
  }

  const pPr = child(doc, style, 'w:pPr');
  if (spec.align) child(doc, pPr, 'w:jc').setAttribute('w:val', spec.align);

  if (spec.indentLeftDxa !== undefined || spec.indentRightDxa !== undefined) {
    const ind = child(doc, pPr, 'w:ind');
    if (spec.indentLeftDxa !== undefined) ind.setAttribute('w:left', String(spec.indentLeftDxa));
    if (spec.indentRightDxa !== undefined) ind.setAttribute('w:right', String(spec.indentRightDxa));
  }

  if (
    spec.spaceBeforePt !== undefined ||
    spec.spaceAfterPt !== undefined ||
    spec.lineSpacing !== undefined
  ) {
    const spacing = child(doc, pPr, 'w:spacing');
    if (spec.spaceBeforePt !== undefined) {
      spacing.setAttribute('w:before', String(Math.round(spec.spaceBeforePt * 20)));
    }
    if (spec.spaceAfterPt !== undefined) {
      spacing.setAttribute('w:after', String(Math.round(spec.spaceAfterPt * 20)));
    }
    if (spec.lineSpacing) {
      spacing.setAttribute('w:line', String(spec.lineSpacing.value));
      spacing.setAttribute('w:lineRule', spec.lineSpacing.rule);
    }
  }

  if (spec.pageBreakBefore !== undefined) {
    setFlag(doc, pPr, 'w:pageBreakBefore', spec.pageBreakBefore);
  }
  if (spec.keepNext !== undefined) setFlag(doc, pPr, 'w:keepNext', spec.keepNext);
  if (spec.keepLines !== undefined) setFlag(doc, pPr, 'w:keepLines', spec.keepLines);

  if (spec.outlineLevel !== undefined && spec.outlineLevel !== null) {
    child(doc, pPr, 'w:outlineLvl').setAttribute('w:val', String(spec.outlineLevel));
  }
}

function findStyle(doc: Document, styleId: string): Element | null {
  const styles = doc.getElementsByTagName('w:style');
  for (let i = 0; i < styles.length; i++) {
    const style = styles.item(i)!;
    if (style.getAttribute('w:styleId') === styleId) return style;
  }
  return null;
}

function ensureStyle(doc: Document, styleId: string, isCharacter: boolean): Element {
  const existing = findStyle(doc, styleId);
  if (existing) return existing;

  const style = doc.createElementNS(W, 'w:style');
  style.setAttribute('w:type', isCharacter ? 'character' : 'paragraph');
  style.setAttribute('w:customStyle', '1');
  style.setAttribute('w:styleId', styleId);
  doc.documentElement.appendChild(style);
  return style;
}

/** cardmirror only takes its native import path — the one that recognises
 *  cite and underline marks by styleId — when these three are present. they
 *  are inert in word, and without them a cite mark is dropped on reimport. */
function ensureNativePathSentinels(doc: Document): void {
  for (const group of REQUIRED_FOR_NATIVE_PATH) {
    const id = group[0]!;
    if (group.some((candidate) => findStyle(doc, candidate))) continue;
    const style = ensureStyle(doc, id, true);
    child(doc, style, 'w:name').setAttribute('w:val', id);
  }
}

export function buildStylesXml(profile: Profile): string {
  const source = profile.donorStylesXml.trim() || EMPTY_STYLES;
  const doc = new DOMParser().parseFromString(source, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error('donor styles.xml did not parse');
  }

  for (const key of Object.keys(profile.types) as (BlockType | RunType)[]) {
    const spec = profile.types[key];
    const isCharacter = CHARACTER_TYPES.has(key);
    applySpec(doc, ensureStyle(doc, spec.styleId, isCharacter), spec, isCharacter);
  }

  ensureNativePathSentinels(doc);

  return new XMLSerializer().serializeToString(doc);
}
