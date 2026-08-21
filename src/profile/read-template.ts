// a school's template docx IS the profile. it already carries font, size,
// weight, casing, underline, colour, alignment, indent, spacing and
// page-break-before per style, plus page setup and the header and footer.
// so nothing here invents a format — it reads one.

import { unzip, readText, type Parts } from '../docx/zip.js';
import { HEADING_LEVEL_TO_TYPE } from './mapping.js';
import { FONT_FALLBACKS } from './defaults.js';
import type { BlockType, PageSetup, Profile, RunType, TypeSpec } from './profile.js';

const STYLES = 'word/styles.xml';
const THEME = 'word/theme/theme1.xml';
const DOCUMENT = 'word/document.xml';
const SETTINGS_RELS = 'word/_rels/settings.xml.rels';

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

const first = (el: Element | Document, tag: string): Element | null =>
  el.getElementsByTagName(tag).item(0);

const attr = (el: Element | null, name: string): string | null =>
  el?.getAttribute(name) ?? null;

/** ooxml booleans: present means true unless w:val says otherwise. */
function onOff(parent: Element | null, tag: string): boolean | undefined {
  if (!parent) return undefined;
  const el = first(parent, tag);
  if (!el) return undefined;
  const val = el.getAttribute('w:val');
  return val === null ? true : val !== '0' && val !== 'false';
}

const num = (value: string | null): number | undefined => {
  if (value === null) return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
};

interface RawStyle {
  id: string;
  name: string;
  basedOn: string | null;
  type: string;
  spec: Partial<TypeSpec>;
}

function themeFonts(parts: Parts): { major: string | null; minor: string | null } {
  const xml = readText(parts, THEME);
  if (!xml) return { major: null, minor: null };
  const doc = parseXml(xml);
  const read = (tag: string): string | null =>
    attr(first(first(doc, tag) ?? doc, 'a:latin'), 'typeface');
  return { major: read('a:majorFont'), minor: read('a:minorFont') };
}

function readFont(rPr: Element | null, theme: { major: string | null; minor: string | null }) {
  const fonts = rPr ? first(rPr, 'w:rFonts') : null;
  if (!fonts) return undefined;
  const explicit = attr(fonts, 'w:ascii');
  if (explicit) return explicit;
  // asciiTheme="minorHAnsi" | "majorHAnsi" resolves through the theme
  const themed = attr(fonts, 'w:asciiTheme');
  if (!themed) return undefined;
  return (themed.startsWith('major') ? theme.major : theme.minor) ?? undefined;
}

function readStyleSpec(
  style: Element,
  theme: { major: string | null; minor: string | null },
): Partial<TypeSpec> {
  const pPr = first(style, 'w:pPr');
  const rPr = first(style, 'w:rPr');
  const spec: Partial<TypeSpec> = {};

  const font = readFont(rPr, theme);
  if (font) spec.font = font;

  // w:sz is half-points
  const sz = num(attr(rPr ? first(rPr, 'w:sz') : null, 'w:val'));
  if (sz !== undefined) spec.sizePt = sz / 2;

  const bold = onOff(rPr, 'w:b');
  if (bold !== undefined) spec.bold = bold;
  const italic = onOff(rPr, 'w:i');
  if (italic !== undefined) spec.italic = italic;
  const smallCaps = onOff(rPr, 'w:smallCaps');
  if (smallCaps !== undefined) spec.smallCaps = smallCaps;

  const u = attr(rPr ? first(rPr, 'w:u') : null, 'w:val');
  if (u) spec.underline = u as TypeSpec['underline'];

  const color = attr(rPr ? first(rPr, 'w:color') : null, 'w:val');
  if (color && color !== 'auto') spec.color = color;

  const jc = attr(pPr ? first(pPr, 'w:jc') : null, 'w:val');
  if (jc) spec.align = jc as TypeSpec['align'];

  const ind = pPr ? first(pPr, 'w:ind') : null;
  const left = num(attr(ind, 'w:left'));
  if (left !== undefined) spec.indentLeftDxa = left;
  const right = num(attr(ind, 'w:right'));
  if (right !== undefined) spec.indentRightDxa = right;

  const spacing = pPr ? first(pPr, 'w:spacing') : null;
  // w:before / w:after are twips; 20 to the point
  const before = num(attr(spacing, 'w:before'));
  if (before !== undefined) spec.spaceBeforePt = before / 20;
  const after = num(attr(spacing, 'w:after'));
  if (after !== undefined) spec.spaceAfterPt = after / 20;
  const line = num(attr(spacing, 'w:line'));
  if (line !== undefined) {
    const rule = (attr(spacing, 'w:lineRule') ?? 'auto') as 'auto' | 'exact' | 'atLeast';
    spec.lineSpacing = { rule, value: line };
  }

  const pageBreakBefore = onOff(pPr, 'w:pageBreakBefore');
  if (pageBreakBefore !== undefined) spec.pageBreakBefore = pageBreakBefore;
  const keepNext = onOff(pPr, 'w:keepNext');
  if (keepNext !== undefined) spec.keepNext = keepNext;
  const keepLines = onOff(pPr, 'w:keepLines');
  if (keepLines !== undefined) spec.keepLines = keepLines;

  const outline = num(attr(pPr ? first(pPr, 'w:outlineLvl') : null, 'w:val'));
  if (outline !== undefined) spec.outlineLevel = outline;

  return spec;
}

function readStyles(parts: Parts): Map<string, RawStyle> {
  const xml = readText(parts, STYLES);
  const out = new Map<string, RawStyle>();
  if (!xml) return out;

  const doc = parseXml(xml);
  const theme = themeFonts(parts);

  // docDefaults become the implicit base of every style
  const defaults = first(doc, 'w:rPrDefault');
  if (defaults) {
    out.set('__defaults__', {
      id: '__defaults__',
      name: '__defaults__',
      basedOn: null,
      type: 'paragraph',
      spec: readStyleSpec(defaults, theme),
    });
  }

  const styles = doc.getElementsByTagName('w:style');
  for (let i = 0; i < styles.length; i++) {
    const style = styles.item(i)!;
    const id = style.getAttribute('w:styleId');
    if (!id) continue;
    out.set(id, {
      id,
      name: attr(first(style, 'w:name'), 'w:val') ?? id,
      basedOn: attr(first(style, 'w:basedOn'), 'w:val'),
      type: style.getAttribute('w:type') ?? 'paragraph',
      spec: readStyleSpec(style, theme),
    });
  }
  return out;
}

/** follow basedOn to the root, then merge back down so the nearest style
 *  wins. word resolves inheritance this way and css needs the result. */
function resolve(id: string, styles: Map<string, RawStyle>): TypeSpec | null {
  const chain: RawStyle[] = [];
  const seen = new Set<string>();
  let cursor: string | null = id;
  while (cursor && !seen.has(cursor)) {
    seen.add(cursor);
    const style: RawStyle | undefined = styles.get(cursor);
    if (!style) break;
    chain.unshift(style);
    cursor = style.basedOn;
  }
  if (chain.length === 0) return null;

  const defaults = styles.get('__defaults__');
  const merged: Partial<TypeSpec> = { ...(defaults?.spec ?? {}) };
  for (const style of chain) Object.assign(merged, style.spec);

  const own = styles.get(id)!;
  return { ...merged, styleId: own.id, styleName: own.name };
}

/** find the donor's style for a cardmirror type: preferred ids first, then
 *  names, both case-insensitively. */
function pick(
  styles: Map<string, RawStyle>,
  ids: readonly string[],
  names: readonly string[] = [],
): string | null {
  for (const id of ids) if (styles.has(id)) return id;
  const wanted = new Set(names.map((n) => n.toLowerCase()));
  for (const style of styles.values()) {
    if (wanted.has(style.name.toLowerCase())) return style.id;
  }
  return null;
}

/** headings are identified by outline level, matching how the importer reads
 *  them back. */
function pickHeading(styles: Map<string, RawStyle>, level: number): string | null {
  for (const style of styles.values()) {
    if (style.type !== 'paragraph') continue;
    if (style.spec.outlineLevel === level) return style.id;
  }
  return pick(styles, [`Heading${level + 1}`], [`heading ${level + 1}`]);
}

function readPage(parts: Parts): PageSetup {
  const fallback: PageSetup = {
    widthTwips: 12240,
    heightTwips: 15840,
    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 },
  };
  const xml = readText(parts, DOCUMENT);
  if (!xml) return fallback;
  const sect = first(parseXml(xml), 'w:sectPr');
  if (!sect) return fallback;

  const size = first(sect, 'w:pgSz');
  const mar = first(sect, 'w:pgMar');
  return {
    widthTwips: num(attr(size, 'w:w')) ?? fallback.widthTwips,
    heightTwips: num(attr(size, 'w:h')) ?? fallback.heightTwips,
    margin: {
      top: num(attr(mar, 'w:top')) ?? fallback.margin.top,
      right: num(attr(mar, 'w:right')) ?? fallback.margin.right,
      bottom: num(attr(mar, 'w:bottom')) ?? fallback.margin.bottom,
      left: num(attr(mar, 'w:left')) ?? fallback.margin.left,
      header: num(attr(mar, 'w:header')) ?? fallback.margin.header,
      footer: num(attr(mar, 'w:footer')) ?? fallback.margin.footer,
    },
  };
}

/** donors carry an absolute path through the author's home directory. word
 *  only basename-matches, so the basename is both the safe answer and the
 *  correct one. */
export function readAttachedTemplate(parts: Parts): string | null {
  const xml = readText(parts, SETTINGS_RELS);
  if (!xml) return null;
  const rels = parseXml(xml).getElementsByTagName('Relationship');
  for (let i = 0; i < rels.length; i++) {
    const rel = rels.item(i)!;
    if (!rel.getAttribute('Type')?.endsWith('/attachedTemplate')) continue;
    const target = rel.getAttribute('Target');
    if (!target) continue;
    const base = target.split(/[\\/]/).pop() ?? target;
    try {
      return decodeURIComponent(base);
    } catch {
      return base;
    }
  }
  return null;
}

const WANTED: Record<BlockType | RunType, { ids: readonly string[]; names: readonly string[] }> = {
  paragraph: { ids: ['Normal'], names: ['normal'] },
  pocket: { ids: [], names: [] },
  hat: { ids: [], names: [] },
  block: { ids: [], names: [] },
  tag: { ids: ['Tag', 'Tags'], names: ['tag', 'tags', 'debate tag'] },
  cite_paragraph: { ids: ['Cite', 'Cites'], names: ['cite', 'cites', 'debate cite main'] },
  card_body: { ids: ['card', 'Cards'], names: ['card', 'cards', 'card text'] },
  analytic: { ids: ['Analytic'], names: ['analytic'] },
  undertag: { ids: ['Undertag'], names: ['undertag'] },
  underline_mark: { ids: ['Underline', 'StyleUnderline'], names: ['underline', 'style underline'] },
  cite_mark: { ids: ['Style13ptBold', 'StyleStyleBold12pt'], names: ['style 13 pt bold'] },
  emphasis_mark: { ids: ['Emphasis'], names: ['emphasis'] },
  analytic_mark: { ids: ['AnalyticChar'], names: [] },
  undertag_mark: { ids: ['UndertagChar'], names: [] },
};

export interface TemplateResult {
  profile: Profile;
  /** types the donor had no style for; they keep the fallback profile's. */
  missing: (BlockType | RunType)[];
}

export function readTemplate(bytes: Uint8Array, fallback: Profile): TemplateResult {
  const parts = unzip(bytes);
  const styles = readStyles(parts);
  const missing: (BlockType | RunType)[] = [];

  const types = { ...fallback.types };
  for (const key of Object.keys(WANTED) as (BlockType | RunType)[]) {
    const heading = key === 'pocket' ? 0 : key === 'hat' ? 1 : key === 'block' ? 2 : null;
    const id =
      heading !== null
        ? pickHeading(styles, heading)
        : pick(styles, WANTED[key].ids, WANTED[key].names);

    const resolved = id ? resolve(id, styles) : null;
    if (!resolved) {
      missing.push(key);
      continue;
    }
    types[key] = resolved;
  }

  return {
    profile: {
      ...fallback,
      types,
      page: readPage(parts),
      headerXml: readText(parts, 'word/header1.xml'),
      footerXml: readText(parts, 'word/footer1.xml'),
      attachedTemplate: readAttachedTemplate(parts) ?? fallback.attachedTemplate,
      donorStylesXml: readText(parts, STYLES) ?? '',
      fontFallbacks: { ...FONT_FALLBACKS, ...fallback.fontFallbacks },
    },
    missing,
  };
}

/** the headings the importer will reconstruct, for the settings ui. */
export const HEADING_TYPES = HEADING_LEVEL_TO_TYPE;
