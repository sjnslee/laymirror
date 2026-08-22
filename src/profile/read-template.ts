// turning a school's .docx or .dotx into a profile.
//
// this used to parse the donor's typography into a model and re-emit it,
// which dropped every property nobody thought to parse. now it only does two
// things: take the identity snapshot verbatim, and work out which of the
// template's style ids each cardmirror export style should become.

import { captureSnapshot } from '../docx/snapshot.js';
import { isDocx, readText, unzip } from '../docx/zip.js';
import { EXPORT_STYLE_BY_TYPE, LEGACY_BY_ID, LEGACY_BY_NAME } from './mapping.js';
import type { Profile, StyleInfo } from './profile.js';

const STYLES = 'word/styles.xml';

export type ReadResult =
  | { ok: true; profile: Profile }
  | { ok: false; error: string };

const attr = (tag: string, name: string): string | null =>
  new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1] ?? null;

export function readStyles(stylesXml: string): StyleInfo[] {
  const out: StyleInfo[] = [];
  for (const match of stylesXml.matchAll(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g)) {
    const block = match[0];
    const open = /<w:style\b[^>]*>/.exec(block)![0];
    const id = attr(open, 'w:styleId');
    if (!id) continue;
    const kind = (attr(open, 'w:type') ?? 'paragraph') as StyleInfo['kind'];
    const name = /<w:name\b[^>]*w:val="([^"]*)"/.exec(block)?.[1] ?? id;
    out.push({ id, name, kind });
  }
  return out;
}

/** cardmirror's own legacy tables say which of a school's styles means what:
 *  it matches paragraph styles by lowercased `w:name` and character styles by
 *  a small id table. reading the same tables here means the mapping we pick
 *  is the one cardmirror will agree with when the file comes back. */
function roleOf(style: StyleInfo): string | null {
  return LEGACY_BY_NAME[style.name.toLowerCase()] ?? LEGACY_BY_ID[style.id] ?? null;
}

/** the role each cardmirror export style is looking for a home for, and the
 *  kind of style it must land on. a run style mapped onto a paragraph style
 *  would be written as an `rStyle` word cannot resolve, so the kinds are
 *  checked rather than assumed. */
const WANTED: Record<string, { role: string; kind: StyleInfo['kind'] }> = {
  Heading4: { role: 'tag', kind: 'paragraph' },
  Style13ptBold: { role: 'char-cite', kind: 'character' },
  StyleUnderline: { role: 'char-underline', kind: 'character' },
};

/** every style playing each role, in definition order.
 *
 *  all of them, not just the first: cardmirror's own `Heading4` is named
 *  "heading 4", which its legacy table also reads as a tag — so a template's
 *  `Tag` would lose to the very style we are trying to move away from. */
function rolesIn(styles: StyleInfo[], kind?: StyleInfo['kind']): Map<string, string[]> {
  const byRole = new Map<string, string[]>();
  for (const style of styles) {
    if (kind && style.kind !== kind) continue;
    const role = roleOf(style);
    if (!role) continue;
    byRole.set(role, [...(byRole.get(role) ?? []), style.id]);
  }
  return byRole;
}

/** the two paragraph types cardmirror exports with no style at all. */
export function deriveBareStyles(styles: StyleInfo[]): {
  cite_paragraph: string | null;
  card_body: string | null;
} {
  const byRole = rolesIn(styles, 'paragraph');
  return {
    cite_paragraph: byRole.get('cite')?.[0] ?? null,
    card_body: byRole.get('body')?.[0] ?? null,
  };
}

/** map cardmirror's export ids onto the template's own.
 *
 *  identity is the default and is usually right for Heading1–3. it is wrong
 *  for a tag: cardmirror exports one as `Heading4`, but a lay template's tag
 *  style is its own `Tag`, and leaving identity would render every tag in
 *  word's stock italic blue. so where the template defines a style whose role
 *  matches, that wins. */
export function deriveStyleMap(styles: StyleInfo[]): Record<string, string> {
  const defined = new Set(styles.map((s) => s.id));
  const byKind = {
    paragraph: rolesIn(styles, 'paragraph'),
    character: rolesIn(styles, 'character'),
  };

  const map: Record<string, string> = {};
  for (const exportId of Object.values(EXPORT_STYLE_BY_TYPE)) {
    if (!exportId) continue;
    const wanted = WANTED[exportId];
    // the school's own style, never cardmirror's — a candidate equal to the
    // id we are remapping is the thing we are trying to get away from
    const preferred = wanted
      ? (byKind[wanted.kind as 'paragraph' | 'character'] ?? byKind.paragraph)
          .get(wanted.role)
          ?.find((id) => id !== exportId)
      : undefined;
    // never map onto a style the template does not define — word would show
    // the text unstyled, which is worse than cardmirror's own default
    if (preferred && preferred !== exportId) map[exportId] = preferred;
    else if (defined.has(exportId)) map[exportId] = exportId;
  }
  return map;
}

export function readTemplate(bytes: Uint8Array, name: string): ReadResult {
  let parts;
  try {
    parts = unzip(bytes);
  } catch {
    return { ok: false, error: `could not read ${name} — is it a word document or template?` };
  }
  if (!isDocx(parts)) {
    return { ok: false, error: `${name} is not a word document or template` };
  }

  const snapshot = captureSnapshot(parts);
  const stylesXml = readText(parts, STYLES);
  const styles = stylesXml ? readStyles(stylesXml) : [];

  if (!snapshot && styles.length === 0) {
    return { ok: false, error: `${name} has no styles or header to copy` };
  }

  return {
    ok: true,
    profile: {
      // one profile per template, so two schools' profiles cannot collide
      id: `template:${name}`,
      name,
      snapshot,
      styleMap: deriveStyleMap(styles),
      bareStyles: deriveBareStyles(styles),
      styles,
    },
  };
}
