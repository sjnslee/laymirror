// the cardmirror <-> template vocabulary bridge, and the one place that knows
// whether a mapping survives reimport.
//
// cardmirror's importer has two paths, and which one runs decides everything:
//
//   native   — taken when the document's styles look like cardmirror's own.
//              marks are matched by styleId, so `Style13ptBold` comes back as
//              a cite mark and `Underline` as an underline mark.
//   legacy   — taken otherwise. paragraph styles are matched by lowercased
//              w:name, and character styles only by the small table below.
//
// the lay donor names (Tag, Cite, card, Underline) are all in the legacy
// table, which is what makes lay style names safe. `Style13ptBold` is NOT —
// under the legacy path a cite mark is silently lost. see
// REQUIRED_FOR_NATIVE_PATH.
//
// verified against the shipped 1.3.0 parse worker.



/** cardmirror's block vocabulary — the node types its exporter can emit. */
export type BlockType =
  | 'pocket'
  | 'hat'
  | 'block'
  | 'tag'
  | 'analytic'
  | 'undertag'
  | 'cite_paragraph'
  | 'card_body'
  | 'paragraph';

/** cardmirror's run vocabulary. */
export type RunType =
  | 'underline_mark'
  | 'emphasis_mark'
  | 'cite_mark'
  | 'analytic_mark'
  | 'undertag_mark';

/** lowercased w:name -> legacy role. */
export const LEGACY_BY_NAME: Record<string, string> = {
  tags: 'tag',
  tag: 'tag',
  'debate tag': 'tag',
  'heading 4': 'tag',
  'block headings': 'heading',
  'block heading': 'heading',
  'block title': 'heading',
  'hidden block header': 'heading',
  'heading 1': 'heading',
  'heading 2': 'heading',
  'heading 3': 'heading',
  cites: 'cite',
  cite: 'cite',
  'debate cite main': 'cite',
  'debate secondary cite': 'cite',
  normalcite: 'cite',
  cards: 'body',
  card: 'body',
  'card text': 'body',
  'card (indented)': 'body',
  nothing: 'body',
  'normal text': 'body',
  'evidence text': 'body',
  'author-date': 'char-cite',
  'debate underline': 'char-underline',
  'debate highlighted': 'char-underline',
  underline: 'char-underline',
  'dotted underline': 'char-underline',
  'style bold underline': 'char-underline',
  'style style bold + 12 pt': 'char-cite',
};

/** styleId -> legacy role, consulted only after the name misses. */
export const LEGACY_BY_ID: Record<string, string> = {
  Tags: 'tag',
  BlockHeadings: 'heading',
  BlockTitle: 'heading',
  Cites: 'cite',
  Cards: 'body',
  Nothing: 'body',
  'Author-Date': 'char-cite',
  DebateUnderline: 'char-underline',
  DebateHighlighted: 'char-underline',
  DottedUnderline: 'char-underline',
  StyleBoldUnderline: 'char-underline',
  StyleStyleBold12pt: 'char-cite',
};

/** styleId -> mark, on the native path. */
export const NATIVE_MARK_BY_ID: Record<string, RunType> = {
  StyleUnderline: 'underline_mark',
  Underline: 'underline_mark',
  StyleBoldUnderline: 'underline_mark',
  Style13ptBold: 'cite_mark',
  StyleStyleBold12pt: 'cite_mark',
  Cite: 'cite_mark',
  Emphasis: 'emphasis_mark',
  UndertagChar: 'undertag_mark',
  AnalyticChar: 'analytic_mark',
};

/** what cardmirror's *exporter* writes for each type, which is not the same
 *  question as what its importer reads back. `null` means it writes no
 *  `w:pStyle` at all — a cite paragraph and a card body leave the editor as
 *  bare paragraphs, indistinguishable from body text by style alone, which is
 *  why `rewrite.ts` has to recognise them from their runs.
 *
 *  note `tag` leaves as `Heading4`, not as any style named "tag".
 *
 *  verified against the shipped 1.3.0 exporter. */
export const EXPORT_STYLE_BY_TYPE: Record<BlockType | RunType, string | null> = {
  pocket: 'Heading1',
  hat: 'Heading2',
  block: 'Heading3',
  tag: 'Heading4',
  analytic: 'Analytic',
  undertag: 'Undertag',
  cite_paragraph: null,
  card_body: null,
  paragraph: null,
  cite_mark: 'Style13ptBold',
  underline_mark: 'StyleUnderline',
  emphasis_mark: 'Emphasis',
  undertag_mark: 'UndertagChar',
  analytic_mark: 'AnalyticChar',
};

/** the same table read the other way, for turning an export back into types. */
export const TYPE_BY_EXPORT_STYLE: Record<string, BlockType | RunType> = Object.fromEntries(
  Object.entries(EXPORT_STYLE_BY_TYPE)
    .filter((entry): entry is [BlockType | RunType, string] => entry[1] !== null)
    .map(([type, styleId]) => [styleId, type]),
);

/** cardmirror decides a document is one of its own when its styles contain
 *  all three of these, matched by id OR name. emitting them is what keeps the
 *  native path — and therefore cite and underline marks — alive through a
 *  round-trip. they cost nothing in word. */
export const REQUIRED_FOR_NATIVE_PATH: readonly (readonly string[])[] = [
  ['Style13ptBold', 'Style 13 pt Bold'],
  ['StyleUnderline', 'Style Underline'],
  ['Emphasis'],
];

export function takesNativePath(styleIds: Iterable<string>, styleNames: Iterable<string>): boolean {
  const ids = new Set(styleIds);
  const names = new Set(styleNames);
  return REQUIRED_FOR_NATIVE_PATH.every((group) =>
    group.some((s) => ids.has(s) || names.has(s)),
  );
}

export interface MappingWarning {
  styleId: string;
  styleName: string;
  message: string;
}

/** warn where a mapping will not survive the trip home.
 *
 *  cardmirror reads a reopened document one of two ways. it takes the native
 *  path only when the styles look like its own, and otherwise falls back to
 *  matching paragraph styles by lowercased `w:name` and character styles by a
 *  small id table. so a template style that is in neither table exports
 *  perfectly into word and comes back as an ordinary paragraph — which is the
 *  failure worth telling the user about before they cut a whole file. */
export function validateMapping(
  styles: readonly StyleInfo[],
  styleMap: Record<string, string>,
  bareStyles: BareStyles,
): MappingWarning[] {
  const byId = new Map(styles.map((style) => [style.id, style]));
  const targets = [
    ...Object.values(styleMap),
    bareStyles.cite_paragraph,
    bareStyles.card_body,
  ].filter((id): id is string => !!id);

  // on the native path cardmirror matches paragraph styles by id, so any id
  // its own exporter emits comes back as the type it left as. the legacy
  // tables are the fallback, and `Analytic` is in neither of them.
  const native = takesNativePath(
    styles.map((style) => style.id),
    styles.map((style) => style.name),
  );

  const warnings: MappingWarning[] = [];
  const seen = new Set<string>();

  for (const id of targets) {
    if (seen.has(id)) continue;
    seen.add(id);
    const style = byId.get(id);
    const name = style?.name ?? id;
    // headings resolve by outline level rather than by name, so they are safe
    if (/^Heading\d$/.test(id)) continue;
    if (native && TYPE_BY_EXPORT_STYLE[id]) continue;
    if (LEGACY_BY_NAME[name.toLowerCase()] || LEGACY_BY_ID[id]) continue;
    if (NATIVE_MARK_BY_ID[id]) continue;
    warnings.push({
      styleId: id,
      styleName: name,
      message: `cardmirror does not recognise a style called "${name}", so this text comes back as an ordinary paragraph when the file is reopened`,
    });
  }

  return warnings;
}

// ── reading a template's own styles ───────────────────────────────────

export interface StyleInfo {
  id: string;
  name: string;
  kind: 'paragraph' | 'character' | 'table' | 'numbering';
  basedOn: string | null;
}

const attr = (tag: string, name: string): string | null =>
  new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1] ?? null;

export function readStyles(stylesXml: string): StyleInfo[] {
  const out: StyleInfo[] = [];
  for (const match of stylesXml.matchAll(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g)) {
    const block = match[0];
    const open = /<w:style\b[^>]*>/.exec(block)![0];
    const id = attr(open, 'w:styleId');
    if (!id) continue;
    out.push({
      id,
      name: /<w:name\b[^>]*w:val="([^"]*)"/.exec(block)?.[1] ?? id,
      kind: (attr(open, 'w:type') ?? 'paragraph') as StyleInfo['kind'],
      basedOn: /<w:basedOn\b[^>]*w:val="([^"]*)"/.exec(block)?.[1] ?? null,
    });
  }
  return out;
}

/** cardmirror's own legacy tables say which of a school's styles means what:
 *  it matches paragraph styles by lowercased `w:name` and character styles by
 *  a small id table. reading the same tables here means the mapping we pick is
 *  the one cardmirror will agree with when the file comes back. */
const roleOf = (style: StyleInfo): string | null =>
  LEGACY_BY_NAME[style.name.toLowerCase()] ?? LEGACY_BY_ID[style.id] ?? null;

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
function rolesIn(styles: readonly StyleInfo[], kind?: StyleInfo['kind']): Map<string, string[]> {
  const byRole = new Map<string, string[]>();
  for (const style of styles) {
    if (kind && style.kind !== kind) continue;
    const role = roleOf(style);
    if (!role) continue;
    byRole.set(role, [...(byRole.get(role) ?? []), style.id]);
  }
  return byRole;
}

export interface BareStyles {
  cite_paragraph: string | null;
  card_body: string | null;
}

/** the two paragraph types cardmirror exports with no style at all. */
export function deriveBareStyles(styles: readonly StyleInfo[]): BareStyles {
  const byRole = rolesIn(styles, 'paragraph');
  return {
    cite_paragraph: byRole.get('cite')?.[0] ?? null,
    card_body: byRole.get('body')?.[0] ?? null,
  };
}

/** map cardmirror's export ids onto the template's own.
 *
 *  identity is the default and is usually right for Heading1-3. it is wrong
 *  for a tag: cardmirror exports one as `Heading4`, but a template whose tag
 *  style is its own `Tag` would render every tag in word's stock italic blue.
 *  so where the template defines a style whose role matches, that wins. */
export function deriveStyleMap(styles: readonly StyleInfo[]): Record<string, string> {
  const defined = new Set(styles.map((style) => style.id));
  const byKind = {
    paragraph: rolesIn(styles, 'paragraph'),
    character: rolesIn(styles, 'character'),
  };

  const map: Record<string, string> = {};
  for (const exportId of Object.values(EXPORT_STYLE_BY_TYPE)) {
    if (!exportId) continue;
    const wanted = WANTED[exportId];
    // the school's own style, never cardmirror's — a candidate equal to the id
    // we are remapping is the thing we are trying to get away from
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
