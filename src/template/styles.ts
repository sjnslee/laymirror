// the cardmirror <-> template vocabulary bridge, and the one place that knows
// whether a mapping survives reimport.
//
// cardmirror's importer has two paths, and which one runs decides everything:
//
//   native   — taken when the document's styles look like cardmirror's own.
//              marks are matched by styleId.
//   legacy   — taken otherwise. paragraph styles are matched by lowercased
//              w:name, character styles only by the id table below.
//
// the lay names (Tag, Cite, card, Underline) are all in the legacy table, which
// is what makes them safe.
//
// read off the shipped 1.3.0 parse worker; the tables still match 1.11.0.

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

/** lowercased w:name -> legacy role.
 *
 *  this table and `LEGACY_BY_ID` are taken from cardmirror's
 *  `src/ooxml/legacy-styles.ts` (`BY_NAME` and `BY_ID`). they have to match it
 *  entry for entry or a template maps to styles the importer will not read
 *  back, so they are copied rather than paraphrased. see LICENSE. */
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

/** styleId -> legacy role, consulted only after the name misses. cardmirror's
 *  `BY_ID`; see the note on `LEGACY_BY_NAME`. */
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

/** what cardmirror's *exporter* writes for each type — not the same question as
 *  what its importer reads back. `null` means no `w:pStyle` at all, so a cite
 *  paragraph and a card body have to be recognised from their runs instead.
 *  note `tag` leaves as `Heading4`, not as any style named "tag".
 *
 *  transposed from cardmirror's `STYLE_RENAME_MAP` in
 *  `src/ooxml/style-clean/template-styles.ts`; see LICENSE. */
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

export interface StyleInfo {
  id: string;
  name: string;
  kind: 'paragraph' | 'character' | 'table' | 'numbering';
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
    });
  }
  return out;
}

/** reading cardmirror's own legacy tables picks the mapping it will agree with
 *  when the file comes back. */
const roleOf = (style: StyleInfo): string | null =>
  LEGACY_BY_NAME[style.name.toLowerCase()] ?? LEGACY_BY_ID[style.id] ?? null;

/** the role each export style wants a home for, and the kind it must land on:
 *  a run style on a paragraph style writes an `rStyle` word cannot resolve. */
const WANTED: Record<string, { role: string; kind: StyleInfo['kind'] }> = {
  Heading4: { role: 'tag', kind: 'paragraph' },
  Style13ptBold: { role: 'char-cite', kind: 'character' },
  StyleUnderline: { role: 'char-underline', kind: 'character' },
};

/** every style playing each role, in definition order — all of them, because
 *  cardmirror's own `Heading4` is named "heading 4" and its legacy table reads
 *  that as a tag too, so a template's `Tag` would lose to it. */
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

/** every `<w:style>` in a part, by the id it defines. */
function blocksById(stylesXml: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const match of stylesXml.matchAll(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g)) {
    const id = attr(/^<[^>]*>/.exec(match[0])![0], 'w:styleId');
    if (id) out.set(id, match[0]);
  }
  return out;
}

const REFERENCED = /<w:(?:pStyle|rStyle|tblStyle)\b[^>]*\bw:val="([^"]*)"/g;
const INHERITED = /<w:(?:basedOn|link|next)\b[^>]*\bw:val="([^"]*)"/g;

/** the template's `styles.xml` lands whole on cardmirror's, so an id the
 *  document still names and the template does not define resolves to nothing
 *  and word renders that text as Normal. cardmirror's own definition is
 *  carried in beside the template's instead.
 *
 *  null when there is nothing to carry: the template's part stands as it is. */
export function mergeStyles(
  exportXml: string | null,
  templateXml: string | null,
  documentXml: string,
): string | null {
  if (!exportXml || !templateXml || !templateXml.includes('</w:styles>')) return null;

  const defined = new Set(readStyles(templateXml).map((style) => style.id));
  const available = blocksById(exportXml);
  const carried = new Map<string, string>();

  const wanted = [...documentXml.matchAll(REFERENCED)].map((match) => match[1]!);
  for (let id = wanted.shift(); id !== undefined; id = wanted.shift()) {
    if (defined.has(id) || carried.has(id)) continue;
    const block = available.get(id);
    if (!block) continue;
    carried.set(id, block);
    // a style based on or linked to one nothing else names is half a definition
    for (const link of block.matchAll(INHERITED)) wanted.push(link[1]!);
  }

  if (carried.size === 0) return null;
  return templateXml.replace('</w:styles>', `${[...carried.values()].join('')}</w:styles>`);
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
 *  identity is the default and usually right for Heading1-3. it is wrong for a
 *  tag: cardmirror exports one as `Heading4`, which word renders in its stock
 *  italic blue. so a template style whose role matches wins. */
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
    // never cardmirror's own: a candidate equal to the id being remapped is
    // the thing the remap exists to get away from
    const preferred = wanted
      ? (byKind[wanted.kind as 'paragraph' | 'character'] ?? byKind.paragraph)
          .get(wanted.role)
          ?.find((id) => id !== exportId)
      : undefined;
    // never map onto a style the template does not define: the id is left as
    // cardmirror wrote it, and `mergeStyles` carries its definition over
    if (preferred && preferred !== exportId) map[exportId] = preferred;
    else if (defined.has(exportId)) map[exportId] = exportId;
  }
  return map;
}
