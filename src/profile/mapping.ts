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

import type { Profile } from './profile.js';

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

/** what an earlier laymirror wrote into the document to stand in for a page
 *  break. breaks are held outside the document now — see docx/breaks.ts — and
 *  this survives only so an old document can be cleaned up on its next save. */
export const LEGACY_SENTINEL = '[page break]';

/** heading level (from w:outlineLvl + 1) -> block type. */
export const HEADING_LEVEL_TO_TYPE: Record<number, BlockType> = {
  1: 'pocket',
  2: 'hat',
  3: 'block',
  4: 'tag',
  5: 'block',
};

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
export function validateMapping(profile: Profile): MappingWarning[] {
  const byId = new Map(profile.styles.map((s) => [s.id, s]));
  const targets = [
    ...Object.values(profile.styleMap),
    profile.bareStyles.cite_paragraph,
    profile.bareStyles.card_body,
  ].filter((id): id is string => !!id);

  const warnings: MappingWarning[] = [];
  const seen = new Set<string>();

  for (const id of targets) {
    if (seen.has(id)) continue;
    seen.add(id);
    const style = byId.get(id);
    const name = style?.name ?? id;
    // headings resolve by outline level rather than by name, so they are safe
    if (/^Heading\d$/.test(id)) continue;
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
