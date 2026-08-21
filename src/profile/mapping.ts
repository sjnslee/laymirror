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

import type { BlockType, Profile, RunType } from './profile.js';

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
  type: BlockType | RunType;
  styleName: string;
  message: string;
}

/** block types reimport by lowercased name; a name outside the legacy table
 *  exports fine and comes back as a plain paragraph. */
const BLOCK_TYPES: readonly BlockType[] = [
  'pocket',
  'hat',
  'block',
  'tag',
  'cite_paragraph',
  'card_body',
];

export function validateMapping(profile: Profile): MappingWarning[] {
  const warnings: MappingWarning[] = [];

  const ids = Object.values(profile.types).map((t) => t.styleId);
  const names = Object.values(profile.types).map((t) => t.styleName);
  const native = takesNativePath(ids, names);

  for (const type of BLOCK_TYPES) {
    const spec = profile.types[type];
    const byName = LEGACY_BY_NAME[spec.styleName.toLowerCase()];
    const byId = LEGACY_BY_ID[spec.styleId];
    // headings resolve by outline level instead of by name
    const isHeading = spec.outlineLevel !== undefined && spec.outlineLevel !== null;
    if (!byName && !byId && !isHeading) {
      warnings.push({
        type,
        styleName: spec.styleName,
        message: `"${spec.styleName}" is outside cardmirror's vocabulary — it will reimport as a plain paragraph`,
      });
    }
  }

  if (!native) {
    for (const type of ['cite_mark', 'underline_mark'] as const) {
      const spec = profile.types[type];
      if (!LEGACY_BY_NAME[spec.styleName.toLowerCase()] && !LEGACY_BY_ID[spec.styleId]) {
        warnings.push({
          type,
          styleName: spec.styleName,
          message:
            `"${spec.styleName}" only reimports on cardmirror's native path; ` +
            `emit ${REQUIRED_FOR_NATIVE_PATH.map((g) => g[0]).join(', ')} to keep it`,
        });
      }
    }
  }

  return warnings;
}
