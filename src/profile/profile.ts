// one profile, rendered two ways: toCss for the editor, toOoxml for the file.
// they read the same object so the screen cannot lie about the file.

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

export type RunType =
  | 'underline_mark'
  | 'emphasis_mark'
  | 'cite_mark'
  | 'analytic_mark'
  | 'undertag_mark';

export interface TypeSpec {
  /** styleId written into document.xml. */
  styleId: string;
  /** w:name — what cardmirror's legacy remapper reads back on import.
   *  changing this silently breaks round-trip. */
  styleName: string;
  font?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: 'none' | 'single' | 'thick' | 'double';
  smallCaps?: boolean;
  /** hex6, no leading '#'. */
  color?: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  indentLeftDxa?: number;
  indentRightDxa?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacing?: { rule: 'auto' | 'exact' | 'atLeast'; value: number };
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  outlineLevel?: number | null;
}

export interface PageSetup {
  widthTwips: number;
  heightTwips: number;
  margin: {
    top: number;
    right: number;
    bottom: number;
    left: number;
    header: number;
    footer: number;
  };
}

export interface Profile {
  id: string;
  name: string;
  types: Record<BlockType | RunType, TypeSpec>;
  page: PageSetup;
  /** raw header1/footer1 xml from the donor, tokens unresolved. */
  headerXml: string | null;
  footerXml: string | null;
  /** basename only, e.g. "Lay Cut Cards.dotx" — word only basename-matches,
   *  and donors carry absolute paths through someone's home directory. */
  attachedTemplate: string | null;
  /** the donor's styles.xml, the base we merge into. */
  donorStylesXml: string;
  fontFallbacks: Record<string, string>;
}
