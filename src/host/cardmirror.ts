// every undocumented cardmirror internal lives here, and nowhere else.
// each constant is stamped with the version it was verified against, so a
// cardmirror upgrade breaks a canary test instead of breaking the plugin
// mid-round.
//
// verified by reading the shipped build (/Applications/cardmirror.app,
// app.asar) and by the phase 0 spike run against a live editor.

export const VERIFIED_AGAINST = '1.3.0';

/** the editor container. `#editor` observed in the spike; the pane class
 *  is the multi-pane form. */
export const EDITOR_SELECTOR = '#editor, .pmd-pane-editor';

/** stable classes the stylesheet targets. counts observed in the spike on
 *  an imported lay template: tag 1, card-body 12, card 1. */
export const CLASS = {
  tag: 'pmd-tag',
  cardBody: 'pmd-card-body',
  citePara: 'pmd-cite-para',
  undertag: 'pmd-undertag',
  analytic: 'pmd-analytic',
  card: 'pmd-card',
} as const;

/** css custom properties cardmirror writes on :root and #editor. we set the
 *  ones that exist and write direct rules for what has no variable (per-type
 *  font-family). observed live: normal 11pt, tag 13pt, pocket 26pt,
 *  body-font '"Times New Roman", ...'. */
export const CSS_VAR = {
  sizeNormal: '--pmd-size-normal',
  sizePocket: '--pmd-size-pocket',
  sizeHat: '--pmd-size-hat',
  sizeBlock: '--pmd-size-block',
  sizeTag: '--pmd-size-tag',
  sizeAnalytic: '--pmd-size-analytic',
  sizeCite: '--pmd-size-cite',
  sizeUnderline: '--pmd-size-underline',
  sizeEmphasis: '--pmd-size-emphasis',
  sizeUndertag: '--pmd-size-undertag',
  bodyFont: '--pmd-body-font',
} as const;

export const LS = { recents: 'pmd-recent-files' } as const;

/** our marker, stored beside cardmirror's own `cmirDocId`. */
export const MARKER_PROP = 'layMirrorProfile';
export const DOC_ID_PROP = 'cmirDocId';

/** a `pmd-recent-files` entry. `handle` is an absolute path on electron. */
export interface RecentEntry {
  handle: string | null;
  filename: string;
  format: 'cmir' | 'docx' | string;
}

export function readRecents(): RecentEntry[] {
  try {
    const raw = JSON.parse(localStorage.getItem(LS.recents) ?? '[]');
    return Array.isArray(raw) ? (raw as RecentEntry[]) : [];
  } catch {
    return [];
  }
}

/** the open document as a prosemirror node, via the editor dom's
 *  `pmViewDesc` back-reference (prosemirror-view sets `dom.pmViewDesc = this`).
 *  undocumented but stable; verified in the spike — returned a `doc` node with
 *  children `block`, `paragraph`, `card`. null when the shape is not what we
 *  expect, never throws. */
export function readDocNode(): unknown | null {
  try {
    const dom = document.querySelector('.ProseMirror') as
      | (Element & { pmViewDesc?: { node?: unknown } })
      | null;
    return dom?.pmViewDesc?.node ?? null;
  } catch {
    return null;
  }
}
