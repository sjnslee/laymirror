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

/** the class on each block type's element, from cardmirror's own schema —
 *  `pocket` renders as `h1.pmd-pocket`, `hat` as `h2.pmd-hat`, and so on.
 *  read off the shipped 1.3.0 schema's `toDOM`. */
export const CLASS = {
  pocket: 'pmd-pocket',
  hat: 'pmd-hat',
  block: 'pmd-block',
  tag: 'pmd-tag',
  cardBody: 'pmd-card-body',
  citePara: 'pmd-cite-para',
  undertag: 'pmd-undertag',
  analytic: 'pmd-analytic',
  card: 'pmd-card',
} as const;

export const LS = { recents: 'pmd-recent-files' } as const;

/** cardmirror paints the open document's filename here, and into
 *  `document.title`. these are the only signals that name the current
 *  document when it has no doc id — which is every word-authored .docx. */
export const DOC_NAME_CHIP = 'doc-name-chip-text';
export const TITLE_SUFFIX = ' — CardMirror';

/** our marker, stored beside cardmirror's own `cmirDocId`. */
export const MARKER_PROP = 'layMirrorTemplate';
export const DOC_ID_PROP = 'cmirDocId';

/** a `pmd-recent-files` entry. `handle` is an absolute path on electron.
 *
 *  the list is capped at ten and the open document is unshifted to the front
 *  with a fresh `lastOpenedAt`, so it is a history — not a list of what is
 *  open now. treating it as the latter is what made every document look
 *  ambiguous. */
export interface RecentEntry {
  handle: string | null;
  filename: string;
  format: 'cmir' | 'docx' | string;
  lastOpenedAt?: number;
}

/** the filename cardmirror is showing for the document in front of the user. */
export function currentFilename(): string | null {
  const chip = document.getElementById(DOC_NAME_CHIP)?.textContent?.trim();
  if (chip) return chip;

  const title = document.title.endsWith(TITLE_SUFFIX)
    ? document.title.slice(0, -TITLE_SUFFIX.length).trim()
    : '';
  return title || null;
}

export function readRecents(): RecentEntry[] {
  try {
    const raw = JSON.parse(localStorage.getItem(LS.recents) ?? '[]');
    return Array.isArray(raw) ? (raw as RecentEntry[]) : [];
  } catch {
    return [];
  }
}
