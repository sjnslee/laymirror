// every undocumented cardmirror internal lives here and nowhere else, so a
// cardmirror upgrade breaks a canary test rather than the plugin mid-round.
// read off the shipped build (/Applications/cardmirror.app, app.asar) at 1.3.0.

export const LS = { recents: 'pmd-recent-files' } as const;

/** where the sanctioned api keeps a plugin's storage bag: one localStorage entry
 *  per plugin, holding plain json.
 *
 *  read directly because cardmirror only hands the api object to a command's
 *  `run()`, and the save watcher has to start before any command has run. */
export const storageKey = (pluginId: string): string => `plugin:${pluginId}`;

/** cardmirror paints the open document's filename into both. they are the only
 *  signals naming a document with no doc id — which is every word-authored
 *  .docx. */
export const DOC_NAME_CHIP = 'doc-name-chip-text';
export const TITLE_SUFFIX = ' — CardMirror';

/** our marker, stored beside cardmirror's own `cmirDocId`. */
export const MARKER_PROP = 'layMirrorTemplate';

/** a `pmd-recent-files` entry. `handle` is an absolute path on electron.
 *
 *  a history, not a list of what is open now: it is capped at ten and the open
 *  document is unshifted to the front with a fresh `lastOpenedAt`. */
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
