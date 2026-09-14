// every undocumented cardmirror internal lives here and nowhere else, so an
// upgrade has one file to check. read off the shipped app.asar at 1.3.0.

export const LS = { recents: 'pmd-recent-files', plugins: 'pmd-plugins' } as const;

/** the sanctioned api's storage bag: one localStorage entry per plugin, holding
 *  plain json. read directly because the watcher starts before any command. */
export const storageKey = (pluginId: string): string => `plugin:${pluginId}`;

/** the filename is painted into both, and they are the only signals naming a
 *  document with no doc id — which is every word-authored .docx. */
export const DOC_NAME_CHIP = 'doc-name-chip-text';
export const TITLE_SUFFIX = ' — CardMirror';

/** laymirror's marker, stored beside cardmirror's own `cmirDocId`. */
export const MARKER_PROP = 'layMirrorTemplate';

/** a `pmd-recent-files` entry. `handle` is an absolute path on electron.
 *  a history, not a list of what is open: capped at ten, newest first. */
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

/** whether cardmirror still has laymirror switched on.
 *
 *  api v1 has no unload hook, and a bundle that has run cannot be unloaded, so
 *  a plugin switched off mid session keeps whatever timers it started. reading
 *  the enabled flag is what lets laymirror stop writing files on its own. an
 *  unreadable or absent blob reads as enabled: the flag is written when the user
 *  toggles, and refusing to work because it is missing would be worse. */
export function isEnabled(pluginId: string): boolean {
  try {
    const raw: unknown = JSON.parse(localStorage.getItem(LS.plugins) ?? 'null');
    const flags = (raw as { enabled?: Record<string, unknown> } | null)?.enabled;
    if (!flags || typeof flags !== 'object') return true;
    return flags[pluginId] !== false;
  } catch {
    return true;
  }
}

export function readRecents(): RecentEntry[] {
  try {
    const raw = JSON.parse(localStorage.getItem(LS.recents) ?? '[]');
    return Array.isArray(raw) ? (raw as RecentEntry[]) : [];
  } catch {
    return [];
  }
}
