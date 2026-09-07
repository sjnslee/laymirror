// the focused document's absolute path.
//
// not through `api.docInfo()`: a doc id exists only once cardmirror has saved
// the file itself, so docInfo is null for every word-authored .docx. the
// filename chip names the document instead, and `pmd-recent-files` turns that
// name into a path.

import { currentFilename, readRecents, type RecentEntry } from './cardmirror.js';
import type { DocInfo } from './plugin-api.js';

export type Resolved =
  | { kind: 'ok'; path: string }
  | { kind: 'ambiguous'; paths: string[] }
  /** no document, one laymirror cannot touch, or one cardmirror never
   *  wrote a history entry for. */
  | { kind: 'none'; because: 'no-document' | 'not-a-docx' | 'unlisted' };

type Openable = RecentEntry & { handle: string };

const DOCX = /\.docx$/i;

const isDocx = (entry: RecentEntry): entry is Openable =>
  !!entry.handle && (entry.format === 'docx' || DOCX.test(entry.filename));

const openedAt = (entry: RecentEntry): number => entry.lastOpenedAt ?? 0;

export function resolveDocPath(info: DocInfo | null): Resolved {
  const filename = currentFilename() ?? info?.docTitle ?? null;
  if (!filename) return { kind: 'none', because: 'no-document' };

  const named = readRecents().filter(
    (entry) => entry.filename === filename && isDocx(entry),
  ) as Openable[];

  // a .docx missing from the history is not the same as a .cmir: only the
  // first can be rescued by asking the user where the file is
  if (named.length === 0) {
    return { kind: 'none', because: DOCX.test(filename) ? 'unlisted' : 'not-a-docx' };
  }
  if (named.length === 1) return { kind: 'ok', path: named[0]!.handle };

  // two files share this name: the one opened last is the one on screen
  const [first, second] = [...named].sort((a, b) => openedAt(b) - openedAt(a));
  if (openedAt(first!) > openedAt(second!)) return { kind: 'ok', path: first!.handle };

  return { kind: 'ambiguous', paths: named.map((entry) => entry.handle) };
}
