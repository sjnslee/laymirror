// resolving the focused document's absolute path.
//
// not through `api.docInfo()`: cardmirror builds it as
// `docId ? {docId, docTitle} : null`, and the doc id only exists once cardmirror
// has saved the file itself — so it is null for every word-authored .docx, which
// is exactly the case laymirror is for.
//
// the filename chip names the document instead, and `pmd-recent-files` turns
// that name into a path. it is a history, so it is only ever consulted for a
// name we already have.

import { currentFilename, readRecents, type RecentEntry } from './cardmirror.js';
import type { DocInfo } from './plugin-api.js';

export type Resolved =
  | { kind: 'ok'; path: string }
  | { kind: 'ambiguous'; paths: string[] }
  /** no document, or one laymirror cannot touch. */
  | { kind: 'none'; because: 'no-document' | 'not-a-docx' };

type Openable = RecentEntry & { handle: string };

const isDocx = (entry: RecentEntry): entry is Openable =>
  !!entry.handle &&
  (entry.format === 'docx' || entry.filename.toLowerCase().endsWith('.docx'));

const openedAt = (entry: RecentEntry): number => entry.lastOpenedAt ?? 0;

export function resolveDocPath(info: DocInfo | null): Resolved {
  const filename = currentFilename() ?? info?.docTitle ?? null;
  if (!filename) return { kind: 'none', because: 'no-document' };

  const named = readRecents().filter(
    (entry) => entry.filename === filename && isDocx(entry),
  ) as Openable[];

  // a name with no docx behind it is a .cmir, or a document never saved.
  // guessing at another file would rewrite the wrong one.
  if (named.length === 0) return { kind: 'none', because: 'not-a-docx' };
  if (named.length === 1) return { kind: 'ok', path: named[0]!.handle };

  // two files really do share this name: the one opened last is the one in
  // front of the user, and only a tie is undecidable
  const [first, second] = [...named].sort((a, b) => openedAt(b) - openedAt(a));
  if (openedAt(first!) > openedAt(second!)) return { kind: 'ok', path: first!.handle };

  return { kind: 'ambiguous', paths: named.map((entry) => entry.handle) };
}
