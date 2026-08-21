// resolving the focused document's absolute path.
//
// `api.docInfo()` looked like the way in, but it is not: cardmirror builds it
// as `docId ? {docId, docTitle} : null`, and the doc id only exists once
// cardmirror has saved the file itself. a word-authored lay .docx has none,
// so docInfo() returns null in exactly the case laymirror is for.
//
// what does name the open document is the filename cardmirror paints into its
// own chip and into the window title. `pmd-recent-files` then turns that name
// into a path — it is a history rather than a list of open files, so it is
// only ever consulted for a name we already have.

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
  // guessing at another file here would rewrite the wrong one.
  if (named.length === 0) return { kind: 'none', because: 'not-a-docx' };
  if (named.length === 1) return { kind: 'ok', path: named[0]!.handle };

  // two files really do share this name. the one opened most recently is the
  // one in front of the user; only a tie is undecidable.
  const [first, second] = [...named].sort((a, b) => openedAt(b) - openedAt(a));
  if (openedAt(first!) > openedAt(second!)) return { kind: 'ok', path: first!.handle };

  return { kind: 'ambiguous', paths: named.map((entry) => entry.handle) };
}
