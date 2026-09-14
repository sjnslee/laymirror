// the focused document's absolute path.
//
// not through `api.docInfo()`: a doc id exists only once cardmirror has saved
// the file itself, so docInfo is null for every word-authored .docx. the
// filename chip names the document instead, and `pmd-recent-files` turns that
// name into a path.
//
// a name is not a file: two folders can each hold a 1ac.docx. the path the user
// pointed at with locate competes with the history on recency, so pointing at
// the right file wins until a newer open says otherwise. a wrong answer here
// cannot write the wrong file — `saveExisting` refuses a path this window does
// not have open — but it can leave laymirror unable to write until located.

import { currentFilename, readRecents, type RecentEntry } from './cardmirror.js';
import type { DocInfo } from './plugin-api.js';

/** a path the user pointed at, and when. */
export interface Located {
  path: string;
  at: number;
}

export type Resolved =
  | { kind: 'ok'; path: string; at: number }
  | { kind: 'ambiguous'; paths: string[] }
  /** no document, one laymirror cannot touch, or one cardmirror never
   *  wrote a history entry for. */
  | { kind: 'none'; because: 'no-document' | 'not-a-docx' | 'unlisted' };

const DOCX = /\.docx$/i;

const isDocx = (entry: RecentEntry): boolean =>
  !!entry.handle && (entry.format === 'docx' || DOCX.test(entry.filename));

export function resolveDocPath(
  info: DocInfo | null,
  located: (filename: string) => Located | null = () => null,
): Resolved {
  const filename = currentFilename() ?? info?.docTitle ?? null;
  if (!filename) return { kind: 'none', because: 'no-document' };

  // newest sighting per path, from the history and from locate
  const seen = new Map<string, number>();
  for (const entry of readRecents()) {
    if (entry.filename !== filename || !isDocx(entry)) continue;
    const at = entry.lastOpenedAt ?? 0;
    seen.set(entry.handle!, Math.max(seen.get(entry.handle!) ?? at, at));
  }
  const pointed = DOCX.test(filename) ? located(filename) : null;
  if (pointed) seen.set(pointed.path, Math.max(seen.get(pointed.path) ?? 0, pointed.at));

  // a .docx missing from the history is not the same as a .cmir: only the
  // first can be rescued by asking the user where the file is
  if (seen.size === 0) {
    return { kind: 'none', because: DOCX.test(filename) ? 'unlisted' : 'not-a-docx' };
  }

  // two files share this name: the one seen last is the one on screen
  const [first, second] = [...seen].sort((a, b) => b[1] - a[1]);
  if (!second || first![1] > second[1]) return { kind: 'ok', path: first![0], at: first![1] };

  return { kind: 'ambiguous', paths: [...seen.keys()] };
}
