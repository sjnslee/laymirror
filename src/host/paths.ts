// resolving the focused document's absolute path.
//
// api.docInfo() gives a title and (sometimes) a doc id; `pmd-recent-files`
// gives absolute paths. the plan disambiguated same-named documents by
// comparing each candidate's `cmirDocId` against docInfo().docId — but the
// phase 0 spike found that id absent until cardmirror itself saves the file,
// which is exactly never for a word-authored lay docx. so the id is used when
// present and ambiguity is reported honestly when it isn't, rather than
// guessing at a file we are about to rewrite.

import { readRecents } from './cardmirror.js';
import { readFile } from './electron.js';
import type { DocInfo } from './plugin-api.js';
import { unzip } from '../docx/zip.js';
import { readDocId } from '../docx/marker.js';

export type Resolved =
  | { kind: 'ok'; path: string }
  | { kind: 'ambiguous'; paths: string[] }
  | { kind: 'none' };

export async function resolveDocPath(info: DocInfo | null): Promise<Resolved> {
  const docx = readRecents().filter(
    (r): r is typeof r & { handle: string } => !!r.handle && r.format === 'docx',
  );
  if (docx.length === 0) return { kind: 'none' };

  const named = info?.docTitle
    ? docx.filter((r) => r.filename === info.docTitle)
    : [];
  const candidates = named.length > 0 ? named : docx;

  if (candidates.length === 1) return { kind: 'ok', path: candidates[0]!.handle };

  // more than one file could be the open document. an id settles it when the
  // document has one.
  if (info?.docId) {
    const matches: string[] = [];
    for (const c of candidates) {
      const file = await readFile(c.handle);
      if (!file) continue;
      try {
        if (readDocId(unzip(file.bytes)) === info.docId) matches.push(c.handle);
      } catch {
        // unreadable or mid-write; it just isn't a match
      }
    }
    if (matches.length === 1) return { kind: 'ok', path: matches[0]! };
  }

  return { kind: 'ambiguous', paths: candidates.map((c) => c.handle) };
}
