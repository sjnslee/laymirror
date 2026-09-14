// a template, and everything laymirror reads out of one.
//
// what is kept is the template cut down to the parts a blueprint is built from,
// each byte for byte, and everything else is derived from those on demand. the
// rest of a template — macros, sample cards, a thumbnail — never reaches the
// document, and the kept copy sits in localStorage cardmirror shares.

import { findFields, type Field } from '../docx/fields.js';
import { captureSnapshot, type Snapshot } from '../docx/snapshot.js';
import { CONTENT_TYPES, isDocx, readText, unzip, writeText, zip, type Parts } from '../docx/zip.js';
import { deriveBareStyles, deriveStyleMap, readStyles, type BareStyles } from './styles.js';

const STYLES = 'word/styles.xml';
const DOCUMENT = 'word/document.xml';

/** read alongside the snapshot's own parts: which parts are headers, and what
 *  template the file was attached to. */
const READ_WITH = [
  CONTENT_TYPES,
  'word/_rels/document.xml.rels',
  'word/settings.xml',
  'word/_rels/settings.xml.rels',
];

export interface Template {
  /** one per file, so two templates cannot collide. */
  id: string;
  name: string;
  /** where the user picked it from, so an apply can go back and re-read it */
  path: string | null;
  docx: Uint8Array;
}

export interface Blueprint {
  snapshot: Snapshot;
  /** cardmirror's exported style id -> the id this template defines. an id
   *  absent from the map is left as cardmirror wrote it. */
  styleMap: Record<string, string>;
  /** cite paragraphs and card bodies leave cardmirror with no `w:pStyle` at
   *  all, so they cannot be remapped by id — they are recognised from the
   *  marks their runs carry and given these ids instead. */
  bareStyles: BareStyles;
  /** the header and footer text the user may replace. */
  fields: Field[];
}

export type ReadResult =
  | { ok: true; blueprint: Blueprint; /** what to keep: reads back to the same blueprint */ kept: Uint8Array }
  | { ok: false; error: string };

/** the template with nothing a blueprint does not read. the body keeps only its
 *  page setup. */
function cut(parts: Parts, snapshot: Snapshot): Uint8Array {
  const kept: Parts = { ...snapshot.parts };
  for (const name of READ_WITH) {
    if (parts[name]) kept[name] = parts[name]!;
  }
  const open = /<w:document\b[^>]*>/.exec(readText(parts, DOCUMENT) ?? '')?.[0] ?? '<w:document>';
  writeText(kept, DOCUMENT, `${open}<w:body>${snapshot.sectPr ?? ''}</w:body></w:document>`);
  return zip(kept);
}

/** the header and footer parts as text, which is what the field code works in */
export function headerParts(snapshot: Snapshot): Record<string, string> {
  const out: Record<string, string> = {};
  for (const name of Object.keys(snapshot.parts)) {
    if (!/\/(header|footer)\d*\.xml$/.test(name)) continue;
    const xml = readText(snapshot.parts as Parts, name);
    if (xml !== null) out[name] = xml;
  }
  return out;
}

export function read(bytes: Uint8Array, name: string): ReadResult {
  let parts: Parts;
  try {
    parts = unzip(bytes);
  } catch {
    return { ok: false, error: `could not read ${name} — is it a word document or template?` };
  }
  if (!isDocx(parts)) return { ok: false, error: `${name} is not a word document or template` };

  const snapshot = captureSnapshot(parts);
  if (!snapshot) return { ok: false, error: `${name} has no styles, header or page setup to copy` };

  const styles = readStyles(readText(parts, STYLES) ?? '');

  return {
    ok: true,
    kept: cut(parts, snapshot),
    blueprint: {
      snapshot,
      styleMap: deriveStyleMap(styles),
      bareStyles: deriveBareStyles(styles),
      fields: findFields(headerParts(snapshot)),
    },
  };
}
