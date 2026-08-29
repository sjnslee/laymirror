// a school template, and everything laymirror reads out of one.
//
// the template is kept as the file itself — the whole .docx, byte for byte —
// and everything else is derived from it on demand. an earlier design stored a
// digested profile instead, which meant every new thing laymirror learned to
// read (numbering, a crest in the header, a page break in a style) needed the
// user to load their template again. the file is the truth; a blueprint is
// just this session's reading of it.

import { findFields, type Field } from '../docx/fields.js';
import { captureSnapshot, type Snapshot } from '../docx/snapshot.js';
import { isDocx, readText, unzip, type Parts } from '../docx/zip.js';
import {
  deriveBareStyles,
  deriveStyleMap,
  readStyles,
  type BareStyles,
  type StyleInfo,
} from './styles.js';

const STYLES = 'word/styles.xml';

export interface Template {
  /** one per file, so two schools' templates cannot collide. */
  id: string;
  name: string;
  /** where the user picked it from, so an apply can go back and re-read it.
   *  null when it was loaded before laymirror kept paths. */
  path: string | null;
  docx: Uint8Array;
}

export interface Blueprint {
  snapshot: Snapshot;
  styles: StyleInfo[];
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

export type ReadResult = { ok: true; blueprint: Blueprint } | { ok: false; error: string };

/** the header and footer parts as text, which is the form the field machinery
 *  works in. */
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
  const styleMap = deriveStyleMap(styles);

  return {
    ok: true,
    blueprint: {
      snapshot,
      styles,
      styleMap,
      bareStyles: deriveBareStyles(styles),
      fields: findFields(headerParts(snapshot)),
    },
  };
}
