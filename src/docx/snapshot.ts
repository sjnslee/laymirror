// a document's identity: the parts that make it the school's document
// rather than a generic export.
//
// cardmirror's exporter builds a fresh package on every save — one hardcoded
// letter sectPr with 1" margins, its own styles.xml, and no header, footer or
// theme at all (src/export/exporter.ts:111). so these parts are not something
// we decorate a file with; they are something the file keeps losing, and this
// module is how it gets them back.
//
// everything here is carried verbatim. the previous design parsed the donor
// into a model and re-emitted it, which silently dropped every property nobody
// remembered to parse — smallCaps, thick underlines, borders. bytes cannot
// forget.

import { CONTENT_TYPES, readText, writeText, type Parts } from './zip.js';

const DOCUMENT = 'word/document.xml';
const DOC_RELS = 'word/_rels/document.xml.rels';
const SETTINGS_RELS = 'word/_rels/settings.xml.rels';

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HEADER_TYPE = `${REL_BASE}/header`;
const FOOTER_TYPE = `${REL_BASE}/footer`;
const TEMPLATE_TYPE = `${REL_BASE}/attachedTemplate`;

/** carried whole whenever the donor has them. */
const CARRIED = ['word/theme/theme1.xml', 'word/fontTable.xml', 'word/styles.xml'];

export interface Snapshot {
  /** part name -> xml text, exactly as the donor had it. */
  parts: Record<string, string>;
  /** the body's `<w:sectPr>`, with its original r:ids. */
  sectPr: string | null;
  /** header/footer part name -> the donor's relationship id for it. */
  relIds: Record<string, string>;
  /** `[Content_Types].xml` overrides for the carried parts. */
  overrides: Record<string, string>;
  /** basename only — word matches an attached template by basename, and
   *  donors carry an absolute path through somebody's home directory. */
  attachedTemplate: string | null;
}

const attr = (tag: string, name: string): string | null =>
  new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1] ?? null;

const relationships = (xml: string): string[] =>
  [...xml.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => m[0]);

/** resolve a document-relative rels target to a package part name. */
const partFor = (target: string): string =>
  target.startsWith('/') ? target.slice(1) : `word/${target}`.replace(/\/\.\//g, '/');

/** the test that drives the whole save pipeline: only word writes a header
 *  reference, because cardmirror's exporter never emits one. a file that has
 *  one was last written by word, and is therefore authoritative. */
export function hasOwnHeader(parts: Parts): boolean {
  const rels = readText(parts, DOC_RELS);
  if (!rels) return false;
  return relationships(rels).some((rel) => {
    const type = attr(rel, 'Type');
    return type === HEADER_TYPE || type === FOOTER_TYPE;
  });
}

export function readSectPr(documentXml: string): string | null {
  const all = [...documentXml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/g)];
  return all.length ? all[all.length - 1]![0] : null;
}

/** null when the package has nothing worth carrying — a document with no
 *  header, no theme and no styles of its own has no identity to keep, and a
 *  snapshot of nothing would later overwrite something. */
export function captureSnapshot(parts: Parts): Snapshot | null {
  const relsXml = readText(parts, DOC_RELS);
  const snapshotParts: Record<string, string> = {};
  const relIds: Record<string, string> = {};

  if (relsXml) {
    for (const rel of relationships(relsXml)) {
      const type = attr(rel, 'Type');
      if (type !== HEADER_TYPE && type !== FOOTER_TYPE) continue;
      const target = attr(rel, 'Target');
      const id = attr(rel, 'Id');
      if (!target || !id) continue;
      const name = partFor(target);
      const xml = readText(parts, name);
      if (xml === null) continue;
      snapshotParts[name] = xml;
      relIds[name] = id;
    }
  }

  for (const name of CARRIED) {
    const xml = readText(parts, name);
    if (xml !== null) snapshotParts[name] = xml;
  }

  const documentXml = readText(parts, DOCUMENT);
  const sectPr = documentXml ? readSectPr(documentXml) : null;

  if (Object.keys(snapshotParts).length === 0 && !sectPr) return null;

  const overrides: Record<string, string> = {};
  const ct = readText(parts, CONTENT_TYPES);
  if (ct) {
    for (const override of ct.matchAll(/<Override\b[^>]*\/>/g)) {
      const partName = attr(override[0], 'PartName');
      if (partName && snapshotParts[partName.replace(/^\//, '')]) {
        overrides[partName.replace(/^\//, '')] = override[0];
      }
    }
  }

  return {
    parts: snapshotParts,
    sectPr,
    relIds,
    overrides,
    attachedTemplate: readAttachedTemplate(parts),
  };
}

function readAttachedTemplate(parts: Parts): string | null {
  const rels = readText(parts, SETTINGS_RELS);
  if (!rels) return null;
  for (const rel of relationships(rels)) {
    if (attr(rel, 'Type') !== TEMPLATE_TYPE) continue;
    const target = attr(rel, 'Target');
    if (!target) continue;
    const basename = target.split(/[\\/]/).pop();
    return basename || null;
  }
  return null;
}

const WP = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const CONTENT_TYPE_BY_PART: [RegExp, string][] = [
  [/\/header\d*\.xml$/, `${WP}.header+xml`],
  [/\/footer\d*\.xml$/, `${WP}.footer+xml`],
  [/\/styles\.xml$/, `${WP}.styles+xml`],
  [/\/fontTable\.xml$/, `${WP}.fontTable+xml`],
  [/\/theme\d*\.xml$/, 'application/vnd.openxmlformats-officedocument.theme+xml'],
];

function defaultOverride(partName: string): string | null {
  for (const [pattern, contentType] of CONTENT_TYPE_BY_PART) {
    if (pattern.test(partName)) {
      return `<Override PartName="/${partName}" ContentType="${contentType}"/>`;
    }
  }
  return null;
}

/** ids we mint ourselves, so restoring twice reuses an entry rather than
 *  stacking a second one, and so we never collide with the ids cardmirror's
 *  own exporter hands out. */
const mintedId = (partName: string): string =>
  `rIdLayMirror${partName.replace(/[^a-zA-Z0-9]/g, '')}`;

function upsertRelationship(relsXml: string, id: string, type: string, target: string): string {
  const entry = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  if (relsXml.includes(`Id="${id}"`)) {
    return relsXml.replace(new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*/>`), entry);
  }
  return relsXml.replace('</Relationships>', `${entry}</Relationships>`);
}

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>';

/** put the document's identity back onto a package cardmirror has just
 *  rebuilt. the donor's own relationship ids are not reused — the fresh
 *  package hands out its own rIdN and they would collide — so each part gets
 *  a minted id and the sectPr's references are rewritten to match. */
export function restoreSnapshot(parts: Parts, snapshot: Snapshot): void {
  for (const [name, xml] of Object.entries(snapshot.parts)) {
    writeText(parts, name, xml);
  }

  let ct = readText(parts, CONTENT_TYPES);
  if (ct) {
    for (const name of Object.keys(snapshot.parts)) {
      if (ct.includes(`PartName="/${name}"`)) continue;
      // the donor's own override when we have it, otherwise one built from
      // the part's kind: a part with no declared content type makes word
      // call the whole file corrupt
      const override = snapshot.overrides[name] ?? defaultOverride(name);
      if (override) ct = ct.replace('</Types>', `${override}</Types>`);
    }
    writeText(parts, CONTENT_TYPES, ct);
  }

  let rels = readText(parts, DOC_RELS) ?? EMPTY_RELS;
  const remapped: Record<string, string> = {};
  for (const [name, donorId] of Object.entries(snapshot.relIds)) {
    const id = mintedId(name);
    const type = name.includes('/footer') ? FOOTER_TYPE : HEADER_TYPE;
    rels = upsertRelationship(rels, id, type, name.replace(/^word\//, ''));
    remapped[donorId] = id;
  }
  writeText(parts, DOC_RELS, rels);

  const documentXml = readText(parts, DOCUMENT);
  if (documentXml && snapshot.sectPr) {
    writeText(parts, DOCUMENT, applySectPr(documentXml, retargetSectPr(snapshot.sectPr, remapped)));
  }
}

/** rewrite `r:id` on the header/footer references to the ids just minted, and
 *  drop any reference whose part didn't come back — a dangling r:id makes
 *  word declare the file corrupt. */
export function retargetSectPr(sectPr: string, remapped: Record<string, string>): string {
  return sectPr.replace(
    /<w:(header|footer)Reference\b[^>]*\/>/g,
    (ref) => {
      const id = attr(ref, 'r:id');
      const next = id ? remapped[id] : undefined;
      return next ? ref.replace(/r:id="[^"]*"/, `r:id="${next}"`) : '';
    },
  );
}

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function applySectPr(documentXml: string, sectPr: string): string {
  // the restored sectPr references parts by r:id, so the prefix has to be
  // bound even though cardmirror's export never used it
  let xml = documentXml;
  if (!/<w:document\b[^>]*xmlns:r=/.test(xml)) {
    xml = xml.replace(/<w:document\b/, `<w:document xmlns:r="${R_NS}"`);
  }
  if (/<w:sectPr\b/.test(xml)) {
    return xml.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/, sectPr);
  }
  return xml.replace('</w:body>', `${sectPr}</w:body>`);
}
