// a template's identity: the parts that make a document the school's rather
// than a generic export.
//
// cardmirror's exporter builds a fresh package on every save — one hardcoded
// letter sectPr with 1" margins, its own styles.xml, and no header, footer or
// theme at all. so these parts are not something we decorate a file with; they
// are something the file keeps losing, and this module is how it gets them
// back.
//
// everything here is carried verbatim, bytes and all. an earlier design parsed
// the template into a model and re-emitted it, which silently dropped every
// property nobody remembered to parse — smallCaps, thick underlines, borders.
// bytes cannot forget, and they carry a school's logo as readily as its fonts.

import { CONTENT_TYPES, readText, writeText, type Parts } from './zip.js';

const DOCUMENT = 'word/document.xml';
const DOC_RELS = 'word/_rels/document.xml.rels';
const SETTINGS_RELS = 'word/_rels/settings.xml.rels';

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HEADER_TYPE = `${REL_BASE}/header`;
const FOOTER_TYPE = `${REL_BASE}/footer`;
const TEMPLATE_TYPE = `${REL_BASE}/attachedTemplate`;

/** word resolves these parts through a relationship, not by their name, so a
 *  part copied in without one is a part word never reads.
 *
 *  cardmirror's exporter writes relationships for styles and settings, and for
 *  numbering only when it emitted a list of its own — it never writes one for
 *  the theme or the font table. so a template's theme fonts (`asciiTheme=
 *  "minorHAnsi"`) would resolve to nothing, which is most of a school's
 *  typography, silently. */
const RELATED: [RegExp, string][] = [
  [/\/styles\.xml$/, `${REL_BASE}/styles`],
  [/\/numbering\.xml$/, `${REL_BASE}/numbering`],
  [/\/fontTable\.xml$/, `${REL_BASE}/fontTable`],
  [/\/theme\/theme\d*\.xml$/, `${REL_BASE}/theme`],
];

/** carried whole whenever the template has them. `numbering.xml` is here
 *  because a lay template numbers its block headings, and a list that
 *  references a missing `numId` renders as nothing at all. */
const CARRIED = [
  'word/styles.xml',
  'word/theme/theme1.xml',
  'word/fontTable.xml',
  'word/numbering.xml',
];

export interface Snapshot {
  /** part name -> bytes, exactly as the template had them. */
  parts: Parts;
  /** the body's `<w:sectPr>`, with its original r:ids. */
  sectPr: string | null;
  /** header/footer part name -> the template's relationship id for it. */
  relIds: Record<string, string>;
  /** `[Content_Types].xml` declarations the carried parts need. */
  overrides: Record<string, string>;
  defaults: Record<string, string>;
  /** basename only — word matches an attached template by basename, and
   *  templates carry an absolute path through somebody's home directory. */
  attachedTemplate: string | null;
}

const attr = (tag: string, name: string): string | null =>
  new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1] ?? null;

const relationships = (xml: string): string[] =>
  [...xml.matchAll(/<Relationship\b[^>]*\/>/g)].map((m) => m[0]);

/** resolve a rels target, relative to the folder the OWNING PART lives in —
 *  not the `_rels` folder its relationships file sits in. */
function partFor(owner: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const dir = owner.slice(0, owner.lastIndexOf('/'));
  const segments = `${dir}/${target}`.split('/');
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === '.' || segment === '') continue;
    if (segment === '..') out.pop();
    else out.push(segment);
  }
  return out.join('/');
}

const relsFor = (partName: string): string =>
  `${partName.slice(0, partName.lastIndexOf('/'))}/_rels/${partName.slice(partName.lastIndexOf('/') + 1)}.rels`;

export function readSectPr(documentXml: string): string | null {
  const all = [...documentXml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/g)];
  return all.length ? all[all.length - 1]![0] : null;
}

/** a header or footer brings its own rels with it — a school crest in the
 *  header is an image part the header points at, and copying the header
 *  without it puts a red x on every page. external targets are left alone:
 *  they are urls, not parts. */
function carryDependencies(parts: Parts, partName: string, into: Parts): void {
  const relsName = relsFor(partName);
  const relsXml = readText(parts, relsName);
  if (!relsXml) return;
  into[relsName] = parts[relsName]!;

  for (const rel of relationships(relsXml)) {
    if (attr(rel, 'TargetMode') === 'External') continue;
    const target = attr(rel, 'Target');
    if (!target) continue;
    const name = partFor(partName, target);
    if (parts[name]) into[name] = parts[name]!;
  }
}

/** null when the package has nothing worth carrying — a document with no
 *  header, no theme and no styles of its own has no identity to keep, and a
 *  snapshot of nothing would later overwrite something. */
export function captureSnapshot(parts: Parts): Snapshot | null {
  const relsXml = readText(parts, DOC_RELS);
  const carried: Parts = {};
  const relIds: Record<string, string> = {};

  if (relsXml) {
    for (const rel of relationships(relsXml)) {
      const type = attr(rel, 'Type');
      if (type !== HEADER_TYPE && type !== FOOTER_TYPE) continue;
      const target = attr(rel, 'Target');
      const id = attr(rel, 'Id');
      if (!target || !id) continue;
      const name = partFor(DOCUMENT, target);
      if (!parts[name]) continue;
      carried[name] = parts[name]!;
      relIds[name] = id;
      carryDependencies(parts, name, carried);
    }
  }

  for (const name of CARRIED) {
    if (parts[name]) carried[name] = parts[name]!;
  }

  const documentXml = readText(parts, DOCUMENT);
  const sectPr = documentXml ? readSectPr(documentXml) : null;

  if (Object.keys(carried).length === 0 && !sectPr) return null;

  return {
    parts: carried,
    sectPr,
    relIds,
    ...contentTypes(parts, Object.keys(carried)),
    attachedTemplate: readAttachedTemplate(parts),
  };
}

/** the declarations `[Content_Types].xml` needs for the carried parts: the
 *  template's own where it has one, so a part word declared unusually keeps
 *  its declaration, and a Default for every extension the media uses. */
function contentTypes(
  parts: Parts,
  names: readonly string[],
): { overrides: Record<string, string>; defaults: Record<string, string> } {
  const overrides: Record<string, string> = {};
  const defaults: Record<string, string> = {};
  const ct = readText(parts, CONTENT_TYPES);
  if (!ct) return { overrides, defaults };

  const wanted = new Set(names);
  for (const match of ct.matchAll(/<Override\b[^>]*\/>/g)) {
    const partName = attr(match[0], 'PartName')?.replace(/^\//, '');
    if (partName && wanted.has(partName)) overrides[partName] = match[0];
  }

  const extensions = new Set(
    names
      .filter((name) => !overrides[name])
      .map((name) => name.slice(name.lastIndexOf('.') + 1).toLowerCase()),
  );
  for (const match of ct.matchAll(/<Default\b[^>]*\/>/g)) {
    const extension = attr(match[0], 'Extension')?.toLowerCase();
    if (extension && extensions.has(extension)) defaults[extension] = match[0];
  }

  return { overrides, defaults };
}

function readAttachedTemplate(parts: Parts): string | null {
  const rels = readText(parts, SETTINGS_RELS);
  if (!rels) return null;
  for (const rel of relationships(rels)) {
    if (attr(rel, 'Type') !== TEMPLATE_TYPE) continue;
    const target = attr(rel, 'Target');
    if (!target) continue;
    return target.split(/[\\/]/).pop() || null;
  }
  return null;
}

const WP = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const CONTENT_TYPE_BY_PART: [RegExp, string][] = [
  [/\/header\d*\.xml$/, `${WP}.header+xml`],
  [/\/footer\d*\.xml$/, `${WP}.footer+xml`],
  [/\/styles\.xml$/, `${WP}.styles+xml`],
  [/\/numbering\.xml$/, `${WP}.numbering+xml`],
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

/** put the school's document back onto a package cardmirror has just rebuilt.
 *
 *  `override` lets the caller substitute a part it has already edited — the
 *  headers, with the user's own team code filled into them — without the
 *  snapshot itself ever being written to.
 *
 *  the template's own relationship ids are not reused: the fresh package hands
 *  out its own rIdN and they would collide, so each part gets a minted id and
 *  the sectPr's references are rewritten to match. */
export function restoreSnapshot(
  parts: Parts,
  snapshot: Snapshot,
  override: Record<string, Uint8Array> = {},
): void {
  for (const [name, bytes] of Object.entries(snapshot.parts)) {
    parts[name] = override[name] ?? bytes;
  }

  let ct = readText(parts, CONTENT_TYPES);
  if (ct) {
    for (const [extension, declaration] of Object.entries(snapshot.defaults)) {
      if (new RegExp(`<Default\\b[^>]*Extension="${extension}"`, 'i').test(ct)) continue;
      ct = ct.replace(/<Types\b[^>]*>/, (open) => `${open}${declaration}`);
    }
    for (const name of Object.keys(snapshot.parts)) {
      if (ct.includes(`PartName="/${name}"`)) continue;
      // the template's own override when we have it, otherwise one built from
      // the part's kind: a part with no declared content type makes word call
      // the whole file corrupt
      const override = snapshot.overrides[name] ?? defaultOverride(name);
      if (override) ct = ct.replace('</Types>', `${override}</Types>`);
    }
    writeText(parts, CONTENT_TYPES, ct);
  }

  let rels = readText(parts, DOC_RELS) ?? EMPTY_RELS;
  const remapped: Record<string, string> = {};
  for (const [name, templateId] of Object.entries(snapshot.relIds)) {
    const id = mintedId(name);
    const type = name.includes('/footer') ? FOOTER_TYPE : HEADER_TYPE;
    rels = upsertRelationship(rels, id, type, name.replace(/^word\//, ''));
    remapped[templateId] = id;
  }
  for (const name of Object.keys(snapshot.parts)) {
    const type = RELATED.find(([pattern]) => pattern.test(name))?.[1];
    // one relationship per type: cardmirror already points at styles.xml, and
    // a second styles relationship is a package word rejects
    if (!type || rels.includes(`Type="${type}"`)) continue;
    rels = upsertRelationship(rels, mintedId(name), type, name.replace(/^word\//, ''));
  }
  writeText(parts, DOC_RELS, rels);

  const documentXml = readText(parts, DOCUMENT);
  if (documentXml && snapshot.sectPr) {
    writeText(parts, DOCUMENT, applySectPr(documentXml, retargetSectPr(snapshot.sectPr, remapped)));
  }
}

/** rewrite `r:id` on the header/footer references to the ids just minted, and
 *  drop any reference whose part didn't come back — a dangling r:id makes word
 *  declare the file corrupt. */
export function retargetSectPr(sectPr: string, remapped: Record<string, string>): string {
  return sectPr.replace(/<w:(header|footer)Reference\b[^>]*\/>/g, (ref) => {
    const id = attr(ref, 'r:id');
    const next = id ? remapped[id] : undefined;
    return next ? ref.replace(/r:id="[^"]*"/, `r:id="${next}"`) : '';
  });
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
