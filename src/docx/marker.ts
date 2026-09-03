// the lay marker: a custom document property beside cardmirror's own `cmirDocId`.
// it travels with the file and survives a word round-trip, which is what makes
// activation per-file rather than per-machine.
//
// custom.xml is merged, never replaced — cardmirror keeps its doc id there, and
// sharepoint-derived templates a ContentTypeId.

import { MARKER_PROP } from '../host/cardmirror.js';
import { CONTENT_TYPES, readText, writeText, type Parts } from './zip.js';

const CUSTOM = 'docProps/custom.xml';
const RELS = '_rels/.rels';
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';

const CT_CUSTOM =
  '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>';
const REL_CUSTOM =
  '<Relationship Id="rIdLayMirrorCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>';

const EMPTY_CUSTOM =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"' +
  ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>';

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function unescapeXml(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&');
}

function propPattern(name: string): RegExp {
  return new RegExp(`<property\\b[^>]*\\bname="${name}"[^>]*>([\\s\\S]*?)</property>`);
}

function readProp(xml: string, name: string): string | null {
  const block = propPattern(name).exec(xml);
  if (!block?.[1]) return null;
  const value = /<vt:lpwstr>([\s\S]*?)<\/vt:lpwstr>/.exec(block[1]);
  return value?.[1] !== undefined ? unescapeXml(value[1]) : null;
}

/** pids must be unique within the part and start at 2. */
function nextPid(xml: string): number {
  const pids = [...xml.matchAll(/\bpid="(\d+)"/g)].map((m) => Number(m[1]));
  return (pids.length ? Math.max(...pids) : 1) + 1;
}

function upsertProp(xml: string, name: string, value: string): string {
  const existing = propPattern(name);
  const body = `<vt:lpwstr>${escapeXml(value)}</vt:lpwstr>`;
  if (existing.test(xml)) {
    return xml.replace(existing, (m) =>
      m.replace(/<vt:lpwstr>[\s\S]*?<\/vt:lpwstr>/, body),
    );
  }
  const prop = `<property fmtid="${FMTID}" pid="${nextPid(xml)}" name="${name}">${body}</property>`;
  return xml.replace('</Properties>', `${prop}</Properties>`);
}

function removeProp(xml: string, name: string): string {
  return xml.replace(propPattern(name), '');
}

/** the template id this document is marked with, or null when it isn't lay */
export function readMarker(parts: Parts): string | null {
  const xml = readText(parts, CUSTOM);
  return xml ? readProp(xml, MARKER_PROP) : null;
}

export function writeMarker(parts: Parts, templateId: string): void {
  const existing = readText(parts, CUSTOM);
  writeText(parts, CUSTOM, upsertProp(existing ?? EMPTY_CUSTOM, MARKER_PROP, templateId));
  if (!existing) ensureCustomWired(parts);
}

export function clearMarker(parts: Parts): void {
  const xml = readText(parts, CUSTOM);
  if (xml) writeText(parts, CUSTOM, removeProp(xml, MARKER_PROP));
}

/** a package that never had custom properties needs the part declared and
 *  related from the root, or word treats the file as corrupt. */
function ensureCustomWired(parts: Parts): void {
  const ct = readText(parts, CONTENT_TYPES);
  if (ct && !ct.includes('/docProps/custom.xml')) {
    writeText(parts, CONTENT_TYPES, ct.replace('</Types>', `${CT_CUSTOM}</Types>`));
  }
  const rels = readText(parts, RELS);
  if (rels && !rels.includes('docProps/custom.xml')) {
    writeText(parts, RELS, rels.replace('</Relationships>', `${REL_CUSTOM}</Relationships>`));
  }
}
