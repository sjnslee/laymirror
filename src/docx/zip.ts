import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

export type Parts = Record<string, Uint8Array>;

export const CONTENT_TYPES = '[Content_Types].xml';

/** word rejects a package whose first entry isn't `[Content_Types].xml`, and
 *  a part added to the map lands last, so ordering is enforced on the way
 *  out rather than trusted. */
export function zip(parts: Parts): Uint8Array {
  const ordered: Parts = {};
  if (parts[CONTENT_TYPES]) ordered[CONTENT_TYPES] = parts[CONTENT_TYPES];
  for (const [name, bytes] of Object.entries(parts)) {
    if (name !== CONTENT_TYPES) ordered[name] = bytes;
  }
  return zipSync(ordered, { level: 6 });
}

export function unzip(bytes: Uint8Array): Parts {
  return unzipSync(bytes);
}

/** a partial read mid-save must abort rather than corrupt, so callers check
 *  this before rewriting anything. */
export function isDocx(parts: Parts): boolean {
  return CONTENT_TYPES in parts && 'word/document.xml' in parts;
}

export function readText(parts: Parts, name: string): string | null {
  const part = parts[name];
  return part ? strFromU8(part) : null;
}

export function writeText(parts: Parts, name: string, xml: string): void {
  parts[name] = strToU8(xml);
}
