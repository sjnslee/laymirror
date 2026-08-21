import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { zip, unzip, isDocx, readText } from '../src/docx/zip.js';
import { readMarker, writeMarker, clearMarker } from '../src/docx/marker.js';

const donor = new Uint8Array(
  readFileSync(new URL('./fixtures/donor.docx', import.meta.url)),
);

describe('donor', () => {
  it('carries no personal data', () => {
    const parts = unzip(donor);
    const all = Object.keys(parts)
      .map((n) => readText(parts, n) ?? '')
      .join('');
    for (const leak of ['a-name', 'A Name', 'Another Name', '/Users/']) {
      expect(all).not.toContain(leak);
    }
  });

  it('points attachedTemplate at a basename, never an absolute path', () => {
    const rels = readText(unzip(donor), 'word/_rels/settings.xml.rels') ?? '';
    expect(rels).toContain('Target="Lay%20Cut%20Cards.dotx"');
    expect(rels).not.toContain('file:');
  });

  it('is a docx and starts unmarked', () => {
    const parts = unzip(donor);
    expect(isDocx(parts)).toBe(true);
    expect(readMarker(parts)).toBeNull();
  });

  it('takes the marker without disturbing any other part', () => {
    const before = unzip(donor);
    const parts = unzip(donor);
    writeMarker(parts, 'sample-lay');
    const after = unzip(zip(parts));

    expect(readMarker(after)).toBe('sample-lay');
    expect(Object.keys(after)).toEqual(Object.keys(before));
    const changed = Object.keys(before).filter(
      (n) => Buffer.compare(Buffer.from(after[n]!), Buffer.from(before[n]!)) !== 0,
    );
    expect(changed).toEqual(['docProps/custom.xml']);
  });

  it('clears back to the original custom.xml exactly', () => {
    const parts = unzip(donor);
    const original = readText(parts, 'docProps/custom.xml');
    writeMarker(parts, 'sample-lay');
    clearMarker(parts);
    expect(readText(parts, 'docProps/custom.xml')).toBe(original);
  });
});
