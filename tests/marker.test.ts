import { describe, it, expect } from 'vitest';
import { makeDocx } from './fixture.js';
import { zip, unzip, isDocx, readText, CONTENT_TYPES } from '../src/docx/zip.js';
import { readMarker, writeMarker, clearMarker } from '../src/docx/marker.js';

describe('zip', () => {
  it('round-trips every part', () => {
    const parts = makeDocx();
    const back = unzip(zip(parts));
    expect(Object.keys(back).sort()).toEqual(Object.keys(parts).sort());
    expect(readText(back, 'word/document.xml')).toBe(readText(parts, 'word/document.xml'));
  });

  it('keeps [Content_Types].xml first even when it was added last', () => {
    const parts = makeDocx();
    const shuffled = Object.fromEntries(Object.entries(parts).reverse());
    expect(Object.keys(shuffled)[0]).not.toBe(CONTENT_TYPES);
    expect(Object.keys(unzip(zip(shuffled)))[0]).toBe(CONTENT_TYPES);
  });

  it('rejects something that is not a docx', () => {
    const parts = makeDocx();
    delete parts['word/document.xml'];
    expect(isDocx(parts)).toBe(false);
  });
});

describe('marker', () => {
  it('reads null when the document is not a lay document', () => {
    expect(readMarker(makeDocx())).toBeNull();
    expect(readMarker(makeDocx({ custom: false }))).toBeNull();
  });

  it('round-trips through a real zip', () => {
    const parts = makeDocx();
    writeMarker(parts, 'sample-lay');
    expect(readMarker(unzip(zip(parts)))).toBe('sample-lay');
  });

  it('preserves cardmirror doc id and other properties', () => {
    const parts = makeDocx();
    writeMarker(parts, 'sample-lay');
    expect(readText(parts, 'docProps/custom.xml')).toContain('doc-abc-123');
    expect(readText(parts, 'docProps/custom.xml')).toContain('ContentTypeId');
  });

  it('does not duplicate the property when written twice', () => {
    const parts = makeDocx();
    writeMarker(parts, 'one');
    writeMarker(parts, 'two');
    const xml = readText(parts, 'docProps/custom.xml') ?? '';
    expect(xml.match(/name="layMirrorTemplate"/g)).toHaveLength(1);
    expect(readMarker(parts)).toBe('two');
  });

  it('assigns a pid that does not collide', () => {
    const parts = makeDocx();
    writeMarker(parts, 'sample-lay');
    const xml = readText(parts, 'docProps/custom.xml') ?? '';
    const pids = [...xml.matchAll(/pid="(\d+)"/g)].map((m) => Number(m[1]));
    expect(new Set(pids).size).toBe(pids.length);
    expect(Math.min(...pids)).toBeGreaterThanOrEqual(2);
  });

  it('escapes values that would break the xml', () => {
    const parts = makeDocx();
    writeMarker(parts, 'a & b <c> "d"');
    expect(readText(parts, 'docProps/custom.xml')).not.toContain('<c>');
    expect(readMarker(unzip(zip(parts)))).toBe('a & b <c> "d"');
  });

  it('creates and wires custom.xml when the donor has none', () => {
    const parts = makeDocx({ custom: false });
    writeMarker(parts, 'sample-lay');
    expect(readMarker(parts)).toBe('sample-lay');
    expect(readText(parts, CONTENT_TYPES)).toContain('/docProps/custom.xml');
    expect(readText(parts, '_rels/.rels')).toContain('docProps/custom.xml');
  });

  it('does not re-wire custom.xml that was already declared', () => {
    const parts = makeDocx({ custom: false });
    writeMarker(parts, 'one');
    writeMarker(parts, 'two');
    const ct = readText(parts, CONTENT_TYPES) ?? '';
    expect(ct.match(/docProps\/custom\.xml/g)).toHaveLength(1);
  });

  it('clearing leaves the document readable and other properties intact', () => {
    const parts = makeDocx();
    writeMarker(parts, 'sample-lay');
    clearMarker(parts);
    const back = unzip(zip(parts));
    expect(readMarker(back)).toBeNull();
    expect(readText(back, 'docProps/custom.xml')).toContain('doc-abc-123');
  });
});
