// @vitest-environment jsdom
// optional: runs only when a real school template is present at
// local/donor.docx, which is gitignored. the committed suite covers the same
// ground with a synthetic donor — this is here so a real template can be
// checked without ever entering the repo.
import { describe, it, expect } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { zip, unzip, isDocx, readText } from '../src/docx/zip.js';
import { readMarker, writeMarker, clearMarker } from '../src/docx/marker.js';
import { readTemplate, readAttachedTemplate } from '../src/profile/read-template.js';
import { validateMapping, takesNativePath } from '../src/profile/mapping.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { rewriteDocx } from '../src/docx/rewrite.js';
import { makeExport } from './fixture.js';

const PATH = 'local/donor.docx';
const present = existsSync(PATH);
const bytes = present ? new Uint8Array(readFileSync(PATH)) : new Uint8Array();

describe.skipIf(!present)('local donor template', () => {
  it('is a docx and starts unmarked', () => {
    const parts = unzip(bytes);
    expect(isDocx(parts)).toBe(true);
    expect(readMarker(parts)).toBeNull();
  });

  it('takes the marker without disturbing any other part', () => {
    const before = unzip(bytes);
    const parts = unzip(bytes);
    writeMarker(parts, 'local');
    const after = unzip(zip(parts));

    expect(readMarker(after)).toBe('local');
    expect(Object.keys(after)).toEqual(Object.keys(before));
    const changed = Object.keys(before).filter(
      (n) => Buffer.compare(Buffer.from(after[n]!), Buffer.from(before[n]!)) !== 0,
    );
    expect(changed).toEqual(['docProps/custom.xml']);
  });

  it('clears back to the original custom.xml exactly', () => {
    const parts = unzip(bytes);
    const original = readText(parts, 'docProps/custom.xml');
    writeMarker(parts, 'local');
    clearMarker(parts);
    expect(readText(parts, 'docProps/custom.xml')).toBe(original);
  });

  it('never carries an absolute template path into the profile', () => {
    const template = readAttachedTemplate(unzip(bytes));
    if (template !== null) {
      expect(template).not.toContain('/');
      expect(template).not.toContain('\\');
    }
  });

  it('maps onto cardmirror text types without losing any', () => {
    const { profile, missing } = readTemplate(bytes, DEFAULT_LAY);
    expect(missing).not.toContain('tag');
    expect(missing).not.toContain('card_body');
    expect(profile.page.widthTwips).toBeGreaterThan(0);

    const fatal = validateMapping(profile).filter((w) => w.type !== 'cite_mark');
    expect(fatal).toEqual([]);
  });

  it('rewrites a cardmirror export into the school format', () => {
    const { profile } = readTemplate(bytes, DEFAULT_LAY);
    const out = unzip(
      rewriteDocx(makeExport(), { ...profile, id: 'local' }, {
        title: '1AC',
        authors: 'A. Debater',
        teamCode: 'BCP 26-27',
      }),
    );

    // the school's own header, carried through untouched
    expect(readText(out, 'word/header1.xml')).toBe(profile.headerXml);
    expect(readMarker(out)).toBe('local');

    // every relationship the document points at has to exist, or word calls
    // the file corrupt rather than telling anyone why
    const document = readText(out, 'word/document.xml')!;
    const rels = readText(out, 'word/_rels/document.xml.rels')!;
    for (const [, id] of document.matchAll(/r:id="([^"]+)"/g)) {
      expect(rels).toContain(`Id="${id}"`);
    }

    // and every part the content types declare has to be in the package
    const types = readText(out, '[Content_Types].xml')!;
    for (const [, part] of types.matchAll(/PartName="\/([^"]+)"/g)) {
      expect(Object.keys(out)).toContain(part);
    }

    // cardmirror has to recognise its own on the way back in
    const styles = readText(out, 'word/styles.xml')!;
    const ids = [...styles.matchAll(/w:styleId="([^"]+)"/g)].map((m) => m[1]!);
    const names = [...styles.matchAll(/<w:name w:val="([^"]+)"/g)].map((m) => m[1]!);
    expect(takesNativePath(ids, names)).toBe(true);
  });
});
