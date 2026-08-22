// @vitest-environment jsdom
//
// the whole pipeline against a real school template rather than a fixture.
// skipped unless `local/donor.docx` is present — it is gitignored, because a
// school's template is theirs.

import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { makeExport } from './fixture.js';
import { applyProfile } from '../src/docx/rewrite.js';
import { readTemplate } from '../src/profile/read-template.js';
import { hasOwnHeader } from '../src/docx/snapshot.js';
import { validateMapping } from '../src/profile/mapping.js';
import { readText, unzip } from '../src/docx/zip.js';

const DONOR = 'local/donor.docx';
const present = existsSync(DONOR);
const suite = present ? describe : describe.skip;

suite('a real school template', () => {
  const profile = () => {
    const result = readTemplate(readFileSync(DONOR), 'donor.docx');
    if (!result.ok) throw new Error(result.error);
    return result.profile;
  };

  const restored = () => {
    const outcome = applyProfile(makeExport(), profile());
    if (outcome.kind !== 'restored') throw new Error(`expected restored, got ${outcome.kind}`);
    return unzip(outcome.bytes);
  };

  it('reads its header, footer, styles and theme', () => {
    const snapshot = profile().snapshot!;
    expect(snapshot.parts['word/header1.xml']).toBeDefined();
    expect(snapshot.parts['word/footer1.xml']).toBeDefined();
    expect(snapshot.parts['word/styles.xml']).toBeDefined();
    expect(snapshot.parts['word/theme/theme1.xml']).toBeDefined();
  });

  it('finds the school styles rather than word stock ones', () => {
    const map = profile().styleMap;
    // cardmirror exports a tag as Heading4; this school calls it Tag
    expect(map['Heading4']).toBe('Tag');
    expect(profile().bareStyles).toEqual({ cite_paragraph: 'Cite', card_body: 'card' });
  });

  it('keeps the school 0.5in sides and 0.7in bottom', () => {
    expect(profile().snapshot!.sectPr).toContain('w:bottom="1008"');
    expect(profile().snapshot!.sectPr).toContain('w:left="720"');
  });

  // this is the round-trip risk worth surfacing before a whole file is cut
  it('maps onto styles cardmirror can read back', () => {
    expect(validateMapping(profile())).toEqual([]);
  });

  it('adopts the template itself rather than rewriting it', () => {
    expect(applyProfile(readFileSync(DONOR), profile()).kind).toBe('adopted');
  });

  it('puts the school document onto a cardmirror export', () => {
    const parts = restored();
    expect(hasOwnHeader(parts)).toBe(true);
    const doc = readText(parts, 'word/document.xml')!;
    expect(doc).toContain('w:bottom="1008"');
    expect(doc).toContain('w:val="Tag"');
    expect(doc).toContain('w:val="Cite"');
    expect(doc).toContain('w:val="card"');
  });

  it('leaves a valid package — every part it declares is present', () => {
    const parts = restored();
    const ct = readText(parts, '[Content_Types].xml')!;
    for (const match of ct.matchAll(/PartName="\/([^"]+)"/g)) {
      expect(Object.keys(parts)).toContain(match[1]);
    }
  });

  it('leaves no dangling relationship for word to choke on', () => {
    const parts = restored();
    const doc = readText(parts, 'word/document.xml')!;
    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    for (const match of doc.matchAll(/r:id="([^"]+)"/g)) {
      expect(rels).toContain(`Id="${match[1]}"`);
    }
  });
});
