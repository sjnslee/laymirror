// @vitest-environment jsdom
//
// the whole pipeline against a real school template rather than a fixture.
// skipped unless `local/lay-template.docm` is present — it is gitignored,
// because a school's template is theirs.

import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { makeExport } from './fixture.js';
import { applyTemplate } from '../src/docx/apply.js';
import { readText, unzip } from '../src/docx/zip.js';
import { validateMapping } from '../src/template/styles.js';
import { read } from '../src/template/template.js';

const DONOR = 'local/lay-template.docm';
const suite = existsSync(DONOR) ? describe : describe.skip;

suite('a real school template', () => {
  const blueprint = () => {
    const result = read(readFileSync(DONOR), 'lay-template.docm');
    if (!result.ok) throw new Error(result.error);
    return result.blueprint;
  };

  const applied = (values = {}) =>
    unzip(applyTemplate(makeExport(), blueprint(), values, 'template:lay'));

  it('reads a macro-enabled template', () => {
    expect(read(readFileSync(DONOR), 'lay-template.docm').ok).toBe(true);
  });

  // this is where a lay file's page breaks come from — nobody types them
  it('breaks a page before the top-level headings', () => {
    expect(blueprint().breaks).toContain('pocket');
  });

  it('offers the header text a squad actually changes', () => {
    const labels = blueprint().fields.map((field) => field.label);
    expect(labels.length).toBeGreaterThan(0);
    // the live page count is word's to compute, never ours to offer
    expect(labels.join(' ')).not.toMatch(/\bof\b/);
  });

  it('maps every style onto something cardmirror can read back', () => {
    const { styles, styleMap, bareStyles } = blueprint();
    expect(validateMapping(styles, styleMap, bareStyles)).toEqual([]);
  });

  it('puts the school header onto a cardmirror export', () => {
    const parts = applied();
    expect(readText(parts, 'word/header1.xml')).toContain('PAGE');
    expect(readText(parts, 'word/document.xml')).toContain('headerReference');
  });

  it('carries the styles that make the file the school\'s', () => {
    expect(readText(applied(), 'word/styles.xml')).toContain('pageBreakBefore');
  });

  it('writes a team code the user typed into the real header', () => {
    const key = blueprint().fields[0]!.key;
    expect(readText(applied({ [key]: 'WDL 27-28' }), 'word/header1.xml')).toContain('WDL 27-28');
  });
});
