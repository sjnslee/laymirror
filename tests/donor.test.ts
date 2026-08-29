// @vitest-environment jsdom
//
// the whole pipeline against the real school template rather than a fixture.
// skipped unless `local/lay-template.docm` is present — it is gitignored,
// because a school's template is theirs.
//
// this template marks its editable text with zero-width spaces, so it is the
// only thing that exercises the marked field path against something a school
// actually wrote.

import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { makeExport } from './fixture.js';
import { applyTemplate } from '../src/docx/apply.js';
import { readText, unzip } from '../src/docx/zip.js';
import { validateMapping } from '../src/template/styles.js';
import { read } from '../src/template/template.js';

const DONOR = 'local/lay-template.docm';
const suite = existsSync(DONOR) ? describe : describe.skip;

const ZWSP = '​';

suite('a real school template', () => {
  const blueprint = () => {
    const result = read(readFileSync(DONOR), 'lay-template.docm');
    if (!result.ok) throw new Error(result.error);
    return result.blueprint;
  };

  const applied = (values = {}) =>
    unzip(applyTemplate(makeExport(), blueprint(), values, 'template:lay'));

  const header = (values = {}) => readText(applied(values), 'word/header1.xml')!;

  it('reads a macro-enabled template', () => {
    expect(read(readFileSync(DONOR), 'lay-template.docm').ok).toBe(true);
  });

  // the four things a squad changes, and nothing else. the school wrapped each
  // of them in a zero-width space; everything outside those is the template's
  it('offers exactly the text the school marked as editable', () => {
    expect(blueprint().fields.map((field) => field.label)).toEqual([
      'School',
      '26-27',
      'File Title',
      'Name',
    ]);
  });

  // ' Page ' and ' of ' read as plain text but belong to the PAGE/NUMPAGES
  // fields, and the numbers between them are word's to recompute
  it('leaves the live page numbering alone', () => {
    const labels = blueprint().fields.map((field) => field.label);
    expect(labels.join(' ')).not.toMatch(/\bof\b|\bPage\b/i);
    expect(header()).toContain('NUMPAGES');
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

  it("carries the styles that make the file the school's", () => {
    expect(readText(applied(), 'word/styles.xml')).toContain('pageBreakBefore');
  });

  it('writes every field the user typed into the real header', () => {
    const [school, year, title, name] = blueprint().fields;
    const xml = header({
      [school!.key]: 'WDL',
      [year!.key]: '27-28',
      [title!.key]: 'Aff Case Neg',
      [name!.key]: 'Shane Lee',
    });
    for (const value of ['WDL', '27-28', 'Aff Case Neg', 'Shane Lee']) {
      expect(xml).toContain(value);
    }
    // the school's own words are gone, not sitting behind the new ones
    expect(xml).not.toContain('File Title');
  });

  // a value is written between the marks, never over them. lose one and the
  // field stops existing the next time the template is read
  it('keeps every marker, so a value can be typed over', () => {
    const marks = (xml: string) => [...xml].filter((c) => c === ZWSP).length;
    const key = blueprint().fields[1]!.key;
    expect(marks(header())).toBe(8);
    expect(marks(header({ [key]: '27-28' }))).toBe(8);
  });

  // the year arrives split across two runs ('2' then '6-27'); the value has to
  // land whole in the first and empty the rest, or it comes out doubled
  it('fills a field that the school split across runs', () => {
    const key = blueprint().fields[1]!.key;
    const xml = header({ [key]: '27-28' });
    expect(xml).toContain('27-28');
    expect(xml).not.toContain('6-27');
  });
});
