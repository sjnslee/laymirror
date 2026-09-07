// @vitest-environment jsdom
import { describe, it, expect, beforeEach } from 'vitest';
import { resolveDocPath } from '../src/host/paths.js';
import { LS, DOC_NAME_CHIP } from '../src/host/cardmirror.js';
import { stubStorage } from './dom.js';

/** cardmirror's history list, newest first, as `pa()` maintains it. */
const recents = (entries: unknown[]) =>
  localStorage.setItem(LS.recents, JSON.stringify(entries));

const entry = (filename: string, handle: string | null, lastOpenedAt = 1, format = 'docx') => ({
  filename,
  handle,
  format,
  lastOpenedAt,
});

/** the chip cardmirror paints the open document's name into. */
function showing(filename: string): void {
  const chip = document.createElement('span');
  chip.id = DOC_NAME_CHIP;
  chip.textContent = filename;
  document.body.append(chip);
}

beforeEach(() => {
  stubStorage();
  document.body.replaceChildren();
  document.title = '';
});

describe('resolveDocPath', () => {
  it('finds the open document among everything ever opened', () => {
    // a history, not a list of what is open: every entry looks like a candidate
    recents([
      entry('1ac.docx', '/docs/1ac.docx', 5),
      entry('neg block.docx', '/docs/neg block.docx', 4),
      entry('old case.docx', '/docs/old case.docx', 3),
    ]);
    showing('neg block.docx');

    expect(resolveDocPath(null)).toEqual({ kind: 'ok', path: '/docs/neg block.docx' });
  });

  it('works with no docInfo at all, which is the normal case', () => {
    // docInfo() is null until cardmirror has saved the file itself
    recents([entry('1ac.docx', '/docs/1ac.docx')]);
    showing('1ac.docx');

    expect(resolveDocPath(null)).toEqual({ kind: 'ok', path: '/docs/1ac.docx' });
  });

  it('falls back to the window title when the chip is not there', () => {
    recents([entry('1ac.docx', '/docs/1ac.docx')]);
    document.title = '1ac.docx — CardMirror';

    expect(resolveDocPath(null)).toEqual({ kind: 'ok', path: '/docs/1ac.docx' });
  });

  it('takes docTitle when cardmirror does have an id for the document', () => {
    recents([entry('1ac.docx', '/docs/1ac.docx')]);
    expect(resolveDocPath({ docId: 'x', docTitle: '1ac.docx' })).toEqual({
      kind: 'ok',
      path: '/docs/1ac.docx',
    });
  });

  it('says so plainly when the open document is not a docx', () => {
    recents([entry('case.cmir', '/docs/case.cmir', 1, 'cmir')]);
    showing('case.cmir');

    expect(resolveDocPath(null)).toEqual({ kind: 'none', because: 'not-a-docx' });
  });

  it('never guesses at another file when this one has no path', () => {
    // a browser-opened document has no handle, and rewriting the neighbour on
    // the list would be worse than doing nothing
    recents([entry('1ac.docx', null), entry('other.docx', '/docs/other.docx')]);
    showing('1ac.docx');

    expect(resolveDocPath(null)).toEqual({ kind: 'none', because: 'unlisted' });
  });

  it('separates a docx with no entry from a document laymirror cannot touch', () => {
    // a .docx with no entry can still be pointed at; a .cmir never can
    recents([entry('other.docx', '/docs/other.docx')]);
    showing('1ac.docx');
    expect(resolveDocPath(null)).toEqual({ kind: 'none', because: 'unlisted' });

    document.body.replaceChildren();
    showing('case.cmir');
    expect(resolveDocPath(null)).toEqual({ kind: 'none', because: 'not-a-docx' });
  });

  it('reports no document when nothing names one', () => {
    recents([entry('1ac.docx', '/docs/1ac.docx')]);
    expect(resolveDocPath(null)).toEqual({ kind: 'none', because: 'no-document' });
  });

  it('prefers the most recently opened of two files with one name', () => {
    recents([
      entry('1ac.docx', '/a/1ac.docx', 9),
      entry('1ac.docx', '/b/1ac.docx', 2),
    ]);
    showing('1ac.docx');

    expect(resolveDocPath(null)).toEqual({ kind: 'ok', path: '/a/1ac.docx' });
  });

  it('gives up only when two of one name are genuinely tied', () => {
    recents([entry('1ac.docx', '/a/1ac.docx', 7), entry('1ac.docx', '/b/1ac.docx', 7)]);
    showing('1ac.docx');

    expect(resolveDocPath(null)).toEqual({
      kind: 'ambiguous',
      paths: ['/a/1ac.docx', '/b/1ac.docx'],
    });
  });

  it('survives a recents list that is missing or corrupt', () => {
    localStorage.setItem(LS.recents, 'not json');
    showing('1ac.docx');
    expect(resolveDocPath(null)).toEqual({ kind: 'none', because: 'unlisted' });
  });
});
