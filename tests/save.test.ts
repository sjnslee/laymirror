// @vitest-environment jsdom
//
// the path the whole plugin exists for: cardmirror rebuilds the .docx on save,
// throwing the header, theme and page setup away, and laymirror puts them back
// without anyone pressing anything.
//
// nothing here runs a command. cardmirror only hands a plugin its api inside a
// command's `run()`, so a plugin that waits for one does nothing at all in a
// session where the user just opens a file and saves it — which is every
// session.

import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { makeExport, makeTemplate } from './fixture.js';
import { stubStorage } from './dom.js';
import { readText, unzip } from '../src/docx/zip.js';
import { encode } from '../src/state.js';

const PATH = '/Users/x/Documents/1ac.docx';
const TEMPLATE = 'template:lay.docx';

let disk: Uint8Array;
let mtime: number;

/** what cardmirror's storage holds for a document that was turned on in an
 *  earlier session — which is the state every launch after the first starts in. */
function seedStorage(): void {
  localStorage.setItem(
    'plugin:laymirror',
    JSON.stringify({
      templates: { [TEMPLATE]: { name: 'lay.docx', docx: encode(makeTemplate()) } },
      lastTemplate: TEMPLATE,
      docs: { '1ac.docx': { templateId: TEMPLATE, values: {}, on: true } },
    }),
  );
}

beforeEach(async () => {
  vi.useFakeTimers();
  document.body.replaceChildren();
  stubStorage();

  disk = makeExport();
  mtime = 1;

  Object.assign(window as never, {
    __registerCardMirrorPlugin: () => {},
    electronAPI: {
      statFile: async () => ({ mtimeMs: mtime, size: disk.length }),
      readFileAtPath: async (path: string) =>
        path === PATH ? { name: '1ac.docx', bytes: disk, handle: PATH, format: 'docx' } : null,
      writeFileAtPath: async (path: string, bytes: Uint8Array) => {
        if (path === PATH) {
          disk = bytes;
          mtime += 1;
        }
        return undefined;
      },
      openFile: async () => null,
    },
  });

  const chip = document.createElement('div');
  chip.id = 'doc-name-chip-text';
  chip.textContent = '1ac.docx';
  document.body.append(chip);
  localStorage.setItem(
    'pmd-recent-files',
    JSON.stringify([{ handle: PATH, filename: '1ac.docx', format: 'docx', lastOpenedAt: 2 }]),
  );

  seedStorage();
  vi.resetModules();
  await import('../src/main.js');
});

afterEach(() => {
  vi.useRealTimers();
});

/** long enough for the watcher to see a change twice at the same size, which
 *  is what it requires before it calls it a save. generous because jsdom's
 *  window is never focused, so the watcher runs at its backed-off interval. */
const settle = async (): Promise<void> => {
  await vi.advanceTimersByTimeAsync(15_000);
};

/** cardmirror saving: a whole new package, none of the school's document in it. */
function cardmirrorSaves(): void {
  disk = makeExport();
  mtime += 1;
}

it('puts the school header back after a save nobody told it about', async () => {
  await settle();
  cardmirrorSaves();
  expect(readText(unzip(disk), 'word/header1.xml')).toBeNull();

  await settle();
  expect(readText(unzip(disk), 'word/header1.xml')).toContain('PAGE');
  expect(readText(unzip(disk), 'word/document.xml')).toContain('headerReference');
});

it('restores the styles and theme the exporter dropped', async () => {
  await settle();
  cardmirrorSaves();
  await settle();

  const parts = unzip(disk);
  expect(readText(parts, 'word/styles.xml')).toContain('w:styleId="Tag"');
  expect(readText(parts, 'word/theme/theme1.xml')).not.toBeNull();
});

it('writes the header values held for the document', async () => {
  const bag = JSON.parse(localStorage.getItem('plugin:laymirror')!);
  const key = Object.keys(bag.templates)[0];
  expect(key).toBe(TEMPLATE);
  bag.docs['1ac.docx'].values = { 'word/header1.xml#0.0': 'WDL 27-28' };
  localStorage.setItem('plugin:laymirror', JSON.stringify(bag));

  await settle();
  cardmirrorSaves();
  await settle();

  expect(readText(unzip(disk), 'word/header1.xml')).toContain('WDL 27-28');
});

it('leaves a document that was never turned on alone', async () => {
  const bag = JSON.parse(localStorage.getItem('plugin:laymirror')!);
  bag.docs['1ac.docx'].on = false;
  localStorage.setItem('plugin:laymirror', JSON.stringify(bag));

  await settle();
  cardmirrorSaves();
  const before = disk;
  await settle();
  expect(disk).toBe(before);
});
