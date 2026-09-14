// @vitest-environment jsdom
//
// the path the whole plugin exists for: cardmirror rebuilds the .docx on save
// and laymirror puts the template back, with nobody pressing anything.
//
// nothing here runs a command, because a plugin waiting for the api object
// would do nothing in a session where the user opens a file and saves it.

import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { makeExport, makeTemplate } from './fixture.js';
import { stubStorage } from './dom.js';
import { readText, unzip, writeText, zip } from '../src/docx/zip.js';
import { encode } from '../src/state.js';

const PATH = '/Users/x/Documents/1ac.docx';
const TEMPLATE_PATH = '/Users/x/Templates/lay.docx';
const TEMPLATE = 'template:lay.docx';

let disk: Uint8Array;
let templateDisk: Uint8Array;
let mtime: number;

/** what cardmirror's storage holds for a document that was turned on in an
 *  earlier session — which is the state every launch after the first starts in. */
function seedStorage(): void {
  localStorage.setItem(
    'plugin:laymirror',
    JSON.stringify({
      templates: {
        [TEMPLATE]: { name: 'lay.docx', path: TEMPLATE_PATH, docx: encode(makeTemplate()) },
      },
      lastTemplate: TEMPLATE,
      docs: { [PATH]: { templateId: TEMPLATE, values: {}, on: true } },
    }),
  );
}

beforeEach(async () => {
  vi.useFakeTimers();
  document.body.replaceChildren();
  stubStorage();

  disk = makeExport();
  templateDisk = makeTemplate();
  mtime = 1;

  Object.assign(window as never, {
    __registerCardMirrorPlugin: () => {},
    electronAPI: {
      statFile: async () => ({ mtimeMs: mtime, size: disk.length }),
      readFileAtPath: async (path: string) => {
        if (path === PATH) return { name: '1ac.docx', bytes: disk, handle: PATH, format: 'docx' };
        if (path === TEMPLATE_PATH) {
          return { name: 'lay.docx', bytes: templateDisk, handle: path, format: 'docx' };
        }
        return null;
      },
      saveExisting: async (path: string, bytes: Uint8Array) => {
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

/** long enough for the watcher to see a change twice at the same size. jsdom
 *  is never focused, so it runs at its backed-off interval. */
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
  bag.docs[PATH].values = { 'word/header1.xml#0.0': 'WDL 27-28' };
  localStorage.setItem('plugin:laymirror', JSON.stringify(bag));

  await settle();
  cardmirrorSaves();
  await settle();

  expect(readText(unzip(disk), 'word/header1.xml')).toContain('WDL 27-28');
});

// an unzip per save buys nothing: the template cannot change between two
// keystrokes. going back to the file belongs to "apply now"
it('uses the stored template on a save rather than re-reading it', async () => {
  await settle();

  const parts = unzip(makeTemplate());
  writeText(
    parts,
    'word/header1.xml',
    readText(parts, 'word/header1.xml')!.replace('Team ', 'New '),
  );
  templateDisk = zip(parts);

  cardmirrorSaves();
  await settle();

  expect(readText(unzip(disk), 'word/header1.xml')).toContain('Team ');
});

it('leaves a document that was never turned on alone', async () => {
  const bag = JSON.parse(localStorage.getItem('plugin:laymirror')!);
  bag.docs[PATH].on = false;
  localStorage.setItem('plugin:laymirror', JSON.stringify(bag));

  await settle();
  cardmirrorSaves();
  const before = disk;
  await settle();
  expect(disk).toBe(before);
});

// api v1 has no unload hook and a bundle that has run cannot be unloaded, so
// laymirror has to notice on its own. a disabled plugin still rewriting files is
// worse than one that stops a tick late
it('stops writing once cardmirror switches it off', async () => {
  await settle();
  cardmirrorSaves();
  await settle();
  expect(readText(unzip(disk), 'word/header1.xml')).toContain('PAGE');

  localStorage.setItem('pmd-plugins', JSON.stringify({ enabled: { laymirror: false } }));
  await settle();

  cardmirrorSaves();
  const before = disk;
  await settle();
  expect(disk).toBe(before);
  expect(readText(unzip(disk), 'word/header1.xml')).toBeNull();
});

// installed, it reads the flag the way cardmirror does: an uninstall that
// clears it leaves laymirror off, not running until relaunch
it('stops once an installed copy loses its flag', async () => {
  // retire the copy this file booted, which loaded with no flag
  localStorage.setItem('pmd-plugins', JSON.stringify({ enabled: { laymirror: false } }));
  await settle();

  localStorage.setItem('pmd-plugins', JSON.stringify({ enabled: { laymirror: true } }));
  vi.resetModules();
  await import('../src/main.js');
  await settle();
  cardmirrorSaves();
  await settle();
  expect(readText(unzip(disk), 'word/header1.xml')).toContain('PAGE');

  localStorage.setItem('pmd-plugins', JSON.stringify({ enabled: {} }));
  await settle();
  cardmirrorSaves();
  await settle();
  expect(readText(unzip(disk), 'word/header1.xml')).toBeNull();
});

// "load plugin from file" writes no flag, and that is how it runs today
it('keeps working when no enabled flag has been written yet', async () => {
  localStorage.removeItem('pmd-plugins');
  await settle();
  cardmirrorSaves();
  await settle();
  expect(readText(unzip(disk), 'word/header1.xml')).toContain('PAGE');
});

type Electron = Record<string, (...args: never[]) => Promise<unknown>>;
const electron = (): Electron => (window as unknown as { electronAPI: Electron }).electronAPI;

// a save landing while laymirror is still rewriting must win: writing then
// would put the older contents back over it
it('leaves a save that lands mid rewrite in place', async () => {
  await settle();
  const real = electron().readFileAtPath as (path: string) => Promise<unknown>;
  let landed = false;
  electron().readFileAtPath = async (path: string) => {
    const read = await real(path);
    if (path === PATH && !landed) {
      landed = true;
      const parts = unzip(makeExport());
      const xml = readText(parts, 'word/document.xml')!;
      writeText(parts, 'word/document.xml', xml.replace('<w:body>', '<w:body><w:p><w:r><w:t>newer</w:t></w:r></w:p>'));
      disk = zip(parts);
      mtime += 1;
    }
    return read;
  };

  cardmirrorSaves();
  await settle();
  expect(readText(unzip(disk), 'word/document.xml')).toContain('newer');

  // and the newer save gets the template like any other
  await settle();
  expect(readText(unzip(disk), 'word/document.xml')).toContain('newer');
  expect(readText(unzip(disk), 'word/header1.xml')).toContain('PAGE');
});

// identical bytes are still a write, and every other machine syncing the file
// would see it as a change and hand one back
it('does not rewrite a file the template is already on', async () => {
  await settle();
  cardmirrorSaves();
  await settle();

  const real = electron().saveExisting as (path: string, bytes: Uint8Array) => Promise<unknown>;
  let writes = 0;
  electron().saveExisting = async (path: string, bytes: Uint8Array) => {
    writes += 1;
    return real(path, bytes);
  };
  // a sync client touching the file without changing it
  mtime += 1;
  await settle();
  expect(writes).toBe(0);
});
