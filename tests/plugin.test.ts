// @vitest-environment jsdom
//
// the plugin as cardmirror loads it: one script, run in the renderer's main
// world, that registers itself and then has to survive being clicked on. this
// is the only test that exercises the panel, the commands and the save
// pipeline against each other.

import { beforeEach, describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { makeExport, makeTemplate } from './fixture.js';
import { stubStorage } from './dom.js';
import { readText, unzip, writeText, zip } from '../src/docx/zip.js';
import type { Command, PluginDefinition, PluginApi } from '../src/host/plugin-api.js';

const PATH = '/Users/x/Documents/1ac.docx';
const TEMPLATE_PATH = '/Users/x/Templates/lay.docx';

interface Host {
  definition: PluginDefinition;
  api: PluginApi;
  toasts: string[];
  /** what is currently on disk at PATH. */
  disk: () => Uint8Array;
  /** rewrite the template file the user picked, as editing it in word would. */
  editTemplate: (bytes: Uint8Array | null) => void;
  run: (id: string) => Promise<void>;
}

let host: Host;

async function boot(): Promise<Host> {
  document.body.replaceChildren();
  document.head.replaceChildren();
  stubStorage();

  let disk = makeExport();
  let templateDisk: Uint8Array | null = makeTemplate();
  const toasts: string[] = [];
  let definition: PluginDefinition | null = null;

  // cardmirror's own storage is one localStorage entry per plugin, and
  // laymirror reads it directly to start watching before any command has run.
  // an in-memory bag here would hide that half of the plugin from every test.
  const bag = (): Record<string, unknown> =>
    JSON.parse(localStorage.getItem('plugin:laymirror') || '{}');

  const api = {
    appVersion: '1.3.0',
    docInfo: () => null,
    showToast: (message: string) => void toasts.push(message),
    storage: {
      get: (key: string) => bag()[key],
      set: (key: string, value: unknown) =>
        localStorage.setItem('plugin:laymirror', JSON.stringify({ ...bag(), [key]: value })),
    },
    settings: { get: () => undefined, onChanged: () => () => {} },
  } as unknown as PluginApi;

  Object.assign(window as never, {
    __registerCardMirrorPlugin: (def: PluginDefinition) => void (definition = def),
    electronAPI: {
      statFile: async () => ({ mtimeMs: 1, size: disk.length }),
      readFileAtPath: async (path: string) => {
        if (path === PATH) return { name: '1ac.docx', bytes: disk, handle: PATH, format: 'docx' };
        // main serves .cmir and .docx only, and only paths the user put in play
        if (path === TEMPLATE_PATH && templateDisk) {
          return { name: 'lay.docx', bytes: templateDisk, handle: path, format: 'docx' };
        }
        return null;
      },
      writeFileAtPath: async (path: string, bytes: Uint8Array) => {
        if (path === PATH) disk = bytes;
        return undefined;
      },
      openFile: async () => ({
        name: 'lay.docx',
        bytes: templateDisk ?? makeTemplate(),
        handle: TEMPLATE_PATH,
      }),
    },
  });

  // cardmirror names the open document in its chip and turns it into a path
  // through its recent-files history
  const chip = document.createElement('div');
  chip.id = 'doc-name-chip-text';
  chip.textContent = '1ac.docx';
  document.body.append(chip);
  const editor = document.createElement('div');
  editor.className = 'ProseMirror';
  document.body.append(editor);
  localStorage.setItem(
    'pmd-recent-files',
    JSON.stringify([{ handle: PATH, filename: '1ac.docx', format: 'docx', lastOpenedAt: 2 }]),
  );

  vi.resetModules();
  await import('../src/main.js');
  if (!definition) throw new Error('the plugin did not register');

  const commands = new Map<string, Command>(
    (definition as PluginDefinition).commands.map((command) => [command.id, command]),
  );
  return {
    definition: definition as PluginDefinition,
    api,
    toasts,
    disk: () => disk,
    editTemplate: (bytes) => void (templateDisk = bytes),
    run: async (id) => {
      await commands.get(id)!.run(api);
    },
  };
}

const panel = () => document.getElementById('laymirror-panel');
const buttons = () =>
  [...(panel()?.querySelectorAll('button') ?? [])] as HTMLButtonElement[];
const click = async (label: string): Promise<void> => {
  const button = buttons().find((b) => b.textContent === label);
  if (!button) throw new Error(`no "${label}" button — saw ${buttons().map((b) => b.textContent)}`);
  button.click();
  await vi.waitFor(() => expect(panel()).not.toBeNull());
  await new Promise((resolve) => setTimeout(resolve, 0));
};

beforeEach(async () => {
  host = await boot();
});

describe('registration', () => {
  it('registers under its own id, at api version 1', () => {
    expect(host.definition.id).toBe('laymirror');
    expect(host.definition.apiVersion).toBe(1);
  });

  // cardmirror rejects the whole plugin if one command id is wrong
  it('prefixes every command id with the plugin id', () => {
    for (const command of host.definition.commands) {
      expect(command.id.startsWith('laymirror.')).toBe(true);
      expect(command.id.length).toBeGreaterThan('laymirror.'.length);
    }
  });

  it('arrives with the menu already bound to cmd-shift-l', () => {
    const open = host.definition.commands.find((c) => c.id === 'laymirror.panel')!;
    expect(open.defaultKey).toContain('Mod-Shift-l');
  });
});

describe('the panel', () => {
  it('opens, and closes on a second press', async () => {
    await host.run('laymirror.panel');
    expect(panel()).not.toBeNull();
    await host.run('laymirror.panel');
    expect(panel()).toBeNull();
  });

  it('says there is no template until one is loaded', async () => {
    await host.run('laymirror.panel');
    expect(panel()!.textContent).toContain('none');
  });

  it('offers the header fields once a template is loaded', async () => {
    await host.run('laymirror.panel');
    await click('load…');
    const labels = [...panel()!.querySelectorAll('label span')].map((el) => el.textContent);
    expect(labels).toEqual(['Team Code', 'lay']);
  });

  it('says nothing has been written until something has', async () => {
    await host.run('laymirror.panel');
    expect(panel()!.textContent).toContain('nothing written yet');
  });
});

describe('turning it on', () => {
  const turnOn = async () => {
    await host.run('laymirror.panel');
    await click('load…');
    await click('turn on');
  };

  it('puts the school header onto the file straight away', async () => {
    await turnOn();
    expect(readText(unzip(host.disk()), 'word/header1.xml')).toContain('PAGE');
  });

  it('carries the template styles, theme and page setup too', async () => {
    await turnOn();
    const parts = unzip(host.disk());
    expect(readText(parts, 'word/styles.xml')).toContain('w:styleId="Tag"');
    expect(readText(parts, 'word/theme/theme1.xml')).not.toBeNull();
    expect(readText(parts, 'word/document.xml')).toContain('headerReference');
  });

  // loading a template onto a document that is already lay used to change
  // nothing until the next save, which read as the feature not working at all
  it('applies a template loaded after it was turned on', async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    await click('load…');
    expect(readText(unzip(host.disk()), 'word/header1.xml')).toContain('PAGE');
  });

  it('turns back off and leaves the file alone', async () => {
    await turnOn();
    await click('turn off');
    expect(host.toasts).toContain('lay formatting off');
  });
});

describe('applying the header', () => {
  it('writes what was typed into the file', async () => {
    await host.run('laymirror.panel');
    await click('load…');
    const input = panel()!.querySelector('input') as HTMLInputElement;
    input.value = 'WDL 27-28';
    input.dispatchEvent(new Event('input'));
    await click('apply now');
    expect(readText(unzip(host.disk()), 'word/header1.xml')).toContain('WDL 27-28');
  });

  it('remembers it for the next time the panel opens', async () => {
    await host.run('laymirror.panel');
    await click('load…');
    (panel()!.querySelector('input') as HTMLInputElement).value = 'WDL 27-28';
    await click('apply now');
    await host.run('laymirror.panel');
    await host.run('laymirror.panel');
    expect((panel()!.querySelector('input') as HTMLInputElement).value).toBe('WDL 27-28');
  });
});

describe('re-reading the template', () => {
  const load = async () => {
    await host.run('laymirror.panel');
    await click('load…');
  };

  const header = () => readText(unzip(host.disk()), 'word/header1.xml')!;

  /** the same template with a different word in its header, standing in for
   *  someone opening it in word and changing the school's name. */
  const edited = (): Uint8Array => {
    const parts = unzip(makeTemplate());
    writeText(
      parts,
      'word/header1.xml',
      readText(parts, 'word/header1.xml')!.replace('Team ', 'New '),
    );
    return zip(parts);
  };

  it('takes the template file again when apply is pressed', async () => {
    await load();
    host.editTemplate(edited());
    await click('apply now');
    expect(header()).toContain('New ');
  });

  it('says it went back to the file', async () => {
    await load();
    await click('apply now');
    expect(panel()!.textContent).toContain('re-read from the template file');
  });

  // a template that moved, or a .docm — which cardmirror will not read back
  // from a path at all — is not worth losing an apply over
  it('falls back to the stored copy when the file cannot be read', async () => {
    await load();
    host.editTemplate(null);
    await click('apply now');
    expect(header()).toContain('Team ');
    expect(panel()!.textContent).toContain('from the stored copy');
  });

});

describe('the built bundle', () => {
  // esbuild has to produce one self-contained classic script, because that is
  // all cardmirror will run
  it('is an iife with no import left in it', () => {
    const bundle = readFileSync('plugin.js', 'utf8');
    expect(bundle).not.toMatch(/^\s*import\s/m);
    expect(bundle).not.toMatch(/^\s*export\s/m);
    expect(bundle).toContain('__registerCardMirrorPlugin');
  });
});
