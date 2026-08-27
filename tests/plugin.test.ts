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
import { unzip, readText } from '../src/docx/zip.js';
import type { Command, PluginDefinition, PluginApi } from '../src/host/plugin-api.js';

const PATH = '/Users/x/Documents/1ac.docx';

interface Host {
  definition: PluginDefinition;
  api: PluginApi;
  toasts: string[];
  /** what is currently on disk at PATH. */
  disk: () => Uint8Array;
  run: (id: string) => Promise<void>;
}

let host: Host;

async function boot(): Promise<Host> {
  document.body.replaceChildren();
  document.head.replaceChildren();
  stubStorage();

  let disk = makeExport();
  const toasts: string[] = [];
  const bag: Record<string, unknown> = {};
  let definition: PluginDefinition | null = null;

  const api = {
    appVersion: '1.3.0',
    docInfo: () => null,
    showToast: (message: string) => void toasts.push(message),
    storage: { get: (key: string) => bag[key], set: (k: string, v: unknown) => void (bag[k] = v) },
    settings: { get: () => undefined, onChanged: () => () => {} },
  } as unknown as PluginApi;

  Object.assign(window as never, {
    __registerCardMirrorPlugin: (def: PluginDefinition) => void (definition = def),
    electronAPI: {
      statFile: async () => ({ mtimeMs: 1, size: disk.length }),
      readFileAtPath: async (path: string) =>
        path === PATH ? { name: '1ac.docx', bytes: disk, handle: PATH, format: 'docx' } : null,
      writeFileAtPath: async (path: string, bytes: Uint8Array) => {
        if (path === PATH) disk = bytes;
        return undefined;
      },
      openFile: async () => ({ name: 'lay.docx', bytes: makeTemplate(), handle: '/x/lay.docx' }),
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

  it('says where the template breaks its pages', async () => {
    await host.run('laymirror.panel');
    await click('load…');
    expect(panel()!.textContent).toContain('new page before every pocket');
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

  it('draws a rule above every block the template breaks before', async () => {
    await turnOn();
    const sheet = document.getElementById('laymirror-break-rules')!;
    expect(sheet.textContent).toContain('h1.pmd-pocket');
  });

  it('turns back off and leaves the file alone', async () => {
    await turnOn();
    await click('turn off');
    expect(host.toasts).toContain('lay formatting off');
    expect(document.getElementById('laymirror-break-rules')).toBeNull();
  });
});

describe('applying the header', () => {
  it('writes what was typed into the file', async () => {
    await host.run('laymirror.panel');
    await click('load…');
    const input = panel()!.querySelector('input') as HTMLInputElement;
    input.value = 'WDL 27-28';
    await click('apply');
    expect(readText(unzip(host.disk()), 'word/header1.xml')).toContain('WDL 27-28');
  });

  it('remembers it for the next time the panel opens', async () => {
    await host.run('laymirror.panel');
    await click('load…');
    (panel()!.querySelector('input') as HTMLInputElement).value = 'WDL 27-28';
    await click('apply');
    await host.run('laymirror.panel');
    await host.run('laymirror.panel');
    expect((panel()!.querySelector('input') as HTMLInputElement).value).toBe('WDL 27-28');
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
