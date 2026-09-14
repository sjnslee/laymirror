// @vitest-environment jsdom
//
// the plugin as cardmirror loads it: one script in the renderer's main world
// that registers itself. the only test running the panel, the commands and the
// save pipeline against each other.

import { beforeEach, describe, expect, it, vi } from 'vitest';
import { makeExport, makeTemplate } from './fixture.js';
import { stubStorage } from './dom.js';
import { readText, unzip, writeText, zip } from '../src/docx/zip.js';
import { writeMarker } from '../src/docx/marker.js';
import type { Command, PluginDefinition, PluginApi } from '../src/host/plugin-api.js';

const PATH = '/Users/x/Documents/1ac.docx';
const TEMPLATE_PATH = '/Users/x/Templates/lay.docx';

interface Host {
  definition: PluginDefinition;
  api: PluginApi;
  /** what laymirror's status line is currently showing. */
  said: () => string;
  /** what is currently on disk at PATH. */
  disk: () => Uint8Array;
  /** rewrite the template file the user picked, as editing it in word would.
   *  null deletes it. */
  editTemplate: (bytes: Uint8Array | null) => void;
  /** what the os picker hands back next, in place of the template file. */
  pickNext: (file: { name: string; bytes: Uint8Array; handle: string }) => void;
  run: (id: string) => Promise<void>;
}

let host: Host;

/** `listed: false` boots with an empty pmd-recent-files, which is what
 *  cardmirror leaves behind for a document it opened into a spawned window. */
async function boot({ listed = true } = {}): Promise<Host> {
  document.body.replaceChildren();
  document.head.replaceChildren();
  stubStorage();

  let disk = makeExport();
  let templateDisk: Uint8Array | null = makeTemplate();
  // a template is re-read only when its mtime moves, so an edit has to move it
  let templateMtime = 1;
  let nextPick: { name: string; bytes: Uint8Array; handle: string } | null = null;
  let definition: PluginDefinition | null = null;

  // one localStorage entry per plugin, which laymirror reads directly to start
  // watching. an in-memory bag would hide that half of the plugin
  const bag = (): Record<string, unknown> =>
    JSON.parse(localStorage.getItem('plugin:laymirror') || '{}');

  const api = {
    docInfo: () => null,
    storage: {
      get: (key: string) => bag()[key],
      set: (key: string, value: unknown) =>
        localStorage.setItem('plugin:laymirror', JSON.stringify({ ...bag(), [key]: value })),
    },
  } as unknown as PluginApi;

  Object.assign(window as never, {
    __registerCardMirrorPlugin: (def: PluginDefinition) => void (definition = def),
    electronAPI: {
      statFile: async (path: string) => {
        if (path === TEMPLATE_PATH) {
          return templateDisk ? { mtimeMs: templateMtime, size: templateDisk.length } : null;
        }
        return { mtimeMs: 1, size: disk.length };
      },
      readFileAtPath: async (path: string) => {
        if (path === PATH) return { name: '1ac.docx', bytes: disk, handle: PATH, format: 'docx' };
        // main serves .cmir and .docx only, and only paths the user put in play
        if (path === TEMPLATE_PATH && templateDisk) {
          return { name: 'lay.docx', bytes: templateDisk, handle: path, format: 'docx' };
        }
        return null;
      },
      saveExisting: async (path: string, bytes: Uint8Array) => {
        if (path === PATH) disk = bytes;
        return undefined;
      },
      openFile: async () => {
        const once = nextPick;
        nextPick = null;
        return once ?? { name: 'lay.docx', bytes: templateDisk ?? makeTemplate(), handle: TEMPLATE_PATH };
      },
    },
  });

  // the chip names the open document; the history turns that into a path
  const chip = document.createElement('div');
  chip.id = 'doc-name-chip-text';
  chip.textContent = '1ac.docx';
  document.body.append(chip);
  const editor = document.createElement('div');
  editor.className = 'ProseMirror';
  document.body.append(editor);
  localStorage.setItem(
    'pmd-recent-files',
    listed
      ? JSON.stringify([{ handle: PATH, filename: '1ac.docx', format: 'docx', lastOpenedAt: 2 }])
      : '[]',
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
    said: () => document.getElementById('laymirror-status')?.textContent ?? '',
    disk: () => disk,
    editTemplate: (bytes) => {
      templateDisk = bytes;
      templateMtime += 1;
    },
    pickNext: (file) => void (nextPick = file),
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

  // a document laymirror is not touching has no template, no header and
  // nothing written to it, and showing all three reads as if it did
  it('shows only the switch while lay formatting is off', async () => {
    await host.run('laymirror.panel');
    expect(panel()!.textContent).toContain('lay formatting is off');
    expect(panel()!.textContent).not.toContain('the file on disk');
    expect(panel()!.querySelectorAll('input')).toHaveLength(0);
    expect(buttons().map((b) => b.textContent)).toEqual([
      '×',
      'turn on',
      'locate\u2026',
      'diagnostics',
    ]);
  });

  it('opens the rest out once it is turned on', async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    const shown = panel()!.textContent!;
    expect(shown).toContain('template');
    expect(shown).toContain('file status');
  });

  it('offers the header fields once a template is loaded', async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    await click('load…');
    const labels = [...panel()!.querySelectorAll('label span')].map((el) => el.textContent);
    expect(labels).toEqual(['Team Code', 'lay']);
  });

  // turning it on before loading a template is the expected first step
  it('asks for a template rather than reporting a failure', async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    expect(host.said()).toBe('lay formatting on — load a template next');
    expect(panel()!.textContent).toContain('nothing written yet');
  });
});

describe('diagnostics', () => {
  const dialog = () => document.getElementById('laymirror-diagnose');
  const escape = () =>
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

  it('closes on escape', async () => {
    await host.run('laymirror.diagnose');
    expect(dialog()).not.toBeNull();
    escape();
    expect(dialog()).toBeNull();
  });

  it('stops listening once it is closed', async () => {
    await host.run('laymirror.diagnose');
    escape();
    await host.run('laymirror.panel');
    escape();
    expect(panel()).toBeNull();
  });
});

describe('turning it on', () => {
  const turnOn = async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    await click('load…');
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

  it('turns back off and leaves the file alone', async () => {
    await turnOn();
    await click('turn off');
    expect(host.said()).toBe('lay formatting off');
  });
});

describe('applying the header', () => {
  it('writes what was typed into the file', async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    await click('load…');
    const input = panel()!.querySelector('input') as HTMLInputElement;
    input.value = 'WDL 27-28';
    input.dispatchEvent(new Event('input'));
    await click('apply now');
    expect(readText(unzip(host.disk()), 'word/header1.xml')).toContain('WDL 27-28');
  });

  // an empty box means "the template's own text", so clearing a field has to
  // bring that back rather than whatever was typed before
  it('puts the template text back when a field is cleared', async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    await click('load…');
    const field = () => panel()!.querySelector('input') as HTMLInputElement;
    field().value = 'WDL 27-28';
    field().dispatchEvent(new Event('input'));
    await click('apply now');
    expect(readText(unzip(host.disk()), 'word/header1.xml')).toContain('WDL 27-28');

    field().value = '';
    field().dispatchEvent(new Event('input'));
    await click('apply now');
    const xml = readText(unzip(host.disk()), 'word/header1.xml')!;
    expect(xml).not.toContain('WDL 27-28');
    expect(xml).toContain('Team ');
    expect(field().value).toBe('');
  });

  // cardmirror's own save refuses a path this window does not have open, which
  // is what stops a same named file in another folder being rewritten
  it('writes nothing to a file this window does not have open', async () => {
    (window as unknown as { electronAPI: Record<string, unknown> }).electronAPI.saveExisting =
      async () => {
        throw new Error(
          `Error invoking remote method 'host:save-existing': Error: EMODIFIED: "1ac.docx" ` +
            'has no baseline in this window (it was not read here, or was restored without one)',
        );
      };
    const before = host.disk();
    await host.run('laymirror.panel');
    await click('turn on');
    await click('load…');
    expect(host.disk()).toBe(before);
    expect(host.said()).toContain('press locate');
  });

  it('remembers it for the next time the panel opens', async () => {
    await host.run('laymirror.panel');
    await click('turn on');
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
    await click('turn on');
    await click('load…');
  };

  const header = () => readText(unzip(host.disk()), 'word/header1.xml')!;

  /** the same template with a different word in its header, standing in for
   *  someone opening it in word and editing it. */
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

  // a template that moved, or a .docm cardmirror will not read from a path,
  // is not worth losing an apply over
  it('falls back to the stored copy when the file cannot be read', async () => {
    await load();
    host.editTemplate(null);
    await click('apply now');
    expect(header()).toContain('Team ');
  });
});

// cardmirror writes no history entry for a document it hands to a window it
// spawned, which is every open after the first.
describe('a document cardmirror never listed', () => {
  beforeEach(async () => {
    host = await boot({ listed: false });
  });

  // its state is kept under its path, so there is nothing to switch on until
  // the file is found
  it('says the file cannot be placed, not that it is the wrong kind', async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    expect(host.said()).toContain('press locate');
    expect(panel()!.textContent).toContain('lay formatting is off');
  });

  it('applies once the user has pointed at it', async () => {
    await host.run('laymirror.panel');
    host.pickNext({ name: '1ac.docx', bytes: new Uint8Array(), handle: PATH });
    await click('locate…');
    await click('turn on');
    await click('load…');
    await click('apply now');
    expect(readText(unzip(host.disk()), 'word/header1.xml')).toContain('Team ');
  });

  it('refuses a file that is not the open document', async () => {
    await host.run('laymirror.panel');
    host.pickNext({ name: 'somebody else.docx', bytes: new Uint8Array(), handle: '/x/y.docx' });
    await click('locate…');
    expect(host.said()).toContain('somebody else.docx');
    await click('turn on');
    expect(host.said()).toContain('press locate');
  });
});

describe('between documents and sessions', () => {
  const SECOND = '/Users/x/Documents/2ac.docx';

  const turnOnAndLoad = async () => {
    await host.run('laymirror.panel');
    await click('turn on');
    await click('load…');
  };

  /** cardmirror opening another file: the chip is repainted and the history
   *  gains an entry. */
  const openAnother = () => {
    document.getElementById('doc-name-chip-text')!.textContent = '2ac.docx';
    localStorage.setItem(
      'pmd-recent-files',
      JSON.stringify([
        { handle: SECOND, filename: '2ac.docx', format: 'docx', lastOpenedAt: 3 },
        { handle: PATH, filename: '1ac.docx', format: 'docx', lastOpenedAt: 2 },
      ]),
    );
  };

  it('gives a new document the template already in use', async () => {
    await turnOnAndLoad();
    await host.run('laymirror.panel');

    openAnother();
    await host.run('laymirror.panel');
    await click('turn on');
    expect(panel()!.textContent).toContain('lay.docx');
  });

  it('remembers the template and its path across a restart', async () => {
    await turnOnAndLoad();
    await host.run('laymirror.panel');

    // the same storage, a fresh module: what a relaunch looks like from here
    vi.resetModules();
    await import('../src/main.js');
    await host.run('laymirror.panel');

    const shown = panel()!.textContent!;
    expect(shown).toContain('lay.docx');
    expect(shown).toContain(TEMPLATE_PATH);
  });

  // localStorage.setItem throws over quota and cardmirror's storage bag
  // swallows it, so the template looked loaded until the next launch
  it('says so when the template will not fit in storage', async () => {
    const real = localStorage.setItem.bind(localStorage);
    await host.run('laymirror.panel');
    await click('turn on');
    localStorage.setItem = (key: string, value: string) => {
      if (key === 'plugin:laymirror' && value.includes('docx')) return;
      real(key, value);
    };
    await click('load…');
    localStorage.setItem = real;
    expect(host.said()).toContain('too large');
  });

  it('keeps two files that share a name apart', async () => {
    await turnOnAndLoad();
    await host.run('laymirror.panel');

    localStorage.setItem(
      'pmd-recent-files',
      JSON.stringify([
        { handle: '/Users/x/Elsewhere/1ac.docx', filename: '1ac.docx', format: 'docx', lastOpenedAt: 9 },
        { handle: PATH, filename: '1ac.docx', format: 'docx', lastOpenedAt: 2 },
      ]),
    );
    await host.run('laymirror.panel');
    expect(panel()!.textContent).toContain('lay formatting is off');
  });

  /** 2ac.docx as someone else's laymirror left it: marked with their template. */
  const markedWith = (templateId: string) => {
    const parts = unzip(makeExport());
    writeMarker(parts, templateId);
    const bytes = zip(parts);
    const electron = (window as unknown as { electronAPI: Record<string, unknown> }).electronAPI;
    const real = electron.readFileAtPath as (path: string) => Promise<unknown>;
    electron.readFileAtPath = async (path: string) =>
      path === SECOND ? { name: '2ac.docx', bytes, handle: SECOND, format: 'docx' } : real(path);
  };

  // the marker is bytes in a file, and one out of an email carries it as
  // readily as one a teammate sent: switching on starts rewriting that file
  it('does not switch a marked document on by itself', async () => {
    await turnOnAndLoad();
    await host.run('laymirror.panel');

    markedWith('template:lay.docx');
    openAnother();
    await host.run('laymirror.panel');
    await vi.waitFor(() => expect(host.said()).toContain('turn lay formatting on to keep it'));
    expect(panel()!.textContent).toContain('lay formatting is off');
  });

  // or a teammate's file is restyled with whatever this machine loaded last
  it('never gives a marked document a template other than its own', async () => {
    await turnOnAndLoad();
    await host.run('laymirror.panel');

    markedWith('template:theirs.docx');
    openAnother();
    await host.run('laymirror.panel');
    await vi.waitFor(() => expect(host.said()).toContain('not loaded here'));
    await click('turn on');
    expect(host.said()).toContain('load a template next');
  });

  // laymirror rewrites the file it is pointed at, so adopting one nobody asked
  // it to touch is the thing it must never do
  it('does not turn a new document on by itself', async () => {
    await turnOnAndLoad();
    await host.run('laymirror.panel');

    openAnother();
    await host.run('laymirror.panel');
    expect(panel()!.textContent).toContain('lay formatting is off');
  });
});

describe('holding what is typed', () => {
  const held = (): Record<string, string> => {
    const bag = JSON.parse(localStorage.getItem('plugin:laymirror') || '{}');
    return bag.docs?.[PATH]?.values ?? {};
  };

  const typeInto = async (text: string): Promise<HTMLInputElement> => {
    await host.run('laymirror.panel');
    await click('turn on');
    await click('load…');
    const input = panel()!.querySelector('input') as HTMLInputElement;
    input.value = text;
    input.dispatchEvent(new Event('input'));
    return input;
  };

  // one keystroke used to rewrite the whole bag, base64 template bytes and all,
  // synchronously on cardmirror's own thread
  it('does not write a bag on every keystroke', async () => {
    await typeInto('W');
    expect(held()).toEqual({});
  });

  it('writes once typing settles', async () => {
    await typeInto('WDL 27-28');
    await vi.waitFor(() => expect(Object.values(held())).toContain('WDL 27-28'));
  });

  it('writes immediately when a field is left', async () => {
    const input = await typeInto('WDL 27-28');
    input.dispatchEvent(new Event('blur'));
    expect(Object.values(held())).toContain('WDL 27-28');
  });

  // a panel closed mid word must not lose it
  it('writes what is held when the panel closes', async () => {
    await typeInto('WDL 27-28');
    await host.run('laymirror.panel');
    expect(Object.values(held())).toContain('WDL 27-28');
  });
});
