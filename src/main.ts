// cardmirror's exporter rebuilds the .docx on every save and keeps no header,
// footer or theme. laymirror watches the file and puts the template back on
// afterwards, with the user's own words in the header's editable text.

import { applyTemplate } from './docx/apply.js';
import { clearMarker, readMarker } from './docx/marker.js';
import { isDocx, unzip, zip } from './docx/zip.js';
import { currentFilename, enabledFlag } from './host/cardmirror.js';
import {
  DOCX_FILES,
  hasFileApi,
  openFile,
  readFile,
  saveExisting,
  statFile,
  WORD_FILES,
} from './host/electron.js';
import { resolveDocPath } from './host/paths.js';
import { bootApi, register, type PluginApi } from './host/plugin-api.js';
import { watchSaves, type Watcher } from './host/watcher.js';
import { store, TEMPLATE_LIMIT, type Store } from './state.js';
import { read, type Blueprint } from './template/template.js';
import { openDiagnostics } from './ui/diagnose.js';
import { say } from './ui/status.js';
import {
  closePanel,
  isOpen as panelOpen,
  openPanel,
  refresh,
  type Outcome,
} from './ui/panel.js';

const ID = 'laymirror';

/** how often laymirror notices the user has switched documents */
const SYNC_MS = 1500;

let watcher: Watcher | null = null;
let watching: string | null = null;
let syncing: ReturnType<typeof setInterval> | null = null;
/** parsing a template is a few ms of unzip; a save should not pay for it twice */
const blueprints = new Map<string, Blueprint>();

/** the name cardmirror shows for the open document. */
const docName = (): string | null => currentFilename();

/** the open document's full path, which is what its state is kept under. null
 *  when laymirror cannot say which file it is. */
function docKey(api: PluginApi): string | null {
  const found = locate(api);
  return 'path' in found ? found.path : null;
}

function blueprintFor(bag: Store, templateId: string | null): Blueprint | null {
  if (!templateId) return null;
  const cached = blueprints.get(templateId);
  if (cached) return cached;

  const template = bag.template(templateId);
  if (!template) return null;
  const result = read(template.docx, template.name);
  if (!result.ok) return null;
  blueprints.set(templateId, result.blueprint);
  return result.blueprint;
}

/** this document's template, or — when it has never had one — the last loaded */
const templateIdFor = (bag: Store, key: string | null): string | null =>
  bag.doc(key).templateId ?? bag.lastTemplateId();

type Found = { path: string } | { error: string };

const UNLISTED =
  'cardmirror has not said where this file is — press locate and point at it';

/** last path stored per name, so an unchanged answer does not rewrite the
 *  storage bag on every tick. */
const knownPath = new Map<string, string>();

/** keep the answer: the history holds ten files, and one that drops out of it
 *  while open still has to be found. */
function remember(api: PluginApi, path: string, at: number): void {
  const name = docName();
  if (!name || knownPath.get(name) === path) return;
  knownPath.set(name, path);
  store(api).setLocated(name, { path, at });
}

function locate(api: PluginApi): Found {
  if (!hasFileApi()) return { error: 'laymirror only works in the desktop app' };

  const bag = store(api);
  const found = resolveDocPath(api.docInfo(), (name) => bag.located(name));
  if (found.kind === 'ok') {
    remember(api, found.path, found.at);
    return { path: found.path };
  }
  if (found.kind === 'ambiguous') {
    return { error: 'two open files have this name, so laymirror cannot tell them apart' };
  }

  return {
    error:
      found.because === 'unlisted'
        ? UNLISTED
        : found.because === 'not-a-docx'
          ? 'this document is not a .docx — save it as one first'
          : 'no document is open',
  };
}

let last: Outcome | null = null;

/** cardmirror serves `readFileAtPath` only for `.cmir` and `.docx`, so a
 *  `.docm` or `.dotx` template can only ever come from the stored copy. */
const REREADABLE = /\.docx$/i;

/** when each template file was last taken off disk. keeping one re-encodes
 *  megabytes of base64, so an unchanged file is not taken again. */
const takenAt = new Map<string, number>();

/** take the template file again, picking up an edit made in word since it was
 *  loaded. doing nothing leaves the stored copy, which is still good. */
async function reread(bag: Store, templateId: string): Promise<void> {
  const info = bag.templateInfo(templateId);
  if (!info?.path || !REREADABLE.test(info.path)) return;

  const stat = await statFile(info.path).catch(() => null);
  if (stat && takenAt.get(templateId) === stat.mtimeMs) return;

  const file = await readFile(info.path);
  if (!file) return;

  const result = read(file.bytes, info.name);
  if (!result.ok) return;
  // grown past the room left: the stored copy is still good
  if (!bag.addTemplate({ id: templateId, name: info.name, path: info.path, docx: result.kept })) {
    return;
  }
  blueprints.set(templateId, result.blueprint);
  if (stat) takenAt.set(templateId, stat.mtimeMs);
}

const SAVED_MEANWHILE =
  'the document was saved again while laymirror was writing — nothing was overwritten';

/** why a write did not happen, in words a user can act on. */
function refused(err: unknown, path: string): string {
  const message = err instanceof Error ? err.message : String(err);
  if (/no baseline in this window/.test(message)) {
    return `this window does not have ${path} open — press locate and point at the open file`;
  }
  if (/EMODIFIED/.test(message)) return 'the file changed on disk — laymirror left it alone';
  return message;
}

const same = (a: Uint8Array, b: Uint8Array): boolean =>
  a.length === b.length && a.every((byte, i) => byte === b[i]);

type Written = { ok: true } | { ok: false; why: string };

/** read the document, change it, and write it back. cardmirror can save again
 *  while a large file is still being rewritten, and writing then would put the
 *  older contents over the newer save, so the file has to be as it was read. */
async function rewrite(
  path: string,
  change: (bytes: Uint8Array) => Uint8Array,
): Promise<Written> {
  const before = await statFile(path).catch(() => null);
  const file = await readFile(path);
  if (!before || !file) {
    return { ok: false, why: 'cardmirror would not let laymirror read the file' };
  }

  try {
    const bytes = change(file.bytes);
    // still a write to cardmirror, and on a synced folder a change for every
    // other machine with it open
    if (same(bytes, file.bytes)) return { ok: true };

    const now = await statFile(path).catch(() => null);
    if (!now || now.mtimeMs !== before.mtimeMs || now.size !== before.size) {
      // the newer save is the watcher's to pick up
      return { ok: false, why: SAVED_MEANWHILE };
    }

    await saveExisting(path, bytes);
    // absorb our own write, or the watcher reports it back as the user saving
    await watcher?.resync();
    return { ok: true };
  } catch (err) {
    return { ok: false, why: refused(err, path) };
  }
}

/** put the template onto the file on disk. failures are recorded rather than
 *  swallowed: a silent no-op and a working plugin look identical on screen. */
async function applyOnce(api: PluginApi, fresh = false): Promise<Outcome> {
  const bag = store(api);
  const found = locate(api);
  if ('error' in found) return record({ ok: false, why: found.error });

  const key = found.path;
  const templateId = templateIdFor(bag, key);
  if (!templateId) return record({ ok: false, why: 'no template loaded — load one first' });

  // asked for by hand: go back to the file first. a background save does not,
  // since the template cannot change between two keystrokes
  if (fresh) await reread(bag, templateId);

  const blueprint = blueprintFor(bag, templateId);
  if (!blueprint) return record({ ok: false, why: 'no template loaded — load one first' });

  const written = await rewrite(key, (bytes) =>
    applyTemplate(bytes, blueprint, bag.valuesFor(key, templateId), templateId),
  );
  if (!written.ok) return record(written);
  return record({
    ok: true,
    at: Date.now(),
    template: bag.templateInfo(templateId)?.name ?? templateId,
  });
}

/** one apply at a time: a save landing mid-apply otherwise reads the same file
 *  and writes it twice, and whichever finished last decides what the panel says. */
let queue: Promise<unknown> = Promise.resolve();

function apply(api: PluginApi, fresh = false): Promise<Outcome> {
  const next = queue.then(() => applyOnce(api, fresh));
  queue = next.catch(() => undefined);
  return next;
}

function record(outcome: Outcome): Outcome {
  last = outcome;
  if (!outcome.ok) console.warn('[laymirror] could not apply —', outcome.why);
  if (panelOpen()) refresh();
  return outcome;
}

/** apply and say what happened. */
async function applyAndReport(api: PluginApi, done: string, fresh = true): Promise<boolean> {
  const outcome = await apply(api, fresh);
  if (outcome.ok) say(done);
  else say(outcome.why, 'problem');
  return outcome.ok;
}

/** cardmirror has just rebuilt the file, so put the template back */
async function onSaved(api: PluginApi): Promise<void> {
  if (!store(api).doc(docKey(api)).on) return;
  const outcome = await apply(api);
  // a read that caught the file half-written, or a save that landed mid write,
  // is followed by a save the watcher picks up: not worth shouting about
  if (outcome.ok || outcome.why === SAVED_MEANWHILE) return;
  if (!/not a complete docx/.test(outcome.why)) say(outcome.why, 'problem');
}

/** take laymirror's marker off the file, so it stops offering itself. */
async function unmark(api: PluginApi): Promise<void> {
  const found = locate(api);
  if ('error' in found) {
    say(found.error, 'problem');
    return;
  }

  const written = await rewrite(found.path, (bytes) => {
    let parts;
    try {
      parts = unzip(bytes);
    } catch {
      throw new Error('the document is not readable as a docx right now');
    }
    // a partial read mid-save must abort, never round-trip into a write
    if (!isDocx(parts)) throw new Error('the document looks incomplete — try again in a moment');
    if (!readMarker(parts)) return bytes;
    clearMarker(parts);
    return zip(parts);
  });
  if (!written.ok) say(written.why, 'problem');
}

/** how many ticks a document gets to become reachable: the history is filled a
 *  moment late, but a `.cmir` never arrives and must not cost a read a tick. */
const ADOPT_TRIES = 4;

/** keys whose marker has been read, or given up on. added before the read: a
 *  tick can fire while the last one is still awaiting. */
const adopted = new Map<string, number>();

/** a document carrying laymirror's marker was formatted with a template,
 *  maybe on a teammate's machine. it does not switch itself on: the marker is
 *  just bytes in a file, and switching on starts rewriting that file. the
 *  document is tied to the marker's own template, never to whichever was loaded
 *  last, and the switch is left to the user. */
async function adopt(api: PluginApi, key: string): Promise<void> {
  const tries = (adopted.get(key) ?? 0) + 1;
  adopted.set(key, tries);

  const file = await readFile(key);
  if (!file) {
    if (tries < ADOPT_TRIES) adopted.delete(key);
    return;
  }

  let marker: string | null = null;
  try {
    marker = readMarker(unzip(file.bytes));
  } catch {
    return;
  }
  if (!marker) return;

  const bag = store(api);
  const doc = bag.doc(key);
  if (doc.on || doc.templateId) return;
  bag.setDoc(key, { templateId: marker });

  const name = docName() ?? 'this document';
  const template = marker.replace(/^template:/, '');
  say(
    bag.templateInfo(marker)
      ? `${name} was formatted with ${template} — turn lay formatting on to keep it`
      : `${name} was formatted with ${template}, which is not loaded here — load it first`,
  );
  if (panelOpen()) refresh();
}

/** the flag as this session found it. an installed laymirror loads only because
 *  its flag is true, and from then on reads it the way cardmirror does: anything
 *  but true is off, which is what an uninstall leaves. one loaded from a file
 *  has no flag, and runs until the session ends or a flag says false. */
const flagAtLoad = enabledFlag(ID);

const switchedOff = (): boolean => {
  const flag = enabledFlag(ID);
  return flagAtLoad === true ? flag !== true : flag === false;
};

/** everything laymirror has running, stopped. api v1 has no unload hook, so
 *  `sync` noticing the enabled flag is the only thing that can call this. */
function stop(): void {
  if (syncing !== null) clearInterval(syncing);
  syncing = null;
  watcher?.stop();
  watcher = null;
  watching = null;
  closePanel();
}

function sync(api: PluginApi): void {
  // switched off in settings: a disabled plugin must not keep rewriting files
  if (switchedOff()) {
    stop();
    return;
  }

  const bag = store(api);
  const key = docKey(api);

  if (key && (adopted.get(key) ?? 0) < ADOPT_TRIES) void adopt(api, key);

  if (!key || !bag.doc(key).on) {
    if (watching !== null) {
      watcher?.stop();
      watching = null;
    }
    return;
  }

  if (key === watching) return;
  watching = key;
  watcher?.start(key);
}

/** the api the background session runs on: a stand-in until a command hands
 *  over cardmirror's own. */
let session: PluginApi = bootApi(ID);

function ensureSession(api: PluginApi): void {
  session = api;
  if (!watcher) watcher = watchSaves(() => void onSaved(session));
  if (!syncing) syncing = setInterval(() => sync(session), SYNC_MS);
  sync(session);
}

async function toggleLay(api: PluginApi): Promise<void> {
  const bag = store(api);
  const found = locate(api);
  if ('error' in found) {
    say(found.error, 'problem');
    return;
  }

  const key = found.path;
  const on = !bag.doc(key).on;
  bag.setDoc(key, { on, templateId: templateIdFor(bag, key) });
  if (on) adopted.set(key, ADOPT_TRIES);
  sync(api);

  if (!on) {
    await unmark(api);
    last = null;
    say('lay formatting off');
    return;
  }

  // no template yet is the expected first step, not a failure
  if (!bag.templateInfo(templateIdFor(bag, key))) {
    say('lay formatting on — load a template next');
    return;
  }

  // apply now rather than at the next save: turning it on and seeing nothing
  // change looks the same as it not having worked
  await applyAndReport(api, 'lay formatting on — template applied');
}

async function loadTemplate(api: PluginApi): Promise<void> {
  const picked = await openFile(WORD_FILES);
  if (!picked) return;

  const result = read(picked.bytes, picked.name);
  if (!result.ok) {
    say(result.error, 'problem');
    return;
  }
  // measured after the cut: macros and sample cards are not kept, so a large
  // file can still be a small template
  if (result.kept.length > TEMPLATE_LIMIT) {
    say(`${picked.name} is too large to keep as a template`, 'problem');
    return;
  }

  const bag = store(api);
  const id = `template:${picked.name}`;
  if (!bag.addTemplate({ id, name: picked.name, path: picked.handle, docx: result.kept })) {
    say(`${picked.name} does not fit beside the templates documents already use`, 'problem');
    return;
  }

  // the bag swallows a failed localStorage write, so a template over quota
  // looks loaded until the next launch. read it back rather than trust it
  if (!bag.templateInfo(id)) {
    say(
      `${picked.name} is too large for cardmirror to keep — laymirror needs a smaller template`,
      'problem',
    );
    return;
  }
  blueprints.set(id, result.blueprint);
  const stat = await statFile(picked.handle).catch(() => null);
  if (stat) takenAt.set(id, stat.mtimeMs);

  const key = docKey(api);
  if (key) bag.setDoc(key, { templateId: id });
  sync(api);

  const fields = result.blueprint.fields.length;
  const found = `${picked.name} — ${fields} header field${fields === 1 ? '' : 's'}`;

  // waiting for the next save made loading a template look like a no-op
  if (key && bag.doc(key).on) {
    await applyAndReport(api, `${found}, applied`);
    return;
  }
  say(`${found} — turn lay formatting on to apply it`);
}

/** cardmirror never said where this document is, so ask. the picker is also
 *  what grants read scope on the path; a path from anywhere else has none. */
async function locateDoc(api: PluginApi): Promise<void> {
  const name = docName();
  if (!name) {
    say('no document is open', 'problem');
    return;
  }

  const picked = await openFile(DOCX_FILES);
  if (!picked) return;

  // a different file with the template written onto it is worse than no file
  if (picked.name !== name) {
    say(`that is ${picked.name}, and the open document is ${name}`, 'problem');
    return;
  }

  // newer than the history's entry, so it wins until the next open
  knownPath.set(name, picked.handle);
  store(api).setLocated(name, { path: picked.handle, at: Date.now() });
  sync(api);
  if (panelOpen()) refresh();
  say(`${name} found`);
}

function openLaymirror(api: PluginApi): void {
  ensureSession(api);
  if (panelOpen()) {
    closePanel();
    return;
  }

  const bag = () => store(api);
  const template = () => blueprintFor(bag(), templateIdFor(bag(), docKey(api)));

  openPanel({
    on: () => bag().doc(docKey(api)).on,
    templateName: () => bag().templateInfo(templateIdFor(bag(), docKey(api)))?.name ?? null,
    templatePath: () => bag().templateInfo(templateIdFor(bag(), docKey(api)))?.path ?? null,
    fields: () => template()?.fields ?? [],
    values: () => {
      const key = docKey(api);
      return bag().valuesFor(key, templateIdFor(bag(), key));
    },
    problem: () => {
      const located = locate(api);
      return 'error' in located ? located.error : null;
    },
    outcome: () => last,
    onToggle: () => toggleLay(api),
    onLoadTemplate: () => loadTemplate(api),
    // saved as typed, so a plain ⌘S picks up what is on screen
    onChange: (values) => {
      const key = docKey(api);
      if (key) bag().setValues(key, templateIdFor(bag(), key), values);
    },
    onApply: async (values) => {
      const key = docKey(api);
      if (key) bag().setValues(key, templateIdFor(bag(), key), values);
      await applyAndReport(api, 'template applied');
    },
    actions: [
      { label: 'locate…', run: () => locateDoc(api) },
      { label: 'diagnostics', run: () => openDiagnostics(api) },
    ],
  });
}

register({
  id: ID,
  name: 'laymirror',
  apiVersion: 1,
  commands: [
    {
      id: `${ID}.panel`,
      label: 'laymirror: open',
      keywords: ['lay', 'debate', 'template', 'header', 'settings'],
      // the alt chord is a fallback for when cardmirror has taken the first
      defaultKey: ['Mod-Shift-l', 'Mod-Alt-l'],
      run: (api) => openLaymirror(api),
    },
    {
      id: `${ID}.toggle-lay`,
      label: 'laymirror: turn lay formatting on or off',
      keywords: ['lay', 'debate', 'parent', 'judge'],
      run: async (api) => {
        ensureSession(api);
        await toggleLay(api);
      },
    },
    {
      id: `${ID}.apply`,
      label: 'laymirror: apply the template now',
      keywords: ['template', 'header', 'apply', 'format'],
      run: async (api) => {
        ensureSession(api);
        await applyAndReport(api, 'template applied');
      },
    },
    {
      id: `${ID}.locate`,
      label: 'laymirror: point at the open document on disk',
      keywords: ['locate', 'find', 'path', 'file', 'missing'],
      run: async (api) => {
        ensureSession(api);
        await locateDoc(api);
      },
    },
    {
      id: `${ID}.diagnose`,
      label: 'laymirror: diagnostics',
      keywords: ['debug', 'diagnose', 'why'],
      run: (api) => openDiagnostics(api),
    },
  ],
}) || console.warn('[laymirror] __registerCardMirrorPlugin unavailable');

// watch from the moment the script loads: the api arrives only inside a
// command's run(), and a document that was lay yesterday has to keep working
// without the user opening anything.
ensureSession(session);
