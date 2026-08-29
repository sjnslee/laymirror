// laymirror — lay debate documents for cardmirror.
//
// cardmirror's exporter rebuilds the .docx from scratch on every save and
// keeps no header, footer or theme. laymirror watches the file, and every time
// the exporter runs it puts the school's document back on: the template's
// styles, fonts, page setup and header, with the user's own words filled into
// the header's editable text.
//
// off-state leaves the file alone. a document that has never been turned on is
// read once to see whether it carries the marker, and never written.

import { applyTemplate } from './docx/apply.js';
import { clearMarker, readMarker } from './docx/marker.js';
import { isDocx, unzip, zip, type Parts } from './docx/zip.js';
import { currentFilename } from './host/cardmirror.js';
import {
  hasFileApi,
  openFile,
  readFile,
  writeFile,
  WORD_FILES,
} from './host/electron.js';
import { resolveDocPath } from './host/paths.js';
import { bootApi, register, type PluginApi } from './host/plugin-api.js';
import { watchSaves, type Watcher } from './host/watcher.js';
import { store, TEMPLATE_LIMIT, type Store } from './state.js';
import { read, type Blueprint } from './template/template.js';
import { openDiagnostics } from './ui/diagnose.js';
import {
  closePanel,
  isOpen as panelOpen,
  openPanel,
  refresh,
  type Outcome,
  type Source,
} from './ui/panel.js';

const ID = 'laymirror';

/** how often laymirror notices the user has switched documents. cheap: a
 *  string compare against the filename chip, and nothing else unless it moved. */
const SYNC_MS = 1500;

let watcher: Watcher | null = null;
let watching: string | null = null;
let syncing: ReturnType<typeof setInterval> | null = null;
/** parsing a template is a few milliseconds of unzip; a save should not pay
 *  for it twice. */
const blueprints = new Map<string, Blueprint>();

/** documents are keyed by name rather than by path: a path can be missing
 *  while the document is plainly open, and losing the key would lose the
 *  template and the header values with it. */
const docKey = (): string | null => currentFilename();

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

/** the template this document should wear: its own, or — for a document that
 *  has never had one — whichever was loaded last, which is nearly always the
 *  right guess. */
const templateIdFor = (bag: Store, key: string | null): string | null =>
  bag.doc(key).templateId ?? bag.lastTemplateId();

type Located = { path: string } | { error: string };

function locate(api: PluginApi): Located {
  if (!hasFileApi()) return { error: 'laymirror only works in the desktop app' };

  const found = resolveDocPath(api.docInfo());
  if (found.kind === 'ok') return { path: found.path };
  if (found.kind === 'ambiguous') {
    return { error: 'two open files have this name, so laymirror cannot tell them apart' };
  }
  return {
    error:
      found.because === 'not-a-docx'
        ? 'save this document as a .docx first'
        : 'no document is open',
  };
}

// ── applying the template ─────────────────────────────────────────────

let last: Outcome | null = null;

/** main serves `readFileAtPath` only for `.cmir` and `.docx`, so a `.docm` or
 *  `.dotx` template — which is most of them — can never be read back from its
 *  path. those keep working from the stored copy; re-loading the template is
 *  how they pick up an edit. */
const REREADABLE = /\.docx$/i;

/** go back to the template file the user picked and take it again, so an edit
 *  made in word since it was loaded is picked up.
 *
 *  false when there is no path, the format cannot be read back, the file has
 *  moved, or what came back does not parse — in every one of those the stored
 *  copy is still good, and refusing to apply would be worse than applying the
 *  version we already had. */
async function reread(bag: Store, templateId: string): Promise<boolean> {
  const template = bag.template(templateId);
  if (!template?.path || !REREADABLE.test(template.path)) return false;

  const file = await readFile(template.path);
  if (!file) return false;

  const result = read(file.bytes, template.name);
  if (!result.ok) return false;

  bag.addTemplate({ ...template, docx: file.bytes });
  blueprints.set(templateId, result.blueprint);
  return true;
}

/** put the school's document onto the file on disk.
 *
 *  every failure is recorded and returned rather than swallowed: laymirror
 *  writes a file nobody is looking at, so a silent no-op and a working plugin
 *  are indistinguishable from inside cardmirror. */
async function apply(api: PluginApi, fresh = false): Promise<Outcome> {
  const bag = store(api);
  const key = docKey();
  const templateId = templateIdFor(bag, key);

  if (!key) return record({ ok: false, why: 'no document is open' });
  if (!templateId) return record({ ok: false, why: 'no template loaded — load one first' });

  // asked for by hand: go back to the file first. a background save does not,
  // because the template does not change between two keystrokes and an unzip
  // per save is a cost with nothing to show for it
  const source: Source = fresh && (await reread(bag, templateId)) ? 'file' : 'stored';

  const blueprint = blueprintFor(bag, templateId);
  if (!blueprint) return record({ ok: false, why: 'no template loaded — load one first' });

  const located = locate(api);
  if ('error' in located) return record({ ok: false, why: located.error });

  const file = await readFile(located.path);
  if (!file) {
    return record({ ok: false, why: 'cardmirror would not let laymirror read the file' });
  }

  try {
    const bytes = applyTemplate(
      file.bytes,
      blueprint,
      bag.valuesFor(key, templateId),
      templateId,
    );
    await writeFile(located.path, bytes);
    // absorb our own write, or the watcher reports it back as the user saving
    await watcher?.resync();
    return record({
      ok: true,
      at: Date.now(),
      template: bag.template(templateId)?.name ?? templateId,
      fields: blueprint.fields.length,
      source,
    });
  } catch (err) {
    return record({ ok: false, why: err instanceof Error ? err.message : String(err) });
  }
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
  api.showToast(outcome.ok ? done : `laymirror: ${outcome.why}`);
  return outcome.ok;
}

/** cardmirror has just rebuilt the file from scratch, so put the school's
 *  document back onto it. */
async function onSaved(api: PluginApi): Promise<void> {
  if (!store(api).doc(docKey()).on) return;
  const outcome = await apply(api);
  // a read that caught the file half-written is not worth shouting about: the
  // next save lands on a whole file. everything else the user needs to know.
  if (!outcome.ok && !/not a complete docx/.test(outcome.why)) {
    api.showToast(`laymirror: ${outcome.why}`);
  }
}

/** read the open document, hand its parts to `edit`, write it back. */
async function withOpenDocx(
  api: PluginApi,
  edit: (parts: Parts) => void,
): Promise<boolean> {
  const located = locate(api);
  if ('error' in located) {
    api.showToast(located.error);
    return false;
  }

  const file = await readFile(located.path);
  if (!file) {
    api.showToast('could not read the document — reopen it and try again');
    return false;
  }

  let parts: Parts;
  try {
    parts = unzip(file.bytes);
  } catch {
    api.showToast('the document is not readable as a docx right now');
    return false;
  }
  // a partial read mid-save must abort, never round-trip into a write
  if (!isDocx(parts)) {
    api.showToast('the document looks incomplete — try again in a moment');
    return false;
  }

  edit(parts);
  await writeFile(located.path, zip(parts));
  await watcher?.resync();
  return true;
}

// ── keeping the watcher on the right document ─────────────────────────

/** how many ticks a document gets to become reachable before laymirror stops
 *  reading it. cardmirror fills its recent-files history a moment after the
 *  document appears, so the first attempt often misses — but a `.cmir`, which
 *  never becomes reachable, must not cost a file read every 1.5 seconds. */
const ADOPT_TRIES = 4;

/** document keys whose marker has been read, or given up on. a key is added
 *  before the read rather than after, because a tick can fire while the last
 *  one is still awaiting. */
const adopted = new Map<string, number>();

/** a document carrying laymirror's marker turns itself on. the marker travels
 *  inside the .docx, so a file a teammate marked arrives already lay — which is
 *  the whole reason it is in the file rather than in this machine's storage.
 *
 *  a document we could not reach is retried for a few ticks: cardmirror fills
 *  its recent-files history a moment after the document appears, and giving up
 *  on that moment would leave the file plain. */
async function adopt(api: PluginApi, key: string): Promise<void> {
  const tries = (adopted.get(key) ?? 0) + 1;
  adopted.set(key, tries);

  const retry = (): void => {
    if (tries < ADOPT_TRIES) adopted.delete(key);
  };

  const located = locate(api);
  if ('error' in located) {
    retry();
    return;
  }

  const file = await readFile(located.path);
  if (!file) {
    retry();
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
  if (bag.doc(key).on) return;
  bag.setDoc(key, { on: true, templateId: bag.template(marker) ? marker : bag.doc(key).templateId });
  sync(api);
}

function sync(api: PluginApi): void {
  const bag = store(api);
  const key = docKey();

  if (key && (adopted.get(key) ?? 0) < ADOPT_TRIES) void adopt(api, key);

  if (!bag.doc(key).on) {
    if (watching !== null) {
      watcher?.stop();
      watching = null;
    }
    return;
  }

  const located = locate(api);
  const path = 'path' in located ? located.path : null;
  if (path === watching) return;

  watching = path;
  if (path) watcher?.start(path);
  else watcher?.stop();
}

/** the api the background session runs on. it starts as the stand-in and is
 *  upgraded the first time a command hands us cardmirror's own — which is what
 *  turns the console messages into toasts. */
let session: PluginApi = bootApi(ID);

function ensureSession(api: PluginApi): void {
  session = api;
  if (!watcher) watcher = watchSaves(() => void onSaved(session));
  if (!syncing) syncing = setInterval(() => sync(session), SYNC_MS);
  sync(session);
}

// ── commands ──────────────────────────────────────────────────────────

async function toggleLay(api: PluginApi): Promise<void> {
  const bag = store(api);
  const key = docKey();
  if (!key) {
    api.showToast('no document is open');
    return;
  }

  const on = !bag.doc(key).on;
  bag.setDoc(key, { on, templateId: templateIdFor(bag, key) });
  if (on) adopted.set(key, ADOPT_TRIES);
  sync(api);

  if (!on) {
    await withOpenDocx(api, clearMarker);
    last = null;
    api.showToast('lay formatting off');
    return;
  }

  // no template yet is the expected first step, not a failure worth a red line
  // in the panel — the menu has just opened out with the load button in it
  if (!bag.template(templateIdFor(bag, key))) {
    api.showToast('lay formatting on — load a template next');
    return;
  }

  // apply straight away rather than waiting for a save: turning it on and
  // seeing nothing change is indistinguishable from it not having worked
  await applyAndReport(api, 'lay formatting on — template applied');
}

async function loadTemplate(api: PluginApi): Promise<void> {
  const picked = await openFile(WORD_FILES);
  if (!picked) return;

  if (picked.bytes.length > TEMPLATE_LIMIT) {
    api.showToast(`${picked.name} is too large to keep as a template`);
    return;
  }

  const result = read(picked.bytes, picked.name);
  if (!result.ok) {
    api.showToast(result.error);
    return;
  }

  const bag = store(api);
  const id = `template:${picked.name}`;
  bag.addTemplate({ id, name: picked.name, path: picked.handle ?? null, docx: picked.bytes });

  // cardmirror's storage bag swallows a failed localStorage write, so a
  // template too big for the quota looks loaded until the next launch, when it
  // is simply gone. read it back rather than trust the write
  if (!bag.template(id)) {
    api.showToast(`${picked.name} is too large for cardmirror to keep — laymirror needs a smaller template`);
    return;
  }
  blueprints.set(id, result.blueprint);

  const key = docKey();
  if (key) bag.setDoc(key, { templateId: id });
  sync(api);

  const fields = result.blueprint.fields.length;
  const found = `${picked.name} — ${fields} header field${fields === 1 ? '' : 's'}`;

  // a template loaded onto a document that is already lay applies now. waiting
  // for the next save is what made loading a template look like it did nothing
  if (key && bag.doc(key).on) {
    await applyAndReport(api, `${found}, applied`);
    return;
  }
  api.showToast(`${found} — turn lay formatting on to apply it`);
}

function openLaymirror(api: PluginApi): void {
  ensureSession(api);
  if (panelOpen()) {
    closePanel();
    return;
  }

  const bag = () => store(api);
  const template = () => blueprintFor(bag(), templateIdFor(bag(), docKey()));

  openPanel({
    on: () => bag().doc(docKey()).on,
    templateName: () => bag().template(templateIdFor(bag(), docKey()))?.name ?? null,
    templatePath: () => bag().template(templateIdFor(bag(), docKey()))?.path ?? null,
    fields: () => template()?.fields ?? [],
    values: () => {
      const key = docKey();
      return bag().valuesFor(key, templateIdFor(bag(), key));
    },
    problem: () => {
      const located = locate(api);
      return 'error' in located ? located.error : null;
    },
    outcome: () => last,
    onToggle: () => toggleLay(api),
    onLoadTemplate: () => loadTemplate(api),
    // typing is saved as it happens, so ⌘S picks up what is on screen without
    // the user having to press apply first
    onChange: (values) => {
      const key = docKey();
      if (key) bag().setValues(key, templateIdFor(bag(), key), values);
    },
    onApply: async (values) => {
      const key = docKey();
      if (key) bag().setValues(key, templateIdFor(bag(), key), values);
      await applyAndReport(api, 'template applied');
    },
    actions: [{ label: 'diagnostics', run: () => openDiagnostics(api) }],
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
      // a plugin cannot put itself on the ribbon, so it had better arrive with
      // a key already bound. the alt chord is a fallback for a cardmirror that
      // has already taken the first.
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
      id: `${ID}.diagnose`,
      label: 'laymirror: diagnostics',
      keywords: ['debug', 'diagnose', 'why'],
      run: (api) => openDiagnostics(api),
    },
  ],
}) || console.warn('[laymirror] __registerCardMirrorPlugin unavailable');

// watch from the moment the script loads. cardmirror only hands a plugin its
// api inside a command's run(), so waiting for one would mean a document that
// was lay yesterday sat there doing nothing until the user opened the panel —
// and "press save and the school's document comes back" is the whole feature.
ensureSession(session);
