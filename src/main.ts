// laymirror — lay debate documents for cardmirror.
//
// off-state is genuinely inert: nothing is injected, watched or rewritten
// until a document is turned on.

import { applyTemplate } from './docx/apply.js';
import { clearMarker, readMarker } from './docx/marker.js';
import { isDocx, unzip, zip, type Parts } from './docx/zip.js';
import { currentFilename, EDITOR_SELECTOR } from './host/cardmirror.js';
import {
  hasFileApi,
  openFile,
  readFile,
  writeFile,
  WORD_FILES,
} from './host/electron.js';
import { resolveDocPath } from './host/paths.js';
import { register, type PluginApi } from './host/plugin-api.js';
import { watchSaves, type Watcher } from './host/watcher.js';
import { clear as clearRules, draw as drawRules, shown as rulesShown } from './render/break-rules.js';
import { closePreview, isPreviewOpen, openPreview } from './render/preview.js';
import { store, TEMPLATE_LIMIT, type Store } from './state.js';
import { read, type Blueprint } from './template/template.js';
import { openDiagnostics } from './ui/diagnose.js';
import { closePanel, isOpen as panelOpen, openPanel } from './ui/panel.js';

const ID = 'laymirror';

/** how often laymirror notices the user has switched documents. cheap: a
 *  string compare against the filename chip, and nothing else unless it moved. */
const SYNC_MS = 1500;

let watcher: Watcher | null = null;
let watching: string | null = null;
let syncing: ReturnType<typeof setInterval> | null = null;
let drawn: string | null = null;
/** parsing a template is a few milliseconds of unzip; a save should not pay
 *  for it twice. */
const blueprints = new Map<string, Blueprint>();

/** documents are keyed by name rather than by path: a path can be missing
 *  while the document is plainly open, and losing the key would lose the
 *  template and the header values with it. */
const docKey = (): string | null => currentFilename();

const editor = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.ProseMirror') ??
  document.querySelector<HTMLElement>(EDITOR_SELECTOR);

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

// ── the save pipeline ─────────────────────────────────────────────────

/** cardmirror has rebuilt the file from scratch, so put the school's document
 *  back onto it. */
async function onSaved(api: PluginApi, path: string): Promise<void> {
  const bag = store(api);
  const key = docKey();
  if (!bag.doc(key).on) return;

  const templateId = templateIdFor(bag, key);
  const blueprint = blueprintFor(bag, templateId);
  if (!blueprint || !templateId) return;

  const file = await readFile(path);
  if (!file) return;

  let bytes: Uint8Array;
  try {
    bytes = applyTemplate(file.bytes, blueprint, bag.valuesFor(key, templateId), templateId);
  } catch {
    // a read that caught the file half-written is not a failure worth
    // shouting about — the next save lands on a whole file
    return;
  }

  await writeFile(path, bytes);
  await watcher?.resync();
}

/** read the open document, hand its parts to `edit`, write it back. */
async function withOpenDocx(
  api: PluginApi,
  edit: (parts: Parts) => void,
): Promise<string | null> {
  const located = locate(api);
  if ('error' in located) {
    api.showToast(located.error);
    return null;
  }

  const file = await readFile(located.path);
  if (!file) {
    api.showToast('could not read the document — reopen it and try again');
    return null;
  }

  let parts: Parts;
  try {
    parts = unzip(file.bytes);
  } catch {
    api.showToast('the document is not readable as a docx right now');
    return null;
  }
  // a partial read mid-save must abort, never round-trip into a write
  if (!isDocx(parts)) {
    api.showToast('the document looks incomplete — try again in a moment');
    return null;
  }

  edit(parts);
  await writeFile(located.path, zip(parts));
  await watcher?.resync();
  return located.path;
}

// ── keeping the screen and the watcher on the right document ──────────

/** document keys whose marker has already been read, so the file is not
 *  reopened on every tick. added before the read rather than after, because a
 *  tick can fire while the last one is still awaiting. */
const adopted = new Set<string>();

/** a document carrying laymirror's marker turns itself on. the marker travels
 *  inside the .docx, so a file a teammate marked arrives already lay — which is
 *  the whole reason it is in the file rather than in this machine's storage.
 *
 *  a document we could not reach is left unadopted so the next tick tries
 *  again: cardmirror fills its recent-files history a moment after the document
 *  appears, and giving up on that moment would leave the file plain. */
async function adopt(api: PluginApi, key: string): Promise<void> {
  adopted.add(key);
  const located = locate(api);
  if ('error' in located) {
    adopted.delete(key);
    return;
  }

  const file = await readFile(located.path);
  if (!file) {
    adopted.delete(key);
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
  const state = bag.doc(key);

  if (key && !adopted.has(key)) void adopt(api, key);

  if (!state.on) {
    if (watching !== null) {
      watcher?.stop();
      watching = null;
    }
    if (rulesShown()) clearRules();
    drawn = null;
    return;
  }

  const located = locate(api);
  const path = 'path' in located ? located.path : null;
  if (path !== watching) {
    watching = path;
    if (path) watcher?.start(path);
    else watcher?.stop();
  }

  // redrawing costs a stylesheet rebuild and a walk of the editor's text, so
  // it happens when the answer could have changed, not on every tick
  const templateId = templateIdFor(bag, key);
  const stamp = `${key}\u0000${templateId}`;
  if (stamp === drawn && rulesShown()) return;

  const host = editor();
  if (!host) return;
  drawRules(host, blueprintFor(bag, templateId)?.breaks ?? []);
  drawn = stamp;
}

function ensureSession(api: PluginApi): void {
  if (!watcher) watcher = watchSaves((path) => void onSaved(api, path));
  if (!syncing) syncing = setInterval(() => sync(api), SYNC_MS);
  sync(api);
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
  const templateId = templateIdFor(bag, key);
  bag.setDoc(key, { on, templateId });
  if (on) adopted.add(key);
  sync(api);

  if (!on) {
    await withOpenDocx(api, clearMarker);
    api.showToast('lay formatting off');
    return;
  }

  const name = bag.template(templateId)?.name;
  if (!name) {
    api.showToast('lay formatting on — load a template next');
    return;
  }

  // apply straight away rather than waiting for a save: turning it on and
  // seeing nothing change is indistinguishable from it not having worked
  if (await applyNow(api)) api.showToast(`lay formatting on — ${name}`);
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
  bag.addTemplate({ id, name: picked.name, docx: picked.bytes });
  blueprints.set(id, result.blueprint);

  const key = docKey();
  if (key) bag.setDoc(key, { templateId: id });

  const fields = result.blueprint.fields.length;
  api.showToast(
    `${picked.name} loaded — ${fields} header field${fields === 1 ? '' : 's'}`,
  );
  sync(api);
}

/** write the header values straight through, so the user sees the change
 *  without having to save first. */
async function applyNow(api: PluginApi): Promise<boolean> {
  const bag = store(api);
  const key = docKey();
  const templateId = templateIdFor(bag, key);
  const blueprint = blueprintFor(bag, templateId);
  if (!key || !templateId || !blueprint) {
    api.showToast('load a template first');
    return false;
  }

  const located = locate(api);
  if ('error' in located) {
    api.showToast(located.error);
    return false;
  }

  const file = await readFile(located.path);
  if (!file) {
    api.showToast('could not read the document — save it and try again');
    return false;
  }

  try {
    const bytes = applyTemplate(
      file.bytes,
      blueprint,
      bag.valuesFor(key, templateId),
      templateId,
    );
    await writeFile(located.path, bytes);
    await watcher?.resync();
    return true;
  } catch (err) {
    api.showToast(`could not apply — ${err instanceof Error ? err.message : String(err)}`);
    return false;
  }
}

async function showPageView(api: PluginApi): Promise<void> {
  if (isPreviewOpen()) {
    closePreview();
    return;
  }

  // page view renders a file, and working out which file is the one thing
  // cardmirror gives a plugin no reliable way to do. rather than refuse, fall
  // back to asking — a preview that always opens beats one that is right
  // about why it cannot.
  const located = locate(api);
  const file =
    'path' in located ? await readFile(located.path) : await openFile(WORD_FILES);
  if (!file) {
    api.showToast('error' in located ? located.error : 'no document chosen');
    return;
  }

  try {
    await openPreview(file.bytes, { label: file.name });
  } catch (err) {
    // say what actually went wrong: a silent failure here is indistinguishable
    // from the command not running at all
    api.showToast(`page view failed — ${err instanceof Error ? err.message : String(err)}`);
    console.error('[laymirror] page view failed', err);
  }
}

function toggleRules(api: PluginApi): void {
  if (rulesShown()) {
    clearRules();
    api.showToast('page break marks off');
    return;
  }
  const host = editor();
  if (!host) {
    api.showToast('no document is open');
    return;
  }
  const bag = store(api);
  const blueprint = blueprintFor(bag, templateIdFor(bag, docKey()));
  const { styled, literal } = drawRules(host, blueprint?.breaks ?? []);
  const total = styled + literal;
  api.showToast(`${total} page break${total === 1 ? '' : 's'}`);
}

const BREAK_NAMES: Record<string, string> = {
  pocket: 'pocket',
  hat: 'hat',
  block: 'block heading',
  tag: 'tag',
  analytic: 'analytic',
  undertag: 'undertag',
};

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
    breaks: () => {
      const types = (template()?.breaks ?? []).map((type) => BREAK_NAMES[type] ?? type);
      if (types.length === 0) return null;
      const last = types.pop()!;
      const named = types.length > 0 ? `${types.join(', ')} and ${last}` : last;
      return `starts a new page before every ${named}`;
    },
    fields: () => template()?.fields ?? [],
    values: () => {
      const key = docKey();
      return bag().valuesFor(key, templateIdFor(bag(), key));
    },
    problem: () => {
      const located = locate(api);
      return 'error' in located ? located.error : null;
    },
    onToggle: () => toggleLay(api),
    onLoadTemplate: () => loadTemplate(api),
    onApply: async (values) => {
      const key = docKey();
      if (!key) return;
      bag().setValues(key, templateIdFor(bag(), key), values);
      if (await applyNow(api)) api.showToast('header applied');
    },
    actions: [
      { label: 'page view', run: () => showPageView(api) },
      { label: 'page break marks', run: () => toggleRules(api) },
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
      // a plugin cannot put itself on the ribbon, so it had better arrive with
      // a key already bound. the alt chord is a fallback for a cardmirror that
      // has already taken the first.
      defaultKey: ['Mod-Shift-l', 'Mod-Alt-l'],
      run: (api) => openLaymirror(api),
    },
    {
      id: `${ID}.page-view`,
      label: 'laymirror: page view',
      keywords: ['page', 'print', 'pdf', 'layout', 'preview'],
      defaultKey: ['Mod-Shift-p', 'Mod-Alt-p'],
      run: (api) => showPageView(api),
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
      id: `${ID}.break-marks`,
      label: 'laymirror: page break marks',
      keywords: ['page', 'break', 'marks'],
      run: (api) => toggleRules(api),
    },
    {
      id: `${ID}.diagnose`,
      label: 'laymirror: diagnostics',
      keywords: ['debug', 'diagnose', 'why'],
      run: (api) => openDiagnostics(api),
    },
  ],
}) || console.warn('[laymirror] __registerCardMirrorPlugin unavailable');
