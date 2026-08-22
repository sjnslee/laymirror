// laymirror — lay debate documents for cardmirror.
//
// off-state is genuinely inert: nothing is injected, watched or rewritten
// until a document is marked as lay.

import { register, type PluginApi } from './host/plugin-api.js';
import { hasFileApi, pickDocx, readFile, writeFile } from './host/electron.js';
import { resolveDocPath } from './host/paths.js';
import { currentFilename, EDITOR_SELECTOR } from './host/cardmirror.js';
import {
  anchorFor,
  caretBlock,
  editorBlocks,
  rememberCaret,
  toggleBreak,
} from './host/anchors.js';
import { openDiagnostics } from './ui/diagnose.js';
import { watchSaves, type Watcher } from './host/watcher.js';
import { zip, unzip, isDocx, type Parts } from './docx/zip.js';
import { readMarker, writeMarker, clearMarker } from './docx/marker.js';
import { applyProfile } from './docx/rewrite.js';
import type { PageBreak } from './docx/breaks.js';
import {
  adopt,
  currentProfile,
  enterLay,
  isLay,
  leaveLay,
  setProfile,
  useWatcher,
  watchFile,
} from './lay.js';
import { openPanel } from './ui/settings-panel.js';
import { closePreview, isPreviewOpen, openPreview } from './render/preview.js';
import { breakMarksShown, clearBreakMarks, showBreakMarks } from './render/break-marks.js';
import { DEFAULT_PROFILE } from './profile/defaults.js';
import type { Profile } from './profile/profile.js';
import { store, type DocState } from './store.js';

const ID = 'laymirror';

let watcher: Watcher | null = null;

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

/** documents are keyed by name rather than by path: a path can be missing
 *  while the document is plainly open, and losing the key would lose the
 *  profile and the page breaks with it. */
const docKey = (): string | null => currentFilename();

function editor(): HTMLElement | null {
  return (
    document.querySelector<HTMLElement>('.ProseMirror') ??
    document.querySelector<HTMLElement>(EDITOR_SELECTOR)
  );
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

function docState(api: PluginApi): DocState {
  const key = docKey();
  return store(api).doc(key);
}

function profileFor(api: PluginApi): Profile {
  const state = docState(api);
  const bag = store(api);
  // a document keeps its own profile; a document that has never had one
  // adopts whichever was used last, which is nearly always the right guess
  return bag.profile(state.profileId ?? bag.lastProfileId()) ?? DEFAULT_PROFILE;
}

/** the save pipeline: cardmirror has rebuilt the file from scratch, so put
 *  the school's document back onto it. */
async function onSaved(api: PluginApi, path: string): Promise<void> {
  const file = await readFile(path);
  if (!file) return;

  const bag = store(api);
  const key = docKey();
  const state = bag.doc(key);
  const profile = profileFor(api);

  let outcome;
  try {
    outcome = applyProfile(file.bytes, profile, state.breaks);
  } catch {
    api.showToast('laymirror could not read that save — the file is unchanged');
    return;
  }

  if (outcome.kind === 'adopted') {
    // word wrote this file, so its header is the truth — take it, and every
    // later cardmirror save restores what the user actually typed. a document
    // that already looks right is a template in its own right, so one marked
    // before any template was loaded teaches laymirror its own format.
    const adopted =
      profile.snapshot || !key
        ? { ...profile, snapshot: outcome.snapshot }
        : { ...profile, id: `document:${key}`, name: key, snapshot: outcome.snapshot };
    bag.setProfile(adopted);
    if (key) bag.setDoc(key, { profileId: adopted.id });
    return;
  }
  if (outcome.kind === 'skipped') return;

  await writeFile(path, outcome.bytes);
  await watcher?.resync();
}

function ensureWatcher(api: PluginApi): void {
  if (watcher) return;
  watcher = watchSaves((path) => void onSaved(api, path));
  useWatcher(watcher);
}

async function toggleLay(api: PluginApi): Promise<void> {
  const profile = profileFor(api);
  const wasLay = isLay();

  // the screen changes first, and whatever becomes of the file: a document
  // laymirror cannot reach must never be left wearing lay type after the user
  // has turned it off
  if (wasLay) leaveLay();
  else await enterLay();

  const path = await withOpenDocx(api, (parts) => {
    if (wasLay) clearMarker(parts);
    else writeMarker(parts, profile.id);
  });

  if (path === null) {
    api.showToast(wasLay ? 'lay off for this session only' : 'lay on for this session only');
    return;
  }

  watchFile(wasLay ? null : path);
  const key = docKey();
  if (!wasLay && key) store(api).setDoc(key, { profileId: profile.id });
  api.showToast(wasLay ? 'lay formatting off' : `lay formatting on — ${profile.name}`);
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
  const path =
    'path' in located ? located.path : await pickDocx('which document should laymirror lay out?');
  if (!path) {
    api.showToast('error' in located ? located.error : 'no document chosen');
    return;
  }

  const file = await readFile(path);
  if (!file) {
    api.showToast(`could not read ${path}`);
    return;
  }

  try {
    await openPreview(file.bytes);
  } catch (err) {
    // say what actually went wrong: a silent failure here is indistinguishable
    // from the command not running at all
    api.showToast(`page view failed — ${err instanceof Error ? err.message : String(err)}`);
    console.error('[laymirror] page view failed', err);
  }
}

function toggleBreakMarks(api: PluginApi): void {
  if (breakMarksShown()) {
    clearBreakMarks();
    api.showToast('page break marks off');
    return;
  }
  const host = editor();
  if (!host) {
    api.showToast('no document is open');
    return;
  }
  const drawn = showBreakMarks(host, docState(api).breaks);
  api.showToast(`${drawn} page break${drawn === 1 ? '' : 's'}`);
}

/** a page break is anchored to the nearest heading and held outside the
 *  document — cardmirror's model cannot carry one, and the text sentinel this
 *  replaces was visible, searchable and corruptible. */
function togglePageBreak(api: PluginApi): void {
  const host = editor();
  if (!host) {
    api.showToast('no document is open');
    return;
  }

  const block = caretBlock(host);
  if (!block) {
    api.showToast('put the cursor where the page should break');
    return;
  }

  const mark = anchorFor(editorBlocks(host), block);
  if (!mark) {
    api.showToast('a page break needs a pocket, hat, block or tag above it');
    return;
  }

  const key = docKey();
  if (!key) {
    api.showToast('no document is open');
    return;
  }

  const bag = store(api);
  const before = bag.doc(key).breaks;
  const after = toggleBreak(before, mark);
  bag.setDoc(key, { breaks: after });

  if (breakMarksShown()) showBreakMarks(host, after);
  api.showToast(
    after.length < before.length ? 'page break removed' : 'page break added — save to apply it',
  );
}

async function openLaymirror(api: PluginApi): Promise<void> {
  ensureWatcher(api);
  const host = editor();
  if (host) rememberCaret(host);

  const found = resolveDocPath(api.docInfo());
  const path = found.kind === 'ok' ? found.path : null;

  // reconciliation only. a marker we cannot read leaves lay exactly as it is —
  // reading the file must never be able to switch the screen off.
  let marker: string | null = null;
  if (path) {
    const file = await readFile(path);
    if (file) {
      try {
        marker = readMarker(unzip(file.bytes));
      } catch {
        marker = null;
      }
    }
  }
  await setProfile(profileFor(api));
  await adopt(marker, path);

  openPanel({
    profile: () => profileFor(api),
    onProfile: async (profile) => {
      const key = docKey();
      store(api).setProfile(profile);
      if (key) store(api).setDoc(key, { profileId: profile.id });
      await setProfile(profile);
    },
    isLay,
    onToggleLay: () => toggleLay(api),
    breakCount: () => docState(api).breaks.length,
    actions: [
      { label: 'page view', run: () => void showPageView(api) },
      { label: 'page break marks', run: () => toggleBreakMarks(api) },
      { label: 'insert page break', run: () => togglePageBreak(api) },
      { label: 'diagnostics', run: () => void openDiagnostics(api) },
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
      keywords: ['lay', 'debate', 'template', 'profile', 'settings'],
      // a plugin cannot put itself on the ribbon, so it had better arrive
      // with a key already bound
      defaultKey: 'Mod-Alt-l',
      run: (api) => openLaymirror(api),
    },
    {
      id: `${ID}.page-view`,
      label: 'laymirror: page view',
      keywords: ['page', 'print', 'layout', 'preview'],
      defaultKey: 'Mod-Alt-p',
      run: (api) => showPageView(api),
    },
    {
      id: `${ID}.page-break`,
      label: 'laymirror: insert page break',
      keywords: ['page', 'break'],
      defaultKey: 'Mod-Alt-Enter',
      run: (api) => togglePageBreak(api),
    },
    {
      id: `${ID}.diagnose`,
      label: 'laymirror: diagnostics',
      keywords: ['debug', 'diagnose', 'why'],
      run: (api) => openDiagnostics(api),
    },
    {
      id: `${ID}.toggle-lay`,
      label: 'laymirror: mark this document as lay',
      keywords: ['lay', 'debate', 'parent', 'judge'],
      run: async (api) => {
        ensureWatcher(api);
        await setProfile(profileFor(api));
        await toggleLay(api);
      },
    },
  ],
}) || console.warn('[laymirror] __registerCardMirrorPlugin unavailable');
