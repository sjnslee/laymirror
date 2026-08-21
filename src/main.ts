// laymirror — lay debate documents for cardmirror.
//
// off-state is genuinely inert: nothing is injected, watched or rewritten
// until a document is marked as lay.

import { register, type PluginApi } from './host/plugin-api.js';
import { hasFileApi, readFile, writeFile } from './host/electron.js';
import { resolveDocPath } from './host/paths.js';
import { watchSaves, type Watcher } from './host/watcher.js';
import { zip, unzip, isDocx, type Parts } from './docx/zip.js';
import { readMarker, writeMarker, clearMarker } from './docx/marker.js';
import { rewriteDocx } from './docx/rewrite.js';
import type { DocMeta } from './docx/headers.js';
import { syncTo, currentProfile, setProfile, activeProfile, useWatcher } from './lay.js';
import { openPanel } from './ui/settings-panel.js';
import type { Profile } from './profile/profile.js';

const ID = 'laymirror';
const PROFILE_KEY = 'profile';
const META_KEY = 'meta';

let watcher: Watcher | null = null;

/** resolve the open document and hand its parts to `edit`, then write back.
 *  every failure path reports rather than half-writing. */
async function withOpenDocx(
  api: PluginApi,
  edit: (parts: Parts) => string,
): Promise<string | null> {
  if (!hasFileApi()) {
    api.showToast('laymirror needs the desktop app');
    return null;
  }

  const found = await resolveDocPath(api.docInfo());
  if (found.kind === 'none') {
    api.showToast('no open .docx to work on — save the document first');
    return null;
  }
  if (found.kind === 'ambiguous') {
    api.showToast(
      `several open documents share this name; laymirror can't tell them apart ` +
        `(${found.paths.length} candidates)`,
    );
    return null;
  }

  const file = await readFile(found.path);
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

  const message = edit(parts);
  await writeFile(found.path, zip(parts));
  await watcher?.resync();
  api.showToast(message);
  return found.path;
}

/** the document's own path and lay marker, in one pass. */
async function readState(api: PluginApi): Promise<{ path: string | null; marker: string | null }> {
  const found = await resolveDocPath(api.docInfo());
  if (found.kind !== 'ok') return { path: null, marker: null };

  const file = await readFile(found.path);
  if (!file) return { path: found.path, marker: null };
  try {
    return { path: found.path, marker: readMarker(unzip(file.bytes)) };
  } catch {
    return { path: found.path, marker: null };
  }
}

function readMeta(api: PluginApi): DocMeta {
  const stored = api.storage.get(META_KEY);
  const saved = (stored && typeof stored === 'object' ? stored : {}) as Partial<DocMeta>;
  return {
    // the document names itself; the rest is the team's, and outlives it
    title: (api.docInfo()?.docTitle ?? saved.title ?? '').replace(/\.[^.]+$/, ''),
    authors: saved.authors ?? '',
    teamCode: saved.teamCode ?? '',
  };
}

/** the save pipeline: cardmirror has written its own docx over the file, so
 *  read it back and put the school's format on it. */
async function applyLay(api: PluginApi, path: string): Promise<void> {
  const file = await readFile(path);
  if (!file) return;

  let bytes: Uint8Array;
  try {
    bytes = rewriteDocx(file.bytes, currentProfile(), readMeta(api));
  } catch {
    api.showToast('laymirror could not read that save — the file is unchanged');
    return;
  }

  await writeFile(path, bytes);
  await watcher?.resync();
}

/** created once, started only by `enterLay`. */
function ensureWatcher(api: PluginApi): void {
  if (watcher) return;
  watcher = watchSaves((path) => void applyLay(api, path));
  useWatcher(watcher);
}

async function toggleLay(api: PluginApi): Promise<void> {
  let marker: string | null = null;
  const path = await withOpenDocx(api, (parts) => {
    if (readMarker(parts)) {
      clearMarker(parts);
      return 'lay formatting off for this document';
    }
    const profile = currentProfile();
    writeMarker(parts, profile.id);
    marker = profile.id;
    return `lay formatting on — ${profile.name}`;
  });

  if (path) syncTo(marker, path);
}

/** the profile outlives a restart; the marker only records which one. */
function restoreProfile(api: PluginApi): void {
  const stored = api.storage.get(PROFILE_KEY);
  if (stored && typeof stored === 'object') setProfile(stored as Profile);
}

register({
  id: ID,
  name: 'laymirror',
  apiVersion: 1,
  commands: [
    {
      // the one worth putting on the ribbon — everything else is in here
      id: `${ID}.panel`,
      label: 'laymirror: open',
      keywords: ['lay', 'debate', 'template', 'profile', 'settings'],
      run: async (api) => {
        restoreProfile(api);
        ensureWatcher(api);
        const { path, marker } = await readState(api);
        syncTo(marker, path);

        openPanel({
          profile: currentProfile,
          onProfile: (profile) => {
            setProfile(profile);
            api.storage.set(PROFILE_KEY, profile);
          },
          meta: () => readMeta(api),
          onMeta: (meta) => api.storage.set(META_KEY, meta),
          isLay: () => activeProfile() !== null,
          onToggleLay: () => toggleLay(api),
        });
      },
    },
    {
      id: `${ID}.toggle-lay`,
      label: 'laymirror: mark this document as lay',
      keywords: ['lay', 'debate', 'parent', 'judge'],
      run: async (api) => {
        restoreProfile(api);
        ensureWatcher(api);
        await toggleLay(api);
      },
    },
  ],
}) || console.warn('[laymirror] __registerCardMirrorPlugin unavailable');
