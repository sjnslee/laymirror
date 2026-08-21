// laymirror — lay debate documents for cardmirror.
//
// off-state is genuinely inert: nothing is injected, watched or rewritten
// until a document is marked as lay.

import { register, type PluginApi } from './host/plugin-api.js';
import { hasFileApi, readFile, writeFile } from './host/electron.js';
import { resolveDocPath } from './host/paths.js';
import { zip, unzip, isDocx, type Parts } from './docx/zip.js';
import { readMarker, writeMarker, clearMarker } from './docx/marker.js';
import { syncTo, currentProfile, setProfile, activeProfile } from './lay.js';
import { openPanel } from './ui/settings-panel.js';
import type { Profile } from './profile/profile.js';

const ID = 'laymirror';
const PROFILE_KEY = 'profile';

/** resolve the open document and hand its parts to `edit`, then write back.
 *  every failure path reports rather than half-writing. */
async function withOpenDocx(
  api: PluginApi,
  edit: (parts: Parts) => string,
): Promise<void> {
  if (!hasFileApi()) {
    api.showToast('laymirror needs the desktop app');
    return;
  }

  const found = await resolveDocPath(api.docInfo());
  if (found.kind === 'none') {
    api.showToast('no open .docx to work on — save the document first');
    return;
  }
  if (found.kind === 'ambiguous') {
    api.showToast(
      `several open documents share this name; laymirror can't tell them apart ` +
        `(${found.paths.length} candidates)`,
    );
    return;
  }

  const file = await readFile(found.path);
  if (!file) {
    api.showToast('could not read the document — reopen it and try again');
    return;
  }

  let parts: Parts;
  try {
    parts = unzip(file.bytes);
  } catch {
    api.showToast('the document is not readable as a docx right now');
    return;
  }
  // a partial read mid-save must abort, never round-trip into a write
  if (!isDocx(parts)) {
    api.showToast('the document looks incomplete — try again in a moment');
    return;
  }

  const message = edit(parts);
  await writeFile(found.path, zip(parts));
  api.showToast(message);
}

async function readMarkerOf(api: PluginApi): Promise<string | null> {
  const found = await resolveDocPath(api.docInfo());
  if (found.kind !== 'ok') return null;
  const file = await readFile(found.path);
  if (!file) return null;
  try {
    return readMarker(unzip(file.bytes));
  } catch {
    return null;
  }
}

async function toggleLay(api: PluginApi): Promise<void> {
  await withOpenDocx(api, (parts) => {
    if (readMarker(parts)) {
      clearMarker(parts);
      syncTo(null);
      return 'lay formatting off for this document';
    }
    const profile = currentProfile();
    writeMarker(parts, profile.id);
    syncTo(profile.id);
    return `lay formatting on — ${profile.name}`;
  });
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
        const marker = await readMarkerOf(api);
        syncTo(marker);
        openPanel({
          profile: currentProfile,
          onProfile: (profile) => {
            setProfile(profile);
            api.storage.set(PROFILE_KEY, profile);
          },
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
        await toggleLay(api);
      },
    },
  ],
}) || console.warn('[laymirror] __registerCardMirrorPlugin unavailable');
