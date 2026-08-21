// laymirror — lay debate documents for cardmirror.
//
// off-state is genuinely inert: nothing is injected, watched or rewritten
// until a document is marked as lay.

import { register, type PluginApi } from './host/plugin-api.js';
import { hasFileApi, readFile, writeFile } from './host/electron.js';
import { resolveDocPath } from './host/paths.js';
import { zip, unzip, isDocx } from './docx/zip.js';
import { readMarker, writeMarker, clearMarker } from './docx/marker.js';

const ID = 'laymirror';
const DEFAULT_PROFILE = 'sample-lay';

/** resolve the open document and hand its parts to `edit`, then write back.
 *  every failure path reports rather than half-writing. */
async function withOpenDocx(
  api: PluginApi,
  edit: (parts: Record<string, Uint8Array>, path: string) => string,
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

  let parts;
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

  const message = edit(parts, found.path);
  await writeFile(found.path, zip(parts));
  api.showToast(message);
}

register({
  id: ID,
  name: 'laymirror',
  apiVersion: 1,
  commands: [
    {
      id: `${ID}.toggle-lay`,
      label: 'laymirror: mark this document as lay',
      keywords: ['lay', 'debate', 'parent', 'judge'],
      run: (api) =>
        withOpenDocx(api, (parts) => {
          const current = readMarker(parts);
          if (current) {
            clearMarker(parts);
            return 'lay formatting off for this document';
          }
          writeMarker(parts, DEFAULT_PROFILE);
          return `lay formatting on — profile "${DEFAULT_PROFILE}"`;
        }),
    },
    {
      id: `${ID}.status`,
      label: 'laymirror: status',
      keywords: ['lay', 'status'],
      run: async (api) => {
        const found = await resolveDocPath(api.docInfo());
        if (found.kind !== 'ok') {
          api.showToast(`no single document resolved (${found.kind})`);
          return;
        }
        const file = await readFile(found.path);
        const marker = file ? readMarker(unzip(file.bytes)) : null;
        api.showToast(marker ? `lay document — profile "${marker}"` : 'not a lay document');
      },
    },
  ],
}) || console.warn('[laymirror] __registerCardMirrorPlugin unavailable');
