// what laymirror can actually see, printed.
//
// this exists because three rounds of this plugin were built against
// inference about the running app rather than observation of it, and every
// one of them was wrong somewhere. when a command does nothing, this says
// which lookup returned nothing.

import { currentFilename, DOC_NAME_CHIP, LS } from '../host/cardmirror.js';
import { resolveDocPath } from '../host/paths.js';
import type { PluginApi } from '../host/plugin-api.js';

const PANEL_ID = 'laymirror-diagnose';

interface Line {
  label: string;
  value: string;
}

const show = (value: unknown): string => {
  if (value === null) return 'null';
  if (value === undefined) return 'undefined';
  if (typeof value === 'string') return value === '' ? '(empty string)' : value;
  return JSON.stringify(value);
};

async function collect(api: PluginApi): Promise<Line[]> {
  const lines: Line[] = [];
  const add = (label: string, value: unknown) => lines.push({ label, value: show(value) });

  const electron = (window as unknown as { electronAPI?: Record<string, unknown> }).electronAPI;
  add('electronAPI', electron ? 'present' : 'MISSING — not the desktop app?');
  if (electron) {
    for (const method of ['readFileAtPath', 'writeFileAtPath', 'statFile', 'openFile']) {
      add(`  .${method}`, typeof electron[method] === 'function' ? 'ok' : 'MISSING');
    }
  }

  add('doc-name chip', document.getElementById(DOC_NAME_CHIP)?.textContent ?? null);
  add('document.title', document.title);
  add('currentFilename()', currentFilename());
  add('.ProseMirror', document.querySelector('.ProseMirror') ? 'found' : 'MISSING');
  add('#editor', document.getElementById('editor') ? 'found' : 'MISSING');

  let recents: unknown = null;
  try {
    recents = JSON.parse(localStorage.getItem(LS.recents) ?? 'null');
  } catch (err) {
    recents = `unreadable: ${String(err)}`;
  }
  if (Array.isArray(recents)) {
    add('pmd-recent-files', `${recents.length} entries`);
    for (const entry of recents.slice(0, 6)) {
      const e = entry as { filename?: string; format?: string; handle?: string | null };
      add(`  ${e.filename ?? '?'}`, `format=${e.format ?? '?'} handle=${e.handle ? 'yes' : 'NO'}`);
    }
  } else {
    add('pmd-recent-files', recents);
  }

  add('api.docInfo()', api.docInfo());
  const resolved = resolveDocPath(api.docInfo());
  add('resolveDocPath()', resolved);

  // does the renderer we depend on actually exist in this bundle?
  try {
    const docx = await import('docx-preview');
    add('docx-preview', typeof docx.renderAsync === 'function' ? 'loaded' : 'BROKEN');
  } catch (err) {
    add('docx-preview', `FAILED TO LOAD: ${String(err)}`);
  }

  return lines;
}

export async function openDiagnostics(api: PluginApi): Promise<void> {
  document.getElementById(PANEL_ID)?.remove();

  const panel = document.createElement('div');
  panel.id = PANEL_ID;
  panel.setAttribute(
    'style',
    'position:fixed;top:40px;left:50%;transform:translateX(-50%);z-index:100000;' +
      'width:min(680px,92vw);max-height:80vh;overflow:auto;padding:16px;border-radius:10px;' +
      'background:#16181d;color:#e8e8ea;box-shadow:0 12px 40px rgba(0,0,0,.6);' +
      'font:12px/1.5 ui-monospace,SFMono-Regular,Menlo,monospace;',
  );

  const title = document.createElement('div');
  title.textContent = 'laymirror diagnostics';
  title.setAttribute('style', 'font-weight:700;margin-bottom:10px');
  panel.append(title);

  const pre = document.createElement('pre');
  pre.setAttribute('style', 'white-space:pre-wrap;margin:0 0 12px');
  pre.textContent = 'collecting…';
  panel.append(pre);

  const buttons = document.createElement('div');
  const copy = document.createElement('button');
  copy.textContent = 'copy';
  copy.addEventListener('click', () => {
    void navigator.clipboard?.writeText(pre.textContent ?? '');
    copy.textContent = 'copied';
  });
  const close = document.createElement('button');
  close.textContent = 'close';
  close.addEventListener('click', () => panel.remove());
  buttons.append(copy, close);
  panel.append(buttons);

  document.body.append(panel);

  let text: string;
  try {
    const lines = await collect(api);
    const width = Math.max(...lines.map((l) => l.label.length));
    text = lines.map((l) => `${l.label.padEnd(width)}  ${l.value}`).join('\n');
  } catch (err) {
    text = `diagnostics threw: ${String(err)}\n${(err as Error)?.stack ?? ''}`;
  }
  pre.textContent = text;
  console.log('[laymirror] diagnostics\n' + text);
}
