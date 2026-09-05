// what laymirror can actually see, printed. when a command does nothing, this
// says which lookup returned nothing.

import { currentFilename, DOC_NAME_CHIP, LS, storageKey } from '../host/cardmirror.js';
import { resolveDocPath } from '../host/paths.js';
import type { PluginApi } from '../host/plugin-api.js';
import { store } from '../state.js';

const ROOT_ID = 'laymirror-diagnose';
const STYLE_ID = 'laymirror-diagnose-style';

const CSS = `
#${ROOT_ID} {
  position: fixed;
  inset: 0;
  z-index: 100000;
  display: flex;
  align-items: flex-start;
  justify-content: center;
  padding: 48px 16px;
  background: var(--pmd-c-overlay, rgba(0, 0, 0, .4));
}
#${ROOT_ID} .lm-dialog {
  display: flex;
  flex-direction: column;
  gap: 10px;
  width: min(680px, 100%);
  max-height: 100%;
  padding: 14px 16px;
  border: 1px solid var(--pmd-c-border-soft, #d0d0d0);
  border-radius: 4px;
  background: var(--pmd-c-bg, #fff);
  color: var(--pmd-c-text, #222);
  font: 13px/1.45 var(--pmd-ui-font, system-ui, -apple-system, sans-serif);
  box-shadow: 0 8px 32px var(--pmd-c-shadow-deep, rgba(0, 0, 0, .25));
}
#${ROOT_ID} h2 { margin: 0; font-weight: 600; font-size: 13px; line-height: 1.4 }
#${ROOT_ID} pre {
  flex: 1;
  overflow: auto;
  margin: 0;
  padding: 10px;
  border-radius: 3px;
  background: var(--pmd-c-surface-alt, #ececec);
  white-space: pre-wrap;
  font: 12px/1.5 ui-monospace, SFMono-Regular, Menlo, monospace;
}
#${ROOT_ID} .lm-actions { display: flex; gap: 8px; justify-content: flex-end }
#${ROOT_ID} button {
  padding: 4px 10px;
  border: 1px solid var(--pmd-c-border, #c8c8c8);
  border-radius: 4px;
  background: var(--pmd-c-bg, #fff);
  color: inherit;
  font: inherit;
  cursor: pointer;
}
#${ROOT_ID} button:hover { background: var(--pmd-c-button-hover, rgba(0, 0, 0, .06)) }
#${ROOT_ID} button:active {
  box-shadow: inset 0 0 0 99px var(--pmd-c-button-press, rgba(0, 0, 0, .15));
}
`;

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

function collect(api: PluginApi): Line[] {
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
  add('resolveDocPath()', resolveDocPath(api.docInfo()));
  // what stands in when the history has no entry to give
  add('path the user pointed at', store(api).doc(currentFilename()).path);

  // the bag laymirror reads before any command has run, and therefore the one
  // that decides whether a plain save is picked up at all
  const bag = localStorage.getItem(storageKey('laymirror'));
  add('plugin:laymirror', bag === null ? 'empty — nothing turned on yet' : `${bag.length} bytes`);

  return lines;
}

function report(api: PluginApi): string {
  try {
    const lines = collect(api);
    const width = Math.max(...lines.map((line) => line.label.length));
    return lines.map((line) => `${line.label.padEnd(width)}  ${line.value}`).join('\n');
  } catch (err) {
    return `diagnostics threw: ${String(err)}\n${(err as Error)?.stack ?? ''}`;
  }
}

function button(label: string, run: () => void): HTMLButtonElement {
  const el = document.createElement('button');
  el.type = 'button';
  el.textContent = label;
  el.addEventListener('click', run);
  return el;
}

export function openDiagnostics(api: PluginApi): void {
  document.getElementById(ROOT_ID)?.remove();

  if (!document.getElementById(STYLE_ID)) {
    const sheet = document.createElement('style');
    sheet.id = STYLE_ID;
    sheet.textContent = CSS;
    document.head.append(sheet);
  }

  const root = document.createElement('div');
  root.id = ROOT_ID;
  root.setAttribute('contenteditable', 'false');

  const close = (): void => {
    document.removeEventListener('keydown', onKey, true);
    root.remove();
  };
  // capture, because cardmirror binds escape too
  const onKey = (event: KeyboardEvent): void => {
    if (event.key !== 'Escape') return;
    event.preventDefault();
    event.stopPropagation();
    close();
  };
  document.addEventListener('keydown', onKey, true);
  root.addEventListener('click', (event) => {
    if (event.target === root) close();
  });

  const dialog = document.createElement('div');
  dialog.className = 'lm-dialog';

  const title = document.createElement('h2');
  title.textContent = 'laymirror diagnostics';

  const text = report(api);
  const pre = document.createElement('pre');
  pre.textContent = text;

  const copy = button('copy', () => {
    void navigator.clipboard?.writeText(text);
    copy.textContent = 'copied';
    setTimeout(() => void (copy.textContent = 'copy'), 1500);
  });

  const actions = document.createElement('div');
  actions.className = 'lm-actions';
  actions.append(copy, button('close', close));

  dialog.append(title, pre, actions);
  root.append(dialog);
  document.body.append(root);

  console.log('[laymirror] diagnostics\n' + text);
}
