// the panel. cardmirror has no command palette a plugin can lean on and no
// way for a plugin to place a ribbon button, so everything laymirror can do
// has to be reachable from here.

import { readTemplate } from '../profile/read-template.js';
import { validateMapping } from '../profile/mapping.js';
import { hasTemplate } from '../profile/defaults.js';
import type { Profile } from '../profile/profile.js';

const PANEL_ID = 'laymirror-panel';
const STYLE_ID = 'laymirror-panel-style';

export interface PanelAction {
  label: string;
  run(): void;
}

export interface PanelHooks {
  profile(): Profile;
  onProfile(profile: Profile): void | Promise<void>;
  isLay(): boolean;
  onToggleLay(): void | Promise<void>;
  breakCount(): number;
  actions: PanelAction[];
}

const CSS = `
#${PANEL_ID} {
  position: fixed;
  top: 64px;
  right: 24px;
  z-index: 99999;
  width: 340px;
  max-height: calc(100vh - 96px);
  overflow: auto;
  padding: 16px;
  border-radius: 10px;
  background: var(--pmd-c-panel, #23252a);
  color: var(--pmd-c-text, #e8e8ea);
  box-shadow: 0 12px 40px rgba(0,0,0,.45);
  font: 13px/1.5 system-ui, sans-serif;
}
#${PANEL_ID} h2 { margin: 0 0 12px; font-size: 14px; font-weight: 600 }
#${PANEL_ID} section { margin-bottom: 16px }
#${PANEL_ID} label { display: block; margin-bottom: 4px; opacity: .75 }
#${PANEL_ID} button {
  font: inherit;
  padding: 4px 10px;
  margin: 0 6px 6px 0;
  border-radius: 6px;
  cursor: pointer;
}
#${PANEL_ID} .lm-note { opacity: .7; margin-top: 6px }
#${PANEL_ID} .lm-warn { color: var(--pmd-c-warn, #e0a458); margin-top: 6px }
#${PANEL_ID} .lm-close { position: absolute; top: 10px; right: 12px }
`;

function ensureStyle(): void {
  if (document.getElementById(STYLE_ID)) return;
  const sheet = document.createElement('style');
  sheet.id = STYLE_ID;
  sheet.textContent = CSS;
  document.head.appendChild(sheet);
}

let onKey: ((e: KeyboardEvent) => void) | null = null;

export const isPanelOpen = (): boolean => document.getElementById(PANEL_ID) !== null;

export function closePanel(): void {
  if (onKey) {
    document.removeEventListener('keydown', onKey, true);
    onKey = null;
  }
  document.getElementById(PANEL_ID)?.remove();
}

const el = <K extends keyof HTMLElementTagNameMap>(
  tag: K,
  text?: string,
  className?: string,
): HTMLElementTagNameMap[K] => {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  if (className) node.className = className;
  return node;
};

export function openPanel(hooks: PanelHooks): void {
  closePanel();
  ensureStyle();

  const panel = el('div');
  panel.id = PANEL_ID;

  // the loaded-template notice lives here rather than in the dom, because a
  // successful load re-renders the panel and a fresh <input type=file> always
  // reads "no file chosen" — which is what made it look like nothing loaded
  let notice: { text: string; warn: boolean } | null = null;

  const render = (): void => {
    panel.replaceChildren();
    const profile = hooks.profile();

    const close = el('button', '✕', 'lm-close');
    close.addEventListener('click', closePanel);
    panel.append(close);

    panel.append(el('h2', 'laymirror'));

    // ── template ────────────────────────────────────────────────────
    const template = el('section');
    template.append(el('label', 'school template'));

    const input = el('input') as HTMLInputElement;
    input.type = 'file';
    input.accept = '.docx,.dotx,.dotm';
    input.addEventListener('change', async () => {
      const file = input.files?.[0];
      if (!file) return;
      const bytes = new Uint8Array(await file.arrayBuffer());
      const result = readTemplate(bytes, file.name);
      if (!result.ok) {
        notice = { text: result.error, warn: true };
        render();
        return;
      }
      await hooks.onProfile(result.profile);
      const warnings = validateMapping(result.profile);
      notice = {
        text: warnings.length
          ? `loaded ${file.name} — ${warnings[0]!.message}`
          : `loaded ${file.name}`,
        warn: warnings.length > 0,
      };
      render();
    });
    template.append(input);

    if (notice) {
      template.append(el('div', notice.text, notice.warn ? 'lm-warn' : 'lm-note'));
    } else if (hasTemplate(profile)) {
      template.append(el('div', `using ${profile.name}`, 'lm-note'));
    } else {
      template.append(
        el(
          'div',
          'no template yet — load your school’s .docx and laymirror will keep its header, styles and margins on every save',
          'lm-note',
        ),
      );
    }
    panel.append(template);

    // ── lay ─────────────────────────────────────────────────────────
    const lay = el('section');
    const toggle = el('button', hooks.isLay() ? 'turn lay off' : 'turn lay on');
    toggle.addEventListener('click', async () => {
      await hooks.onToggleLay();
      render();
    });
    lay.append(toggle);
    lay.append(
      el(
        'div',
        hooks.isLay()
          ? 'this document is a lay document'
          : 'this document is left exactly as cardmirror makes it',
        'lm-note',
      ),
    );
    panel.append(lay);

    // ── actions ─────────────────────────────────────────────────────
    const actions = el('section');
    for (const action of hooks.actions) {
      const button = el('button', action.label);
      button.addEventListener('click', () => {
        // a throwing action must not leave the panel stale and unresponsive —
        // that is indistinguishable from the button doing nothing
        try {
          action.run();
        } catch (err) {
          console.error(`[laymirror] ${action.label} failed`, err);
          notice = { text: `${action.label} failed — ${String(err)}`, warn: true };
        }
        render();
      });
      actions.append(button);
    }
    const breaks = hooks.breakCount();
    actions.append(
      el('div', `${breaks} manual page break${breaks === 1 ? '' : 's'}`, 'lm-note'),
    );
    panel.append(actions);

    // ── reaching it ─────────────────────────────────────────────────
    const help = el('section');
    help.append(
      el(
        'div',
        'shortcuts: ⌘⌥L opens this, ⌘⌥P page view, ⌘⌥↩ a page break. for a ribbon button, add one in settings → ribbon (up to 10).',
        'lm-note',
      ),
    );
    panel.append(help);
  };

  render();
  document.body.append(panel);

  onKey = (event: KeyboardEvent) => {
    if (event.key !== 'Escape') return;
    event.preventDefault();
    event.stopPropagation();
    closePanel();
  };
  // capture, because cardmirror binds escape too
  document.addEventListener('keydown', onKey, true);
}
