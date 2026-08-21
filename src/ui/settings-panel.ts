// laymirror's own modal. cardmirror's declared settings render only
// boolean/text/number/select, which cannot express a mapping table.
//
// the template picker is a plain file input on purpose: `readFileAtPath`
// serves only .cmir/.docx, so a school's .dotx could never come through it.
// an input element hands us the bytes directly, whatever the extension.

import { readTemplate } from '../profile/read-template.js';
import { validateMapping } from '../profile/mapping.js';
import { DEFAULT_LAY } from '../profile/defaults.js';
import type { BlockType, Profile, RunType } from '../profile/profile.js';
import { missingFonts } from './fonts.js';

const PANEL_ID = 'laymirror-panel';

const ROWS: { type: BlockType | RunType; label: string }[] = [
  { type: 'pocket', label: 'pocket' },
  { type: 'hat', label: 'hat' },
  { type: 'block', label: 'block' },
  { type: 'tag', label: 'tag' },
  { type: 'cite_paragraph', label: 'cite' },
  { type: 'card_body', label: 'card body' },
  { type: 'analytic', label: 'analytic' },
  { type: 'undertag', label: 'undertag' },
  { type: 'underline_mark', label: 'underline' },
  { type: 'cite_mark', label: 'cite mark' },
  { type: 'emphasis_mark', label: 'emphasis' },
];

export interface PanelHooks {
  profile(): Profile;
  onProfile(profile: Profile): void;
  isLay(): boolean;
  onToggleLay(): void | Promise<void>;
}

const el = <K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] => {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
};

function describe(profile: Profile, type: BlockType | RunType): string {
  const s = profile.types[type];
  const bits = [s.font ?? 'inherited', s.sizePt ? `${s.sizePt}pt` : null];
  if (s.bold) bits.push('bold');
  if (s.italic) bits.push('italic');
  if (s.smallCaps) bits.push('small caps');
  if (s.underline && s.underline !== 'none') bits.push(`${s.underline} underline`);
  if (s.pageBreakBefore) bits.push('page break before');
  return bits.filter(Boolean).join(' · ');
}

export function closePanel(): void {
  document.getElementById(PANEL_ID)?.remove();
}

export function openPanel(hooks: PanelHooks): void {
  closePanel();

  const overlay = el('div');
  overlay.id = PANEL_ID;
  Object.assign(overlay.style, {
    position: 'fixed',
    inset: '0',
    background: 'rgba(0,0,0,.45)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: '99999',
  });

  const dialog = el('div');
  Object.assign(dialog.style, {
    background: 'var(--pmd-c-surface, #fff)',
    color: 'var(--pmd-c-text, #111)',
    font: '13px/1.5 system-ui, sans-serif',
    padding: '20px 22px',
    borderRadius: '10px',
    width: 'min(680px, 92vw)',
    maxHeight: '86vh',
    overflow: 'auto',
    boxShadow: '0 12px 40px rgba(0,0,0,.35)',
  });

  const render = () => {
    dialog.replaceChildren();
    const profile = hooks.profile();

    const head = el('div');
    Object.assign(head.style, { display: 'flex', justifyContent: 'space-between', gap: '12px' });
    head.append(el('h2', undefined, 'laymirror'));
    const close = el('button', undefined, 'close');
    close.addEventListener('click', closePanel);
    head.append(close);
    dialog.append(head);

    dialog.append(el('p', undefined, `profile: ${profile.name}`));

    // state
    const state = el('p', undefined, hooks.isLay() ? 'this document is lay' : 'this document is not lay');
    dialog.append(state);
    const toggle = el('button', undefined, hooks.isLay() ? 'turn lay off' : 'turn lay on');
    toggle.addEventListener('click', async () => {
      await hooks.onToggleLay();
      render();
    });
    dialog.append(toggle);

    // template picker
    dialog.append(el('h3', undefined, 'template'));
    const picker = el('input');
    picker.type = 'file';
    picker.accept = '.docx,.dotx,.dotm,.docm';
    picker.addEventListener('change', async () => {
      const file = picker.files?.[0];
      if (!file) return;
      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const { profile: next, missing } = readTemplate(bytes, DEFAULT_LAY);
        hooks.onProfile({ ...next, name: file.name.replace(/\.[^.]+$/, '') });
        if (missing.length) {
          dialog.append(el('p', undefined, `no donor style for: ${missing.join(', ')}`));
        }
        render();
      } catch (err) {
        dialog.append(el('p', undefined, `could not read that template: ${String(err)}`));
      }
    });
    dialog.append(picker);

    // mapping
    dialog.append(el('h3', undefined, 'text types'));
    const table = el('table');
    table.style.width = '100%';
    table.style.borderCollapse = 'collapse';
    for (const row of ROWS) {
      const tr = el('tr');
      const name = el('td', undefined, row.label);
      const style = el('td', undefined, profile.types[row.type].styleName);
      const look = el('td', undefined, describe(profile, row.type));
      for (const cell of [name, style, look]) {
        cell.style.padding = '3px 6px';
        cell.style.borderBottom = '1px solid rgba(128,128,128,.25)';
      }
      style.style.opacity = '.75';
      look.style.opacity = '.75';
      tr.append(name, style, look);
      table.append(tr);
    }
    dialog.append(table);

    // round-trip warnings
    const warnings = validateMapping(profile);
    if (warnings.length) {
      dialog.append(el('h3', undefined, 'round-trip'));
      for (const w of warnings) {
        dialog.append(el('p', undefined, `${w.type}: ${w.message}`));
      }
    }

    // fonts
    const absent = missingFonts(profile);
    if (absent.length) {
      dialog.append(el('h3', undefined, 'fonts'));
      dialog.append(
        el(
          'p',
          undefined,
          `not installed: ${absent.join(', ')} — a substitute is used, so page breaks may drift from word`,
        ),
      );
    }
  };

  render();
  overlay.append(dialog);
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) closePanel();
  });
  document.body.append(overlay);
}
