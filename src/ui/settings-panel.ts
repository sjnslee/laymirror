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
import type { DocMeta } from '../docx/headers.js';
import { missingFonts, stackFor, substituteFont, SUBSTITUTES } from './fonts.js';

const PANEL_ID = 'laymirror-panel';

// lay documents are pocket-less in practice — the pocket is a tech-debate
// divider — so it is listed last and its absence from a donor is not worth
// reporting.
const RARE: readonly (BlockType | RunType)[] = ['pocket'];

const ROWS: { type: BlockType | RunType; label: string }[] = [
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
  { type: 'pocket', label: 'pocket (rare in lay)' },
];

export interface PanelAction {
  label: string;
  run(): void | Promise<void>;
}

export interface PanelHooks {
  profile(): Profile;
  onProfile(profile: Profile): void;
  meta(): DocMeta;
  onMeta(meta: DocMeta): void;
  isLay(): boolean;
  onToggleLay(): void | Promise<void>;
  /** page view, draft marks, print — reachable without a keyboard binding,
   *  since cardmirror has no command palette to find them in. */
  actions?: PanelAction[];
}

const META_FIELDS: { key: keyof DocMeta; label: string }[] = [
  { key: 'teamCode', label: 'team code' },
  { key: 'title', label: 'title' },
  { key: 'authors', label: 'authors' },
];

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

let onKey: ((e: KeyboardEvent) => void) | null = null;

export function closePanel(): void {
  if (onKey) {
    document.removeEventListener('keydown', onKey, true);
    onKey = null;
  }
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

  // a re-render rebuilds the file input, which then reads "no file chosen"
  // however well the load went — so what happened has to live in panel state
  // rather than on the element or in a node appended beside it
  let notice: string | null = null;
  let meta = hooks.meta();

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

    if (hooks.actions?.length) {
      const row = el('div');
      Object.assign(row.style, { display: 'flex', gap: '6px', flexWrap: 'wrap', marginTop: '8px' });
      for (const action of hooks.actions) {
        const button = el('button', undefined, action.label);
        button.addEventListener('click', () => void action.run());
        row.append(button);
      }
      dialog.append(row);
    }

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
        const name = file.name.replace(/\.[^.]+$/, '');
        // the id travels in the marker, so two schools' templates must not
        // both call themselves 'default'
        hooks.onProfile({ ...next, id: `template:${name}`, name });
        const absent = missing.filter((type) => !RARE.includes(type));
        notice = absent.length
          ? `loaded ${file.name} — no donor style for ${absent.join(', ')}`
          : `loaded ${file.name}`;
      } catch (err) {
        notice = `could not read ${file.name}: ${String(err)}`;
      }
      render();
    });
    dialog.append(picker);
    if (notice) dialog.append(el('p', undefined, notice));

    // header
    dialog.append(el('h3', undefined, 'header'));
    const donor = profile.headerXml !== null;
    dialog.append(
      el(
        'p',
        undefined,
        donor
          ? 'this template brings its own header — these fill {{team}}, {{title}} and {{authors}} in it'
          : 'no donor header, so laymirror builds one from these',
      ),
    );
    for (const field of META_FIELDS) {
      const row = el('label');
      Object.assign(row.style, { display: 'flex', gap: '8px', alignItems: 'center' });
      row.append(el('span', undefined, field.label));
      const input = el('input');
      input.type = 'text';
      input.value = meta[field.key];
      input.style.flex = '1';
      // no re-render on input: rebuilding the dialog would take the caret
      input.addEventListener('input', () => {
        meta = { ...meta, [field.key]: input.value };
        hooks.onMeta(meta);
      });
      row.append(input);
      dialog.append(row);
    }

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
          'not installed here, so something else is drawn and page breaks may ' +
            'drift from word. the printed file still asks for the real face.',
        ),
      );

      for (const family of absent) {
        const row = el('label');
        Object.assign(row.style, { display: 'flex', gap: '8px', alignItems: 'center' });
        row.append(el('span', undefined, family));

        const choice = document.createElement('select');
        choice.style.flex = '1';
        const current = stackFor(profile, family);
        for (const option of SUBSTITUTES) {
          const item = document.createElement('option');
          item.value = option.stack;
          item.textContent = option.label;
          item.selected = option.stack === current;
          choice.append(item);
        }
        // whatever the donor already asked for, if it is not on the list
        if (!SUBSTITUTES.some((option) => option.stack === current)) {
          const item = document.createElement('option');
          item.value = current;
          item.textContent = 'as the template asks';
          item.selected = true;
          choice.prepend(item);
        }

        choice.addEventListener('change', () => {
          hooks.onProfile(substituteFont(hooks.profile(), family, choice.value));
          render();
        });

        row.append(choice);
        dialog.append(row);
      }
    }
  };

  render();
  overlay.append(dialog);
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) closePanel();
  });
  document.body.append(overlay);

  // capture, because cardmirror binds escape too and the modal is on top
  onKey = (e: KeyboardEvent) => {
    if (e.key !== 'Escape') return;
    e.preventDefault();
    e.stopPropagation();
    closePanel();
  };
  document.addEventListener('keydown', onKey, true);
}
