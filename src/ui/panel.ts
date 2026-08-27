// the laymirror menu.
//
// a plugin cannot put a button on cardmirror's ribbon and cannot add a
// settings page, so this is a floating panel over the editor. it is the only
// place laymirror has to say anything, which is why it says all of it: whether
// the document is lay, which template it is wearing, where that template
// breaks its pages, and what is in the header.

import type { Field, Values } from '../docx/fields.js';

const PANEL_ID = 'laymirror-panel';
const STYLE_ID = 'laymirror-panel-style';

export interface Action {
  label: string;
  run(): void | Promise<void>;
}

export interface PanelHost {
  on(): boolean;
  /** the loaded template's filename, or null when there is none. */
  templateName(): string | null;
  /** plain-english summary of where the template breaks pages. */
  breaks(): string | null;
  fields(): Field[];
  values(): Values;
  /** why laymirror cannot act right now, if it cannot. */
  problem(): string | null;
  onToggle(): void | Promise<void>;
  onLoadTemplate(): void | Promise<void>;
  onApply(values: Values): void | Promise<void>;
  actions: Action[];
}

const CSS = `
#${PANEL_ID} {
  position: fixed;
  top: 64px;
  right: 24px;
  z-index: 99997;
  width: 340px;
  max-height: calc(100vh - 96px);
  overflow: auto;
  padding: 14px 16px 16px;
  border-radius: 10px;
  background: #23252a;
  color: #e9eaec;
  font: 13px/1.45 system-ui, -apple-system, sans-serif;
  box-shadow: 0 10px 40px rgba(0,0,0,.5);
}
#${PANEL_ID} h2 {
  margin: 0 0 10px;
  font: 600 13px/1 system-ui, sans-serif;
  letter-spacing: .08em;
  text-transform: uppercase;
  opacity: .6;
}
#${PANEL_ID} .lm-close {
  position: absolute;
  top: 10px;
  right: 12px;
}
#${PANEL_ID} section { margin-top: 14px }
#${PANEL_ID} .lm-row {
  display: flex;
  gap: 8px;
  align-items: center;
  justify-content: space-between;
}
#${PANEL_ID} .lm-label { opacity: .6 }
#${PANEL_ID} .lm-note { margin-top: 4px; font-size: 12px; opacity: .55 }
#${PANEL_ID} .lm-problem { margin-top: 4px; font-size: 12px; color: #ffb4a2 }
#${PANEL_ID} label.lm-field {
  display: block;
  margin-top: 8px;
}
#${PANEL_ID} label.lm-field span {
  display: block;
  margin-bottom: 3px;
  font-size: 12px;
  opacity: .6;
}
#${PANEL_ID} input {
  width: 100%;
  box-sizing: border-box;
  padding: 5px 7px;
  border: 1px solid #454850;
  border-radius: 5px;
  background: #1a1c20;
  color: inherit;
  font: inherit;
}
#${PANEL_ID} input:focus { outline: 1px solid #7aa2f7; border-color: #7aa2f7 }
#${PANEL_ID} button {
  padding: 4px 11px;
  border: 1px solid #454850;
  border-radius: 5px;
  background: #2e3138;
  color: inherit;
  font: inherit;
  cursor: pointer;
}
#${PANEL_ID} button:hover { background: #3a3e46 }
#${PANEL_ID} button.lm-primary { background: #3b5bdb; border-color: #3b5bdb }
#${PANEL_ID} button.lm-primary:hover { background: #4c6ef5 }
#${PANEL_ID} button[disabled] { opacity: .45; cursor: default }
#${PANEL_ID} .lm-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-top: 16px;
  padding-top: 14px;
  border-top: 1px solid #34373e;
}
`;

let host: PanelHost | null = null;
let onKey: ((event: KeyboardEvent) => void) | null = null;

export const isOpen = (): boolean => document.getElementById(PANEL_ID) !== null;

export function closePanel(): void {
  if (onKey) {
    document.removeEventListener('keydown', onKey, true);
    onKey = null;
  }
  document.getElementById(PANEL_ID)?.remove();
  host = null;
}

function button(label: string, run: () => void, primary = false): HTMLButtonElement {
  const el = document.createElement('button');
  el.type = 'button';
  el.textContent = label;
  if (primary) el.className = 'lm-primary';
  el.addEventListener('click', run);
  return el;
}

function row(label: string, control: HTMLElement): HTMLDivElement {
  const el = document.createElement('div');
  el.className = 'lm-row';
  const name = document.createElement('span');
  name.className = 'lm-label';
  name.textContent = label;
  el.append(name, control);
  return el;
}

function note(text: string, kind: 'lm-note' | 'lm-problem' = 'lm-note'): HTMLParagraphElement {
  const el = document.createElement('p');
  el.className = kind;
  el.textContent = text;
  return el;
}

/** rebuild the body against whatever the host now reports. called after every
 *  action, so the panel never shows a state the plugin has moved on from. */
export function refresh(): void {
  const root = document.getElementById(PANEL_ID);
  if (!root || !host) return;

  const it = host;
  const body = document.createElement('div');

  const title = document.createElement('h2');
  title.textContent = 'laymirror';
  body.append(title, button('×', closePanel));
  (body.lastElementChild as HTMLElement).className = 'lm-close';

  const problem = it.problem();
  if (problem) body.append(note(problem, 'lm-problem'));

  const lay = document.createElement('section');
  lay.append(
    row(
      'lay formatting',
      button(it.on() ? 'on' : 'off', () => void act(() => it.onToggle()), it.on()),
    ),
  );
  body.append(lay);

  const template = document.createElement('section');
  const name = it.templateName();
  template.append(
    row('template', button(name ? 'change…' : 'load…', () => void act(() => it.onLoadTemplate()))),
    note(name ?? 'none — laymirror leaves the file alone until one is loaded'),
  );
  const breaks = it.breaks();
  if (breaks) template.append(note(breaks));
  body.append(template);

  const fields = it.fields();
  if (fields.length > 0) {
    const held = it.values();
    const inputs = new Map<string, HTMLInputElement>();

    const section = document.createElement('section');
    const heading = document.createElement('h2');
    heading.textContent = 'header';
    section.append(heading);

    for (const field of fields) {
      const label = document.createElement('label');
      label.className = 'lm-field';
      const caption = document.createElement('span');
      caption.textContent = field.label;
      const input = document.createElement('input');
      input.type = 'text';
      input.value = held[field.key] ?? field.label;
      input.placeholder = field.label;
      inputs.set(field.key, input);
      label.append(caption, input);
      section.append(label);
    }

    const apply = document.createElement('div');
    apply.className = 'lm-row';
    apply.style.marginTop = '12px';
    apply.style.justifyContent = 'flex-end';
    apply.append(
      button(
        'apply',
        () => {
          const values: Values = {};
          for (const [key, input] of inputs) values[key] = input.value;
          void act(() => it.onApply(values));
        },
        true,
      ),
    );
    section.append(apply);
    body.append(section);
  }

  const actions = document.createElement('div');
  actions.className = 'lm-actions';
  for (const action of it.actions) actions.append(button(action.label, () => void act(action.run)));
  body.append(actions);

  root.replaceChildren(...body.childNodes);
}

/** run an action, then show what it did. the panel is the only feedback
 *  surface laymirror has, so it must not go stale behind an await. */
async function act(run: () => void | Promise<void>): Promise<void> {
  await run();
  refresh();
}

export function openPanel(next: PanelHost): void {
  if (isOpen()) {
    host = next;
    refresh();
    return;
  }

  if (!document.getElementById(STYLE_ID)) {
    const sheet = document.createElement('style');
    sheet.id = STYLE_ID;
    sheet.textContent = CSS;
    document.head.append(sheet);
  }

  const root = document.createElement('div');
  root.id = PANEL_ID;
  root.setAttribute('contenteditable', 'false');
  document.body.append(root);

  host = next;
  refresh();

  onKey = (event: KeyboardEvent) => {
    if (event.key !== 'Escape') return;
    // typing in a field: escape should leave the field, not the panel
    if (document.activeElement?.tagName === 'INPUT') return;
    event.preventDefault();
    event.stopPropagation();
    closePanel();
  };
  // capture, because cardmirror binds escape too
  document.addEventListener('keydown', onKey, true);
}
