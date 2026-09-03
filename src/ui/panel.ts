// the laymirror menu.
//
// a plugin cannot put a button on cardmirror's ribbon and cannot add a settings
// page, so this is a floating panel over the editor. none of laymirror's work
// shows up in the editor, which is why the panel also reports what the last
// write to the file actually did.

import type { Field, Values } from '../docx/fields.js';

const PANEL_ID = 'laymirror-panel';
const STYLE_ID = 'laymirror-panel-style';

export type Outcome = { ok: true; at: number; template: string } | { ok: false; why: string };

export interface Action {
  label: string;
  run(): void | Promise<void>;
}

export interface PanelHost {
  on(): boolean;
  /** the loaded template's filename, or null when there is none. */
  templateName(): string | null;
  /** where that file was picked from, when laymirror knows. */
  templatePath(): string | null;
  fields(): Field[];
  values(): Values;
  /** why laymirror cannot act right now, if it cannot. */
  problem(): string | null;
  /** what the last attempt to write the file did. */
  outcome(): Outcome | null;
  onToggle(): void | Promise<void>;
  onLoadTemplate(): void | Promise<void>;
  /** every keystroke in a header field, so a plain ⌘S picks up what is on
   *  screen without the user pressing apply first. */
  onChange(values: Values): void;
  onApply(values: Values): void | Promise<void>;
  actions: Action[];
}

// cardmirror's own tokens, so the panel follows its light and dark themes. the
// fallbacks are its light values, for a build that renames them.
const CSS = `
#${PANEL_ID} {
  position: fixed;
  top: 56px;
  right: 20px;
  z-index: 99997;
  width: 320px;
  max-height: calc(100vh - 88px);
  overflow: auto;
  padding: 12px 14px 14px;
  border: 1px solid var(--pmd-c-border-soft, #d0d0d0);
  border-radius: 4px;
  background: var(--pmd-c-bg, #fff);
  color: var(--pmd-c-text, #222);
  font: 13px/1.45 var(--pmd-ui-font, system-ui, -apple-system, sans-serif);
  box-shadow: 0 8px 32px var(--pmd-c-shadow-deep, rgba(0, 0, 0, .25));
}
#${PANEL_ID} h2 {
  margin: 0;
  font: 600 12px/1.4 inherit;
  text-transform: uppercase;
  letter-spacing: .06em;
  color: var(--pmd-c-text-muted, #666);
}
#${PANEL_ID} .lm-title {
  font-size: 13px;
  text-transform: none;
  letter-spacing: normal;
  color: var(--pmd-c-text, #222);
}
#${PANEL_ID} section { margin-top: 12px }
#${PANEL_ID} .lm-row {
  display: flex;
  gap: 8px;
  align-items: center;
  justify-content: space-between;
}
#${PANEL_ID} .lm-note {
  margin: 4px 0 0;
  font-size: 12px;
  color: var(--pmd-c-text-muted, #666);
}
#${PANEL_ID} .lm-path { word-break: break-all }
#${PANEL_ID} .lm-problem { color: var(--pmd-c-error, #b00020) }
#${PANEL_ID} .lm-done { color: var(--pmd-c-success, #16a34a) }
#${PANEL_ID} label.lm-field { display: block; margin-top: 8px }
#${PANEL_ID} label.lm-field span {
  display: block;
  margin-bottom: 3px;
  font-size: 12px;
  color: var(--pmd-c-text-muted, #666);
}
#${PANEL_ID} input {
  width: 100%;
  box-sizing: border-box;
  padding: 4px 6px;
  border: 1px solid var(--pmd-c-border, #c8c8c8);
  border-radius: 3px;
  background: var(--pmd-c-bg, #fff);
  color: inherit;
  font: inherit;
}
#${PANEL_ID} input:focus {
  outline: none;
  border-color: var(--pmd-c-focus, #4a90e2);
  box-shadow: 0 0 0 2px var(--pmd-c-focus-glow, rgba(74, 144, 226, .18));
}
#${PANEL_ID} button {
  flex: none;
  padding: 4px 10px;
  border: 1px solid var(--pmd-c-border, #c8c8c8);
  border-radius: 4px;
  background: var(--pmd-c-bg, #fff);
  color: inherit;
  font: inherit;
  cursor: pointer;
}
#${PANEL_ID} button:hover { background: var(--pmd-c-button-hover, rgba(0, 0, 0, .06)) }
#${PANEL_ID} button.lm-primary {
  background: var(--pmd-c-accent, #2563eb);
  border-color: var(--pmd-c-accent, #2563eb);
  color: var(--pmd-c-text-on-accent, #fff);
}
#${PANEL_ID} button.lm-primary:hover { background: var(--pmd-c-accent-hover, #1d4ed8) }
#${PANEL_ID} button.lm-close {
  padding: 0 6px;
  border-color: transparent;
  background: none;
  font-size: 15px;
  line-height: 1.4;
  color: var(--pmd-c-text-muted, #666);
}
#${PANEL_ID} button.lm-close:hover { color: var(--pmd-c-danger, #c00) }
#${PANEL_ID} button[disabled] { opacity: .5; cursor: default }
#${PANEL_ID} .lm-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-top: 14px;
  padding-top: 12px;
  border-top: 1px solid var(--pmd-c-divider, #e0e0e0);
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
  name.textContent = label;
  el.append(name, control);
  return el;
}

function note(
  text: string,
  kind: 'lm-note' | 'lm-problem' | 'lm-done' | 'lm-path' = 'lm-note',
): HTMLParagraphElement {
  const el = document.createElement('p');
  el.className = kind === 'lm-note' ? kind : `lm-note ${kind}`;
  el.textContent = text;
  return el;
}

/** rebuild the body against whatever the host now reports. called after every
 *  action, so the panel never shows a state the plugin has moved on from. */
export function refresh(): void {
  const root = document.getElementById(PANEL_ID);
  if (!root || !host) return;
  // a save can land while a header field is being typed into, and rebuilding
  // the panel under the caret would eat the word
  if (document.activeElement?.tagName === 'INPUT' && root.contains(document.activeElement)) {
    return;
  }

  const it = host;
  const body = document.createElement('div');

  const title = document.createElement('h2');
  title.className = 'lm-title';
  title.textContent = 'laymirror';
  const close = button('×', closePanel);
  close.className = 'lm-close';
  const head = document.createElement('div');
  head.className = 'lm-row';
  head.append(title, close);
  body.append(head);

  const lay = document.createElement('section');
  lay.append(
    row(
      it.on() ? 'lay formatting is on' : 'lay formatting is off',
      button(it.on() ? 'turn off' : 'turn on', () => void act(() => it.onToggle()), !it.on()),
    ),
  );
  body.append(lay);

  // off is off: a document laymirror is not touching has no template, no header
  // and nothing written to it, and offering all three invites the reasonable
  // assumption that something is happening
  if (!it.on()) {
    body.append(
      note('using cardmirror\u2019s own formatting. turn lay formatting on to apply a template every time you save.'),
      actionRow(it),
    );
    root.replaceChildren(...body.childNodes);
    return;
  }

  const problem = it.problem();
  if (problem) body.append(note(problem, 'lm-problem'));

  const template = document.createElement('section');
  const name = it.templateName();
  template.append(
    row('template', button(name ? 'change…' : 'load…', () => void act(() => it.onLoadTemplate()))),
    // the path, not the name: the name is the last segment of it
    note(name ? (it.templatePath() ?? name) : 'none loaded', name ? 'lm-path' : 'lm-note'),
  );
  body.append(template);

  const fields = it.fields();
  const held = it.values();
  const inputs = new Map<string, HTMLInputElement>();

  if (fields.length > 0) {
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
      // empty means "leave whatever the template says", shown greyed out as the
      // placeholder. putting the template's own words in as a *value* meant
      // applying wrote them straight back, so an edit made to the template was
      // overwritten by the text laymirror had been showing
      input.value = held[field.key] ?? '';
      input.placeholder = field.label;
      // held as typed rather than on apply, so a plain ⌘S writes what is on
      // screen. the panel is deliberately not refreshed here: rebuilding it
      // mid-word would take the caret with it
      input.addEventListener('input', () => it.onChange(typed(inputs)));
      inputs.set(field.key, input);
      label.append(caption, input);
      section.append(label);
    }
    body.append(section);
  }

  const done = document.createElement('section');
  done.append(
    row(
      'the file on disk',
      button('apply now', () => void act(() => it.onApply(typed(inputs))), true),
    ),
  );
  const outcome = it.outcome();
  done.append(
    outcome === null
      ? note('nothing written yet')
      : outcome.ok
        ? note(`${outcome.template} applied at ${clock(outcome.at)}`, 'lm-done')
        : note(outcome.why, 'lm-problem'),
  );
  body.append(done);

  body.append(actionRow(it));
  root.replaceChildren(...body.childNodes);
}

function actionRow(it: PanelHost): HTMLDivElement {
  const actions = document.createElement('div');
  actions.className = 'lm-actions';
  for (const action of it.actions) actions.append(button(action.label, () => void act(action.run)));
  return actions;
}

/** run an action, then show what it did. the panel is the only feedback
 *  surface laymirror has, so it must not go stale behind an await. */
/** only the boxes with something in them. an empty box is not a blank header
 *  line, it is "the template's own text is fine". */
const typed = (inputs: ReadonlyMap<string, HTMLInputElement>): Values => {
  const values: Values = {};
  for (const [key, input] of inputs) {
    if (input.value !== '') values[key] = input.value;
  }
  return values;
};

/** minutes matter, seconds do not: this answers "did that save go through?" */
const clock = (at: number): string =>
  new Date(at).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });

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
