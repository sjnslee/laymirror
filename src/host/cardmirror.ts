// every undocumented cardmirror internal lives here, and nowhere else.
// each constant is stamped with the version it was verified against, so a
// cardmirror upgrade breaks a canary test instead of breaking the plugin
// mid-round.
//
// verified by reading the shipped build (/Applications/cardmirror.app,
// app.asar) and by the phase 0 spike run against a live editor.

export const VERIFIED_AGAINST = '1.3.0';

/** the editor container. `#editor` observed in the spike; the pane class
 *  is the multi-pane form. */
export const EDITOR_SELECTOR = '#editor, .pmd-pane-editor';

/** stable classes the stylesheet targets. counts observed in the spike on
 *  an imported lay template: tag 1, card-body 12, card 1. */
export const CLASS = {
  tag: 'pmd-tag',
  cardBody: 'pmd-card-body',
  citePara: 'pmd-cite-para',
  undertag: 'pmd-undertag',
  analytic: 'pmd-analytic',
  card: 'pmd-card',
} as const;

/** css custom properties cardmirror writes on :root and #editor. we set the
 *  ones that exist and write direct rules for what has no variable (per-type
 *  font-family). observed live: normal 11pt, tag 13pt, pocket 26pt,
 *  body-font '"Times New Roman", ...'. */
export const CSS_VAR = {
  sizeNormal: '--pmd-size-normal',
  sizePocket: '--pmd-size-pocket',
  sizeHat: '--pmd-size-hat',
  sizeBlock: '--pmd-size-block',
  sizeTag: '--pmd-size-tag',
  sizeAnalytic: '--pmd-size-analytic',
  sizeCite: '--pmd-size-cite',
  sizeUnderline: '--pmd-size-underline',
  sizeEmphasis: '--pmd-size-emphasis',
  sizeUndertag: '--pmd-size-undertag',
  bodyFont: '--pmd-body-font',
} as const;

export const LS = { recents: 'pmd-recent-files' } as const;

/** cardmirror paints the open document's filename here, and into
 *  `document.title`. these are the only signals that name the current
 *  document when it has no doc id — which is every word-authored .docx. */
export const DOC_NAME_CHIP = 'doc-name-chip-text';
export const TITLE_SUFFIX = ' — CardMirror';

/** `#editor` carries the base typography — font family, size, line height —
 *  and defines every `--pmd-*` variable the rest of cardmirror's stylesheet
 *  reads. `.pmd-pane-editor` defines none of it. so a clone taken out of the
 *  editor inherits nothing and renders at browser defaults, which is what
 *  made the first page view unreadable. copied explicitly instead. */
export function copyEditorStyle(target: HTMLElement): void {
  const editor = document.querySelector<HTMLElement>(EDITOR_SELECTOR);
  if (!editor) return;

  const computed = getComputedStyle(editor);
  for (let i = 0; i < computed.length; i++) {
    const property = computed.item(i);
    if (property.startsWith('--pmd-')) {
      target.style.setProperty(property, computed.getPropertyValue(property));
    }
  }
  for (const property of ['font-family', 'font-size', 'line-height', 'color']) {
    const value = computed.getPropertyValue(property);
    if (value) target.style.setProperty(property, value);
  }
  // .ProseMirror's, and the reason a run of spaces survives
  target.style.setProperty('white-space', 'pre-wrap');
}

/** our marker, stored beside cardmirror's own `cmirDocId`. */
export const MARKER_PROP = 'layMirrorProfile';
export const DOC_ID_PROP = 'cmirDocId';

/** a `pmd-recent-files` entry. `handle` is an absolute path on electron.
 *
 *  the list is capped at ten and the open document is unshifted to the front
 *  with a fresh `lastOpenedAt`, so it is a history — not a list of what is
 *  open now. treating it as the latter is what made every document look
 *  ambiguous. */
export interface RecentEntry {
  handle: string | null;
  filename: string;
  format: 'cmir' | 'docx' | string;
  lastOpenedAt?: number;
}

/** the filename cardmirror is showing for the document in front of the user. */
export function currentFilename(): string | null {
  const chip = document.getElementById(DOC_NAME_CHIP)?.textContent?.trim();
  if (chip) return chip;

  const title = document.title.endsWith(TITLE_SUFFIX)
    ? document.title.slice(0, -TITLE_SUFFIX.length).trim()
    : '';
  return title || null;
}

export function readRecents(): RecentEntry[] {
  try {
    const raw = JSON.parse(localStorage.getItem(LS.recents) ?? '[]');
    return Array.isArray(raw) ? (raw as RecentEntry[]) : [];
  } catch {
    return [];
  }
}

/** the open document as a prosemirror node, via the editor dom's
 *  `pmViewDesc` back-reference (prosemirror-view sets `dom.pmViewDesc = this`).
 *  undocumented but stable; verified in the spike — returned a `doc` node with
 *  children `block`, `paragraph`, `card`. null when the shape is not what we
 *  expect, never throws. */
export function readDocNode(): unknown | null {
  try {
    const dom = document.querySelector('.ProseMirror') as
      | (Element & { pmViewDesc?: { node?: unknown } })
      | null;
    return dom?.pmViewDesc?.node ?? null;
  } catch {
    return null;
  }
}
