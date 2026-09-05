// laymirror's own status line.
//
// cardmirror's toast is positioned at the mouse pointer, so a message about a
// file write lands wherever the cursor happened to be sitting — over the text,
// off in a margin, or under the panel that asked for it. this one is always in
// the same corner, and there is only ever one of it.

const ROOT_ID = 'laymirror-status';
const STYLE_ID = 'laymirror-status-style';

/** long enough to read a filename in, short enough not to sit in the way */
const HOLD_MS = 4000;
const FADE_MS = 250;

const CSS = `
#${ROOT_ID} {
  position: fixed;
  right: 20px;
  bottom: 20px;
  z-index: 99998;
  max-width: 320px;
  box-sizing: border-box;
  padding: 7px 11px;
  border: 1px solid var(--pmd-c-border-soft, #d0d0d0);
  border-radius: 4px;
  background: var(--pmd-c-bg, #fff);
  color: var(--pmd-c-text, #222);
  font: 13px/1.45 var(--pmd-ui-font, system-ui, -apple-system, sans-serif);
  box-shadow: 0 8px 32px var(--pmd-c-shadow-deep, rgba(0, 0, 0, .25));
  transition: opacity ${FADE_MS}ms ease-out;
}
#${ROOT_ID}.lm-problem { color: var(--pmd-c-error, #b00020) }
#${ROOT_ID}.lm-going { opacity: 0 }
`;

let going: ReturnType<typeof setTimeout> | null = null;
let removing: ReturnType<typeof setTimeout> | null = null;

export function say(message: string, kind: 'ok' | 'problem' = 'ok'): void {
  if (typeof document === 'undefined') return;

  if (!document.getElementById(STYLE_ID)) {
    const sheet = document.createElement('style');
    sheet.id = STYLE_ID;
    sheet.textContent = CSS;
    document.head.append(sheet);
  }

  let root = document.getElementById(ROOT_ID);
  if (!root) {
    root = document.createElement('div');
    root.id = ROOT_ID;
    root.setAttribute('contenteditable', 'false');
    document.body.append(root);
  }

  // a second message replaces the first rather than stacking under it: they are
  // all about the same one file
  if (going) clearTimeout(going);
  if (removing) clearTimeout(removing);
  root.className = kind === 'problem' ? 'lm-problem' : '';
  root.textContent = message;

  const el = root;
  going = setTimeout(() => {
    el.classList.add('lm-going');
    removing = setTimeout(() => el.remove(), FADE_MS);
  }, HOLD_MS);
}
