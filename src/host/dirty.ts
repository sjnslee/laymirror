// has the document changed since the last save we saw?
//
// cardmirror tracks this as `currentDocDirty`, a module-local in
// src/editor/index.ts — not on window, not in the dom, not in the plugin api.
// so it cannot be read, and laymirror keeps its own answer instead: the
// editor's own mutations mark it dirty, and the save watcher marks it clean.
//
// page view depends on this. it renders the file on disk, so opening it over
// unsaved edits would quietly show the wrong document — and the alternative,
// saving on the user's behalf, would make a preview command mutate their work.

let dirty = false;
let observer: MutationObserver | null = null;

export const isDirty = (): boolean => dirty;

export function markClean(): void {
  dirty = false;
}

/** watch an editor for edits. idempotent, and safe to call again when the
 *  user opens a different document. */
export function watchEdits(editor: Element): void {
  stopWatchingEdits();
  observer = new MutationObserver(() => {
    dirty = true;
  });
  observer.observe(editor, {
    subtree: true,
    childList: true,
    characterData: true,
    attributes: false,
  });
}

export function stopWatchingEdits(): void {
  observer?.disconnect();
  observer = null;
}
