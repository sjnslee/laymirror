// turning a place in the editor into a durable anchor.
//
// cardmirror puts each heading's stable uuid straight into the dom — pocket,
// hat, block, tag and analytic all render with `data-id` (see their `toDOM`
// in src/schema/nodes.ts) — and the same uuid reaches the file as a
// `pmd-heading-<uuid>` bookmark. so a page break can be pinned to a heading
// and an offset without touching a single prosemirror internal.

import type { PageBreak } from '../docx/breaks.js';

/** every block that becomes a `<w:p>`, in document order.
 *
 *  cardmirror's blocks are plain `p` / `h1`–`h4` elements, including the ones
 *  nested inside a card or a table cell, and the exporter walks the document
 *  in the same order — so this list lines up with the file's paragraphs. */
export function editorBlocks(editor: Element): HTMLElement[] {
  return Array.from(editor.querySelectorAll<HTMLElement>('p, h1, h2, h3, h4'));
}

/** the anchor for a block: the nearest heading at or before it, and how many
 *  blocks past that heading it sits. null when nothing before it carries an
 *  id, which happens only above the document's first heading. */
export function anchorFor(blocks: readonly HTMLElement[], target: Element): PageBreak | null {
  const index = blocks.indexOf(target as HTMLElement);
  if (index === -1) return null;

  for (let i = index; i >= 0; i--) {
    const id = blocks[i]!.getAttribute('data-id');
    if (id) return { headingId: id, offset: index - i };
  }
  return null;
}

/** the block a break resolves to now, for drawing it. */
export function blockForAnchor(
  blocks: readonly HTMLElement[],
  mark: PageBreak,
): HTMLElement | null {
  const anchor = blocks.findIndex((block) => block.getAttribute('data-id') === mark.headingId);
  if (anchor === -1) return null;
  return blocks[anchor + mark.offset] ?? null;
}

const sameAnchor = (a: PageBreak, b: PageBreak): boolean =>
  a.headingId === b.headingId && a.offset === b.offset;

/** add the break, or remove it when it is already there — one command that
 *  toggles, the way word's ctrl+enter reads as a single idea. */
export function toggleBreak(breaks: readonly PageBreak[], mark: PageBreak): PageBreak[] {
  const without = breaks.filter((existing) => !sameAnchor(existing, mark));
  return without.length === breaks.length ? [...breaks, mark] : without;
}

// clicking a button in the panel moves focus out of the editor, so by the
// time a command runs the selection is gone. the caret is therefore
// remembered as it moves, and commands use the remembered one.
let remembered: HTMLElement | null = null;
let onSelectionChange: (() => void) | null = null;

export function rememberCaret(editor: Element): void {
  forgetCaret();
  onSelectionChange = () => {
    const block = blockAtCaret(editor);
    if (block) remembered = block;
  };
  editor.ownerDocument.addEventListener('selectionchange', onSelectionChange);
  onSelectionChange();
}

export function forgetCaret(): void {
  if (onSelectionChange) {
    document.removeEventListener('selectionchange', onSelectionChange);
    onSelectionChange = null;
  }
  remembered = null;
}

/** the caret's block now, or the last one it was in. */
export function caretBlock(editor: Element): HTMLElement | null {
  const live = blockAtCaret(editor);
  if (live) return live;
  return remembered && editor.contains(remembered) ? remembered : null;
}

/** the block containing the caret. */
export function blockAtCaret(editor: Element): HTMLElement | null {
  const selection = editor.ownerDocument.defaultView?.getSelection();
  const node = selection?.focusNode;
  if (!node || !editor.contains(node)) return null;

  const start = (node.nodeType === 1 ? node : node.parentElement) as Element | null;
  const block = start?.closest('p, h1, h2, h3, h4') ?? null;
  return block && editor.contains(block) ? (block as HTMLElement) : null;
}
