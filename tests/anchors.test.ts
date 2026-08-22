// @vitest-environment jsdom
import { describe, expect, it, beforeEach } from 'vitest';
import {
  anchorFor,
  blockAtCaret,
  blockForAnchor,
  editorBlocks,
  toggleBreak,
} from '../src/host/anchors.js';

/** the editor as cardmirror renders it: headings carry `data-id`, a card is a
 *  div wrapping its own paragraphs. */
function editor(): HTMLElement {
  document.body.innerHTML = `
    <div class="ProseMirror">
      <p>intro</p>
      <h2 class="pmd-hat" data-id="hat-1">first hat</h2>
      <div class="pmd-card">
        <h4 class="pmd-tag" data-id="tag-1">the tag</h4>
        <p class="pmd-cite-para">a cite</p>
        <p class="pmd-card-body">the body</p>
      </div>
      <h2 class="pmd-hat" data-id="hat-2">second hat</h2>
      <p>after</p>
    </div>`;
  return document.querySelector<HTMLElement>('.ProseMirror')!;
}

let host: HTMLElement;
beforeEach(() => {
  host = editor();
});

describe('editorBlocks', () => {
  // a card is a div wrapping paragraphs, so its children must be flattened in
  // place or the order stops matching the file's paragraphs
  it('flattens cards and keeps document order', () => {
    expect(editorBlocks(host).map((b) => b.textContent)).toEqual([
      'intro',
      'first hat',
      'the tag',
      'a cite',
      'the body',
      'second hat',
      'after',
    ]);
  });
});

describe('anchorFor', () => {
  const blocks = () => editorBlocks(host);

  it('anchors a heading to itself at offset 0', () => {
    const hat = host.querySelector('[data-id="hat-1"]')!;
    expect(anchorFor(blocks(), hat)).toEqual({ headingId: 'hat-1', offset: 0 });
  });

  it('counts forward from the nearest heading above', () => {
    const body = host.querySelector('.pmd-card-body')!;
    expect(anchorFor(blocks(), body)).toEqual({ headingId: 'tag-1', offset: 2 });
  });

  // nothing above the first heading carries an id, so there is no durable
  // anchor and saying so beats inventing one
  it('has no anchor above the first heading', () => {
    const intro = blocks()[0]!;
    expect(anchorFor(blocks(), intro)).toBeNull();
  });

  it('is null for an element outside the editor', () => {
    expect(anchorFor(blocks(), document.createElement('p'))).toBeNull();
  });
});

describe('blockForAnchor', () => {
  it('finds the block a mark resolves to', () => {
    const block = blockForAnchor(editorBlocks(host), { headingId: 'tag-1', offset: 1 });
    expect(block?.textContent).toBe('a cite');
  });

  it('is null when the anchor has gone', () => {
    expect(blockForAnchor(editorBlocks(host), { headingId: 'gone', offset: 0 })).toBeNull();
  });

  it('is null when the offset runs off the end', () => {
    expect(blockForAnchor(editorBlocks(host), { headingId: 'hat-2', offset: 50 })).toBeNull();
  });
});

describe('toggleBreak', () => {
  const mark = { headingId: 'hat-1', offset: 0 };

  it('adds a break that is not there', () => {
    expect(toggleBreak([], mark)).toEqual([mark]);
  });

  it('removes one that is', () => {
    expect(toggleBreak([mark], mark)).toEqual([]);
  });

  it('leaves other breaks alone', () => {
    const other = { headingId: 'hat-2', offset: 1 };
    expect(toggleBreak([other, mark], mark)).toEqual([other]);
  });
});

describe('blockAtCaret', () => {
  it('finds the block holding the selection', () => {
    const body = host.querySelector('.pmd-card-body')!;
    const range = document.createRange();
    range.selectNodeContents(body);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    expect(blockAtCaret(host)?.textContent).toBe('the body');
  });

  it('is null when the selection is outside the editor', () => {
    window.getSelection()?.removeAllRanges();
    expect(blockAtCaret(host)).toBeNull();
  });
});
