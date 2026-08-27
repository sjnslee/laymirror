// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { clear, draw, shown } from '../src/render/break-rules.js';

const rect = (over: Partial<DOMRect>): DOMRect =>
  ({
    top: 0,
    bottom: 0,
    height: 0,
    left: 0,
    right: 0,
    width: 0,
    x: 0,
    y: 0,
    toJSON: () => ({}),
    ...over,
  }) as DOMRect;

/** jsdom lays nothing out and its Range has no `getBoundingClientRect` at all,
 *  so both are supplied here. an all-zero rect is not a placeholder in these
 *  tests — it is exactly how the overlay reads "chromium skipped this block",
 *  the case it deliberately refuses to draw. */
function layOut(top: number | null): void {
  Range.prototype.getBoundingClientRect = () =>
    top === null ? rect({}) : rect({ top, bottom: top + 14, height: 14, y: top });
  Element.prototype.getBoundingClientRect = () =>
    rect({ bottom: 2000, height: 2000, left: 40, right: 640, width: 600, x: 40 });
}

function editor(html: string): HTMLElement {
  const el = document.createElement('div');
  el.className = 'ProseMirror';
  el.innerHTML = html;
  document.body.append(el);
  return el;
}

afterEach(() => {
  clear();
  vi.restoreAllMocks();
  document.body.replaceChildren();
});

describe('style-driven rules', () => {
  it('draws above the block types the template breaks before', () => {
    draw(editor('<h1 class="pmd-pocket">a</h1><h1 class="pmd-pocket">b</h1>'), ['pocket']);
    const sheet = document.getElementById('laymirror-break-rules')!;
    expect(sheet.textContent).toContain('.ProseMirror h1.pmd-pocket');
    expect(sheet.textContent).toContain('border-top');
  });

  // the first block of a document does not start a new page, it starts the
  // first one
  it('counts every one but the first', () => {
    const host = editor('<h1 class="pmd-pocket">a</h1><h1 class="pmd-pocket">b</h1>');
    expect(draw(host, ['pocket']).styled).toBe(1);
  });

  it('draws nothing for a template that breaks nowhere', () => {
    const host = editor('<h1 class="pmd-pocket">a</h1>');
    expect(draw(host, []).styled).toBe(0);
  });

  it('goes away again', () => {
    draw(editor('<h1 class="pmd-pocket">a</h1>'), ['pocket']);
    expect(shown()).toBe(true);
    clear();
    expect(shown()).toBe(false);
  });
});

describe('page-break characters', () => {
  // a form feed is what a word-derived pipeline leaves behind, and cardmirror
  // imports it as text and draws nothing at all
  it('draws a line across the column at a form feed', () => {
    layOut(300);
    const host = editor('<p class="pmd-card-body">before\fafter</p>');
    expect(draw(host, []).literal).toBe(1);

    const line = document.querySelector<HTMLElement>('#laymirror-break-overlay i')!;
    expect(line.style.top).toBe('307px');
    expect(line.style.left).toBe('40px');
    expect(line.style.width).toBe('600px');
  });

  it('finds every one, not just the first in a run of text', () => {
    layOut(300);
    expect(draw(editor('<p>a\fb\fc</p>'), []).literal).toBe(2);
  });

  // a block chromium never laid out reports nothing, and a line drawn at zero
  // would sit across the toolbar
  it('skips a block that was never laid out', () => {
    layOut(null);
    expect(draw(editor('<p>a\fb</p>'), []).literal).toBe(0);
  });

  it('leaves a document with no break characters alone', () => {
    layOut(300);
    expect(draw(editor('<p>ordinary prose</p>'), []).literal).toBe(0);
    expect(document.querySelectorAll('#laymirror-break-overlay i')).toHaveLength(0);
  });
});
