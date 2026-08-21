// jsdom has no layout engine, so geometry comes from data attributes and
// stubbed rects: block extents from `getBoundingClientRect`, line boxes from
// the range rects `measure` reads to decide where a block may be cut. it is
// the only way to drive the real measuring path.

import { vi } from 'vitest';

const LINE = 24;

const rect = (top: number, height: number): DOMRect =>
  ({
    top,
    height,
    bottom: top + height,
    left: 0,
    right: 624,
    width: 624,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect;

export function stubLayout(): void {
  vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (
    this: Element,
  ) {
    const el = this as HTMLElement;
    return rect(Number(el.dataset?.top ?? 0), Number(el.dataset?.h ?? 0));
  });

  // jsdom does not define this at all, so it is installed rather than spied
  Object.defineProperty(Range.prototype, 'getClientRects', {
    configurable: true,
    value: function (this: Range) {
      const node = this.commonAncestorContainer;
      const el = (node.nodeType === 1 ? node : node.parentElement) as HTMLElement | null;
      const height = Number(el?.dataset?.h ?? 0);
      const top = Number(el?.dataset?.top ?? 0);
      const lines = Math.max(1, Math.round(height / LINE));
      return Array.from({ length: lines }, (_, i) =>
        rect(top + i * LINE, LINE),
      ) as unknown as DOMRectList;
    },
  });
}

export function restoreLayout(): void {
  vi.restoreAllMocks();
  delete (Range.prototype as Partial<Range>).getClientRects;
}

/** an editor holding blocks of the given heights, laid end to end. */
export function makeEditor(heights: number[], text = 'text'): HTMLElement {
  const host = document.createElement('div');
  host.id = 'editor';
  const content = document.createElement('div');
  content.className = 'ProseMirror';
  host.append(content);

  let top = 0;
  for (const height of heights) {
    const block = document.createElement('p');
    block.className = 'pmd-card-body';
    block.textContent = text;
    block.dataset.h = String(height);
    block.dataset.top = String(top);
    content.append(block);
    top += height;
  }

  document.body.append(host);
  return content;
}

/** a canvas that measures a family the machine does not have: every generic
 *  fallback then reports its own width, which is exactly the signal the
 *  metric probe looks for. */
export function stubCanvas(absent: readonly string[]): void {
  const widths: Record<string, number> = { monospace: 100, serif: 110, 'sans-serif': 120 };
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => {
    const ctx = {
      font: '',
      measureText: () => {
        const [, family, fallback] = /72px "([^"]+)", (.+)/.exec(ctx.font) ?? [];
        const missing = absent.includes(family ?? '');
        return { width: missing ? (widths[fallback ?? ''] ?? 0) : 90 };
      },
    };
    return ctx as unknown as CanvasRenderingContext2D;
  });
}
