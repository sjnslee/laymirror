// @vitest-environment jsdom
//
// page view, actually rendered. jsdom has no layout, so this cannot judge how
// it looks — but it does prove the renderer runs, produces page sections, and
// puts the school's header on them, which is what "page view does nothing"
// would have caught.

import { describe, expect, it, beforeAll, beforeEach, afterEach } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { closePreview, isPreviewOpen, openPreview, PREVIEW_ID } from '../src/render/preview.js';

const DONOR = 'local/donor.docx';
const suite = existsSync(DONOR) ? describe : describe.skip;

beforeAll(() => {
  // docx-preview measures an svg to size tab leaders; jsdom implements no
  // svg geometry at all, so the one method it reaches for is installed
  (globalThis as unknown as { SVGElement: { prototype: Record<string, unknown> } }).SVGElement
    ?.prototype && ((SVGElement.prototype as unknown as Record<string, unknown>)['getBBox'] =
    () => ({ x: 0, y: 0, width: 0, height: 0 }));
  globalThis.requestAnimationFrame = ((cb: FrameRequestCallback) =>
    setTimeout(() => cb(Date.now()), 0) as unknown as number) as typeof requestAnimationFrame;
});

beforeEach(() => {
  document.body.replaceChildren();
  document.head.replaceChildren();
});
afterEach(() => closePreview());

suite('page view', () => {
  const bytes = () => new Uint8Array(readFileSync(DONOR));

  it('opens and puts pages on the screen', async () => {
    const result = await openPreview(bytes());
    expect(isPreviewOpen()).toBe(true);
    expect(result.pages).toBeGreaterThan(0);
    expect(document.querySelectorAll(`#${PREVIEW_ID} section.docx`).length).toBe(result.pages);
  });

  it('renders the school header onto the page', async () => {
    await openPreview(bytes());
    const header = document.querySelector(`#${PREVIEW_ID} header`);
    expect(header).not.toBeNull();
  });

  it('sizes the page from the template, not from a guess', async () => {
    await openPreview(bytes());
    const page = document.querySelector<HTMLElement>(`#${PREVIEW_ID} section.docx`)!;
    // 8.5in x 11in, and the school's 0.5in sides / 0.7in bottom
    expect(page.style.width).toBe('612pt');
    expect(page.style.minHeight).toBe('792pt');
    expect(page.style.paddingLeft).toBe('36pt');
    expect(page.style.paddingBottom).toBe('50.4pt');
  });

  it('says how the pages were decided', async () => {
    const result = await openPreview(bytes());
    expect(['word', 'document', 'estimated']).toContain(result.pagination);
    const bar = document.querySelector(`#${PREVIEW_ID} .lm-bar`)!;
    expect(bar.textContent).toContain('page');
  });

  it('offers no open-in-word button', async () => {
    await openPreview(bytes());
    const labels = [...document.querySelectorAll(`#${PREVIEW_ID} button`)].map((b) => b.textContent);
    expect(labels).not.toContain('open in word');
    expect(labels).toContain('close');
  });

  it('closes cleanly and leaves nothing behind', async () => {
    await openPreview(bytes());
    closePreview();
    expect(isPreviewOpen()).toBe(false);
    expect(document.getElementById('laymirror-page-style')).toBeNull();
  });

  it('reopening replaces rather than stacking', async () => {
    await openPreview(bytes());
    await openPreview(bytes());
    expect(document.querySelectorAll(`#${PREVIEW_ID}`)).toHaveLength(1);
  });
});
