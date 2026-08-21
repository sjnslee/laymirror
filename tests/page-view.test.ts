// @vitest-environment jsdom
import { describe, it, expect, afterEach, beforeEach } from 'vitest';
import { openPageView, closePageView, isPageViewOpen, PAGE_VIEW_ID } from '../src/render/page-view.js';
import { clearDraftMarks } from '../src/render/draft-marks.js';
import { printStyles } from '../src/render/print.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { PAGE_BREAK_TEXT } from '../src/profile/mapping.js';
import { makeEditor, restoreLayout, stubLayout } from './dom.js';

const meta = { title: '1AC', authors: 'A. Debater', teamCode: 'BCP 26-27' };

beforeEach(() => stubLayout());

afterEach(() => {
  closePageView();
  clearDraftMarks();
  document.body.replaceChildren();
  document.head.replaceChildren();
  restoreLayout();
});

describe('page view', () => {
  it('has nothing to show without an editor', () => {
    expect(openPageView(DEFAULT_LAY, meta)).toBeNull();
    expect(isPageViewOpen()).toBe(false);
  });

  it('lays a short document onto one page', () => {
    makeEditor([200, 200]);
    expect(openPageView(DEFAULT_LAY, meta)).toEqual({ pages: 1 });
    expect(document.querySelectorAll('.lm-page')).toHaveLength(1);
  });

  it('runs onto a second page when the column is full', () => {
    // the text column is 864px; three 400px blocks do not fit in one
    makeEditor([400, 400, 400]);
    expect(openPageView(DEFAULT_LAY, meta)).toEqual({ pages: 2 });
    expect(document.querySelectorAll('.lm-page')).toHaveLength(2);
  });

  it('numbers every page and heads it with the document metadata', () => {
    makeEditor([400, 400, 400]);
    openPageView(DEFAULT_LAY, meta);

    const feet = [...document.querySelectorAll('.lm-foot')].map((f) => f.textContent);
    expect(feet).toEqual(['Page 1 of 2', 'Page 2 of 2']);
    expect(document.querySelector('.lm-head')?.textContent).toContain('BCP 26-27');
    expect(document.querySelector('.lm-head')?.textContent).toContain('1AC');
  });

  it('clears its measuring stage away', () => {
    makeEditor([200]);
    openPageView(DEFAULT_LAY, meta);
    expect(document.querySelector('.lm-stage')).toBeNull();
  });

  it('never touches the editor it copied', () => {
    const content = makeEditor([200, 200]);
    const before = content.innerHTML;
    openPageView(DEFAULT_LAY, meta);

    expect(content.innerHTML).toBe(before);
    expect(content.querySelector(`#${PAGE_VIEW_ID}`)).toBeNull();
  });

  it('is inert — nothing in it is editable', () => {
    makeEditor([200]);
    openPageView(DEFAULT_LAY, meta);
    expect(document.getElementById(PAGE_VIEW_ID)?.getAttribute('contenteditable')).toBe('false');
    expect(document.querySelector(`#${PAGE_VIEW_ID} [contenteditable="true"]`)).toBeNull();
  });

  it('does not print the page-break marker itself', () => {
    makeEditor([200], PAGE_BREAK_TEXT);
    openPageView(DEFAULT_LAY, meta);
    expect(document.querySelector('.lm-body')?.textContent).toBe('');
  });

  it('closes on escape, and takes its stylesheet with it', () => {
    makeEditor([200]);
    openPageView(DEFAULT_LAY, meta);
    expect(document.getElementById('laymirror-page-style')).not.toBeNull();

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(isPageViewOpen()).toBe(false);
    expect(document.getElementById('laymirror-page-style')).toBeNull();
  });

  it('replaces itself rather than stacking when reopened', () => {
    makeEditor([200]);
    openPageView(DEFAULT_LAY, meta);
    openPageView(DEFAULT_LAY, meta);
    expect(document.querySelectorAll(`#${PAGE_VIEW_ID}`)).toHaveLength(1);
  });
});

describe('print styles', () => {
  it('size the sheet to the profile page and take the margin to zero', () => {
    const css = printStyles(DEFAULT_LAY.page, '.lm-page', PAGE_VIEW_ID);
    expect(css).toContain('size: 8.5000in 11.0000in');
    expect(css).toContain('margin: 0;');
    expect(css).toContain('break-after: page');
  });
});
