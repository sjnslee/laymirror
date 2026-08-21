// @vitest-environment jsdom
import { describe, it, expect, afterEach, beforeEach } from 'vitest';
import { showDraftMarks, clearDraftMarks, draftMarksShown } from '../src/render/draft-marks.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { makeEditor, restoreLayout, stubLayout } from './dom.js';

beforeEach(() => stubLayout());

afterEach(() => {
  clearDraftMarks();
  document.body.replaceChildren();
  document.head.replaceChildren();
  restoreLayout();
});

describe('draft marks', () => {
  it('have nothing to measure without an editor', () => {
    expect(showDraftMarks(DEFAULT_LAY)).toBeNull();
    expect(draftMarksShown()).toBe(false);
  });

  it('draw a rule for each page break', () => {
    makeEditor([400, 400, 400]);
    expect(showDraftMarks(DEFAULT_LAY)).toBe(1);
    expect(document.querySelectorAll('.lm-mark')).toHaveLength(1);
    expect(document.querySelector('.lm-mark span')?.textContent).toBe('page 2');
  });

  it('sit beside the editor content, never inside it', () => {
    // prosemirror owns its children: a node of ours in there is either
    // reconciled away or mistaken for document content
    const content = makeEditor([400, 400, 400]);
    showDraftMarks(DEFAULT_LAY);

    expect(content.querySelector('#laymirror-draft-marks')).toBeNull();
    expect(content.parentElement?.querySelector('#laymirror-draft-marks')).not.toBeNull();
  });

  it('leave nothing behind when cleared', () => {
    makeEditor([400, 400, 400]);
    showDraftMarks(DEFAULT_LAY);
    clearDraftMarks();

    expect(document.getElementById('laymirror-draft-marks')).toBeNull();
    expect(document.getElementById('laymirror-draft-style')).toBeNull();
  });
});

