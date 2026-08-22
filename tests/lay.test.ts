// @vitest-environment jsdom
import { describe, expect, it, beforeEach, vi } from 'vitest';
import { adopt, enterLay, isLay, leaveLay, useWatcher, watchFile } from '../src/lay.js';
import { applyStylesheet, hasStylesheet, removeStylesheet, STYLE_ID } from '../src/render/css.js';

const watcher = () => ({ start: vi.fn(), stop: vi.fn(), resync: vi.fn(async () => {}) });

beforeEach(() => {
  document.head.replaceChildren();
  document.body.replaceChildren();
  leaveLay();
  useWatcher(null);
});

describe('entering and leaving', () => {
  it('starts off', () => {
    expect(isLay()).toBe(false);
    expect(hasStylesheet()).toBe(false);
  });

  it('turns on and off', async () => {
    await enterLay();
    expect(isLay()).toBe(true);
    leaveLay();
    expect(isLay()).toBe(false);
  });

  // dev-loading the plugin again reruns the module with `active` false while
  // the previous sheet is still in the head
  it('re-applies when the sheet is gone but state says on', async () => {
    await enterLay();
    removeStylesheet();
    await enterLay();
    expect(isLay()).toBe(true);
  });

  // this is the bug that made turning lay off never put the fonts back
  it('leaving removes the stylesheet even when state was already off', () => {
    applyStylesheet('.pmd-tag { color: red }');
    leaveLay();
    expect(document.getElementById(STYLE_ID)).toBeNull();
  });
});

describe('adopt', () => {
  it('turns lay on when the file says lay', async () => {
    await adopt('template:lay.docx', '/tmp/1ac.docx');
    expect(isLay()).toBe(true);
  });

  // the whole point of the rule: reading a file must never switch the screen
  // off. a failed read used to call leaveLay() and desynchronise everything
  it('never turns lay off, whatever the file says', async () => {
    await enterLay();
    await adopt(null, null);
    expect(isLay()).toBe(true);
  });

  it('leaves lay off when it was off and the file says nothing', async () => {
    await adopt(null, '/tmp/1ac.docx');
    expect(isLay()).toBe(false);
  });
});

describe('watching', () => {
  it('does not poll while lay is off', () => {
    const w = watcher();
    useWatcher(w);
    watchFile('/tmp/1ac.docx');
    expect(w.start).not.toHaveBeenCalled();
  });

  it('polls the file once lay is on', async () => {
    const w = watcher();
    useWatcher(w);
    await enterLay('/tmp/1ac.docx');
    expect(w.start).toHaveBeenCalledWith('/tmp/1ac.docx');
  });

  it('stops polling when lay is turned off', async () => {
    const w = watcher();
    useWatcher(w);
    await enterLay('/tmp/1ac.docx');
    leaveLay();
    expect(w.stop).toHaveBeenCalled();
  });

  it('does not restart the watcher for the same path', async () => {
    const w = watcher();
    useWatcher(w);
    await enterLay('/tmp/1ac.docx');
    watchFile('/tmp/1ac.docx');
    expect(w.start).toHaveBeenCalledTimes(1);
  });
});
