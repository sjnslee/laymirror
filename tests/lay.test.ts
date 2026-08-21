// @vitest-environment jsdom
import { describe, it, expect, beforeEach } from 'vitest';
import { syncTo, enterLay, leaveLay, activeProfile } from '../src/lay.js';
import { STYLE_ID } from '../src/render/css.js';

const styleEl = () => document.getElementById(STYLE_ID);

beforeEach(() => leaveLay());

describe('lay state', () => {
  it('injects nothing at all while off', () => {
    syncTo(null);
    expect(styleEl()).toBeNull();
    expect(activeProfile()).toBeNull();
    expect(document.head.querySelectorAll('style')).toHaveLength(0);
  });

  it('styles the editor when the document carries a marker', () => {
    syncTo('sample-lay');
    expect(activeProfile()).toBe('sample-lay');
    expect(styleEl()?.textContent).toContain('Palatino Linotype');
  });

  it('leaves nothing behind when the marker is cleared', () => {
    syncTo('sample-lay');
    syncTo(null);
    expect(styleEl()).toBeNull();
    expect(activeProfile()).toBeNull();
  });

  it('is idempotent — repeated syncs do not stack stylesheets', () => {
    syncTo('sample-lay');
    syncTo('sample-lay');
    enterLay('sample-lay');
    expect(document.querySelectorAll(`#${STYLE_ID}`)).toHaveLength(1);
  });

  it('swaps cleanly between profiles', () => {
    syncTo('sample-lay');
    syncTo('other');
    expect(activeProfile()).toBe('other');
    expect(document.querySelectorAll(`#${STYLE_ID}`)).toHaveLength(1);
  });
});
