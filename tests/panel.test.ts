// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import { openPanel, closePanel } from '../src/ui/settings-panel.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';

const hooks = (over: Partial<Parameters<typeof openPanel>[0]> = {}) => ({
  profile: () => DEFAULT_LAY,
  onProfile: vi.fn(),
  isLay: () => false,
  onToggleLay: vi.fn(),
  ...over,
});

afterEach(() => closePanel());

describe('settings panel', () => {
  it('lists every text type with its donor style', () => {
    openPanel(hooks());
    const text = document.body.textContent ?? '';
    for (const label of ['pocket', 'tag', 'cite', 'card body', 'underline']) {
      expect(text).toContain(label);
    }
    expect(text).toContain('Times New Roman');
  });

  it('takes a template through a file input, not the scoped host api', () => {
    openPanel(hooks());
    const input = document.querySelector('input[type=file]') as HTMLInputElement;
    expect(input).not.toBeNull();
    // .dotx must be offered — readFileAtPath would refuse it
    expect(input.accept).toContain('.dotx');
  });

  it('reflects lay state and offers the matching action', () => {
    openPanel(hooks({ isLay: () => true }));
    const text = document.body.textContent ?? '';
    expect(text).toContain('this document is lay');
    expect(text).toContain('turn lay off');
  });

  it('surfaces the round-trip warning about cite marks', () => {
    openPanel(hooks());
    expect(document.body.textContent).toContain('native path');
  });

  it('opens once even if asked twice, and closes cleanly', () => {
    openPanel(hooks());
    openPanel(hooks());
    expect(document.querySelectorAll('#laymirror-panel')).toHaveLength(1);
    closePanel();
    expect(document.querySelector('#laymirror-panel')).toBeNull();
  });
});
