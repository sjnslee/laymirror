// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import { openPanel, closePanel } from '../src/ui/settings-panel.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { zip } from '../src/docx/zip.js';
import { makeDocx, makeTemplate } from './fixture.js';

const meta = { title: '1AC', authors: '', teamCode: 'BCP 26-27' };

const hooks = (over: Partial<Parameters<typeof openPanel>[0]> = {}) => ({
  profile: () => DEFAULT_LAY,
  onProfile: vi.fn(),
  meta: () => meta,
  onMeta: vi.fn(),
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

  it('closes on escape', () => {
    openPanel(hooks());
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(document.querySelector('#laymirror-panel')).toBeNull();
  });

  it('ignores other keys', () => {
    openPanel(hooks());
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'e' }));
    expect(document.querySelector('#laymirror-panel')).not.toBeNull();
  });

  it('stops listening for escape once closed', () => {
    const detach = vi.spyOn(document, 'removeEventListener');
    openPanel(hooks());
    closePanel();
    expect(detach).toHaveBeenCalledWith('keydown', expect.any(Function), true);
    detach.mockRestore();
  });

  it('offers the header fields and reports them as typed', () => {
    const onMeta = vi.fn();
    openPanel(hooks({ onMeta }));
    const fields = [...document.querySelectorAll('input[type=text]')] as HTMLInputElement[];
    expect(fields).toHaveLength(3);
    expect(fields[0]!.value).toBe('BCP 26-27');

    fields[2]!.value = 'A. Debater';
    fields[2]!.dispatchEvent(new Event('input'));
    expect(onMeta).toHaveBeenCalledWith(expect.objectContaining({ authors: 'A. Debater' }));
    // the dialog must not be rebuilt under the caret
    expect(document.activeElement).not.toBe(document.body.firstElementChild);
    expect(document.querySelectorAll('input[type=text]')).toHaveLength(3);
  });

  it('lists the pocket last and says it is rare', () => {
    openPanel(hooks());
    const labels = [...document.querySelectorAll('tr')].map(
      (tr) => tr.firstElementChild?.textContent,
    );
    expect(labels[0]).toBe('hat');
    expect(labels.at(-1)).toBe('pocket (rare in lay)');
  });
});

// the picker is rebuilt on every render, so the browser's own "no file
// chosen" label always comes back — the panel has to say what it loaded.
describe('template picker', () => {
  const notice = () => document.querySelector('input[type=file] + p')?.textContent ?? '';

  const upload = async (bytes: Uint8Array, filename: string) => {
    const input = document.querySelector('input[type=file]') as HTMLInputElement;
    Object.defineProperty(input, 'files', {
      configurable: true,
      value: [new File([bytes as BlobPart], filename)],
    });
    input.dispatchEvent(new Event('change'));
  };

  it('confirms the load by name, and names the profile after it', async () => {
    const onProfile = vi.fn();
    openPanel(hooks({ onProfile }));
    await upload(makeTemplate(), 'westside lay.dotx');

    await vi.waitFor(() => expect(notice()).toContain('loaded westside lay.dotx'));
    expect(onProfile).toHaveBeenCalledWith(
      // the id rides in the marker, so two schools must not both be 'default'
      expect.objectContaining({ id: 'template:westside lay', name: 'westside lay' }),
    );
  });

  it('reports what a donor had no style for, but never the pocket', async () => {
    openPanel(hooks());
    await upload(zip(makeDocx()), 'bare.docx');

    await vi.waitFor(() => expect(notice()).toContain('no donor style for'));
    expect(notice()).toContain('tag');
    expect(notice()).not.toContain('pocket');
  });
});
