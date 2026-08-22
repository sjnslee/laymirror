// @vitest-environment jsdom
import { describe, expect, it, beforeEach, vi } from 'vitest';
import { closePanel, isPanelOpen, openPanel, type PanelHooks } from '../src/ui/settings-panel.js';
import { caretBlock, forgetCaret, rememberCaret } from '../src/host/anchors.js';
import { DEFAULT_PROFILE } from '../src/profile/defaults.js';

const hooks = (over: Partial<PanelHooks> = {}): PanelHooks => ({
  profile: () => DEFAULT_PROFILE,
  onProfile: vi.fn(),
  isLay: () => false,
  onToggleLay: vi.fn(),
  breakCount: () => 0,
  actions: [],
  ...over,
});

const buttons = () => [...document.querySelectorAll<HTMLButtonElement>('#laymirror-panel button')];
const buttonNamed = (label: string) => buttons().find((b) => b.textContent === label);

beforeEach(() => {
  document.body.replaceChildren();
  document.head.replaceChildren();
  closePanel();
  forgetCaret();
});

describe('the panel', () => {
  it('opens and closes', () => {
    openPanel(hooks());
    expect(isPanelOpen()).toBe(true);
    closePanel();
    expect(isPanelOpen()).toBe(false);
  });

  it('reopening replaces rather than stacking', () => {
    openPanel(hooks());
    openPanel(hooks());
    expect(document.querySelectorAll('#laymirror-panel')).toHaveLength(1);
  });

  it('closes on escape', () => {
    openPanel(hooks());
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(isPanelOpen()).toBe(false);
  });

  // "buttons don't do anything" — the run handler must actually fire
  it('runs an action when its button is clicked', () => {
    const run = vi.fn();
    openPanel(hooks({ actions: [{ label: 'page view', run }] }));
    buttonNamed('page view')!.click();
    expect(run).toHaveBeenCalledOnce();
  });

  it('survives an action that throws instead of wedging the panel', () => {
    const run = () => {
      throw new Error('boom');
    };
    openPanel(hooks({ actions: [{ label: 'page view', run }] }));
    expect(() => buttonNamed('page view')!.click()).not.toThrow();
    expect(isPanelOpen()).toBe(true);
  });

  it('toggles lay from its button', async () => {
    const onToggleLay = vi.fn();
    openPanel(hooks({ onToggleLay }));
    buttonNamed('turn lay on')!.click();
    expect(onToggleLay).toHaveBeenCalledOnce();
  });

  it('says the template is missing rather than pretending', () => {
    openPanel(hooks());
    expect(document.getElementById('laymirror-panel')!.textContent).toContain('no template yet');
  });

  it('offers no open-in-word button', () => {
    openPanel(hooks({ actions: [{ label: 'page view', run: vi.fn() }] }));
    expect(buttonNamed('open in word')).toBeUndefined();
  });
});

describe('the caret across a panel click', () => {
  function editor(): HTMLElement {
    document.body.innerHTML = `
      <div class="ProseMirror">
        <h2 class="pmd-hat" data-id="hat-1">a hat</h2>
        <p class="pmd-card-body">the body</p>
      </div>`;
    return document.querySelector<HTMLElement>('.ProseMirror')!;
  }

  // clicking a panel button moves focus out of the editor, so by the time the
  // command runs the selection is gone — which made "insert page break" do
  // nothing at all
  it('remembers the block after the selection leaves the editor', () => {
    const host = editor();
    rememberCaret(host);

    const body = host.querySelector('.pmd-card-body')!;
    const range = document.createRange();
    range.selectNodeContents(body);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    document.dispatchEvent(new Event('selectionchange'));

    // the panel takes focus and the selection is dropped
    selection.removeAllRanges();

    expect(caretBlock(host)?.textContent).toBe('the body');
  });

  it('has nothing to remember before the caret has been anywhere', () => {
    expect(caretBlock(editor())).toBeNull();
  });
});
