// @vitest-environment jsdom
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { watchSaves } from '../src/host/watcher.js';

const stat = vi.fn();

beforeEach(() => {
  vi.useFakeTimers();
  stat.mockReset();
  (window as unknown as { electronAPI: unknown }).electronAPI = {
    statFile: stat,
    readFileAtPath: vi.fn(),
    writeFileAtPath: vi.fn(),
  };
  // jsdom reports no focus by default, which would put every test on the
  // backed-off interval
  vi.spyOn(document, 'hasFocus').mockReturnValue(true);
});

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

const at = (mtimeMs: number, size: number) => ({ mtimeMs, size });

/** let the poll run n times. */
const poll = (n = 1) => vi.advanceTimersByTimeAsync(750 * n);

describe('watchSaves', () => {
  it('says nothing about the state it started in', async () => {
    const saved = vi.fn();
    stat.mockResolvedValue(at(100, 10));
    watchSaves(saved).start('/doc.docx');

    await poll(3);
    expect(saved).not.toHaveBeenCalled();
  });

  it('reports a save once the file has stopped moving', async () => {
    const saved = vi.fn();
    stat.mockResolvedValue(at(100, 10));
    watchSaves(saved).start('/doc.docx');
    await poll();

    // word writes in stages: the first reading of a new mtime is not a
    // finished file, so nothing fires yet
    stat.mockResolvedValue(at(200, 40));
    await poll();
    expect(saved).not.toHaveBeenCalled();

    await poll();
    expect(saved).toHaveBeenCalledWith('/doc.docx');
    expect(saved).toHaveBeenCalledTimes(1);
  });

  it('does not report the same save twice', async () => {
    const saved = vi.fn();
    stat.mockResolvedValue(at(100, 10));
    watchSaves(saved).start('/doc.docx');
    await poll();

    stat.mockResolvedValue(at(200, 40));
    await poll(4);
    expect(saved).toHaveBeenCalledTimes(1);
  });

  it('absorbs our own write instead of chasing it', async () => {
    const saved = vi.fn();
    stat.mockResolvedValue(at(100, 10));
    const watcher = watchSaves(saved);
    watcher.start('/doc.docx');
    await poll();

    // the rewrite lands
    stat.mockResolvedValue(at(300, 90));
    await watcher.resync();

    await poll(4);
    expect(saved).not.toHaveBeenCalled();
  });

  it('stops polling when it is stopped', async () => {
    const saved = vi.fn();
    stat.mockResolvedValue(at(100, 10));
    const watcher = watchSaves(saved);
    watcher.start('/doc.docx');
    await poll();

    watcher.stop();
    const calls = stat.mock.calls.length;
    stat.mockResolvedValue(at(200, 40));
    await poll(5);

    expect(stat.mock.calls.length).toBe(calls);
    expect(saved).not.toHaveBeenCalled();
  });

  it('backs off while the window is unfocused', async () => {
    vi.spyOn(document, 'hasFocus').mockReturnValue(false);
    stat.mockResolvedValue(at(100, 10));
    watchSaves(vi.fn()).start('/doc.docx');
    await vi.advanceTimersByTimeAsync(0);

    const started = stat.mock.calls.length;
    await vi.advanceTimersByTimeAsync(2000);
    expect(stat.mock.calls.length).toBe(started);

    await vi.advanceTimersByTimeAsync(4000);
    expect(stat.mock.calls.length).toBeGreaterThan(started);
  });

  it('survives a stat that fails', async () => {
    const saved = vi.fn();
    stat.mockResolvedValue(at(100, 10));
    watchSaves(saved).start('/doc.docx');
    await poll();

    stat.mockRejectedValue(new Error('gone'));
    await poll(2);

    stat.mockResolvedValue(at(200, 40));
    await poll(2);
    expect(saved).toHaveBeenCalledWith('/doc.docx');
  });
});
