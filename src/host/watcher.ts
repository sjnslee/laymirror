// save detection by polling: there is no save hook, so the mtime is the signal.
//
// a change counts only once it has been seen twice at the same size, so a save
// caught half-written is never read.

import { statFile, type FileStat } from './electron.js';

const FOCUSED_MS = 700;
/** a sync client can still write while the app is blurred, so back off rather
 *  than stop. */
const BLURRED_MS = 5000;

export interface Watcher {
  start(path: string): void;
  stop(): void;
  /** call after writing the file, or the write comes back as a save */
  resync(): Promise<void>;
}

const same = (a: FileStat | null, b: FileStat | null): boolean =>
  !!a && !!b && a.mtimeMs === b.mtimeMs && a.size === b.size;

export function watchSaves(onSaved: (path: string) => void): Watcher {
  let path: string | null = null;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let baseline: FileStat | null = null;
  let pending: FileStat | null = null;

  const delay = (): number =>
    typeof document !== 'undefined' && !document.hasFocus() ? BLURRED_MS : FOCUSED_MS;

  const schedule = (): void => {
    if (path !== null) timer = setTimeout(() => void tick(), delay());
  };

  async function tick(): Promise<void> {
    const watching = path;
    if (watching === null) return;

    let now: FileStat | null = null;
    try {
      now = await statFile(watching);
    } catch {
      // a stat that fails is not a save
    }

    // stop() may have run during the await
    if (path !== watching) return;

    if (now && !same(now, baseline)) {
      if (same(now, pending)) {
        baseline = now;
        pending = null;
        onSaved(watching);
      } else {
        pending = now;
      }
    } else {
      pending = null;
    }

    schedule();
  }

  function stop(): void {
    if (timer !== null) clearTimeout(timer);
    timer = null;
    path = null;
    baseline = null;
    pending = null;
  }

  function start(next: string): void {
    stop();
    path = next;
    // the state we started in is not a save
    void statFile(next)
      .catch(() => null)
      .then((stat) => {
        if (path !== next) return;
        baseline = stat;
        schedule();
      });
  }

  async function resync(): Promise<void> {
    if (path === null) return;
    const stat = await statFile(path).catch(() => null);
    baseline = stat;
    pending = null;
  }

  return { start, stop, resync };
}
