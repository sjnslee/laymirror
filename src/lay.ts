// off / lay, and the profile in force.
//
// the rule this file exists to enforce: **screen state and file state are
// independent, and a failed file read may never change what is on screen.**
//
// the previous version derived lay from a file read. opening the panel called
// `syncTo(marker)`, a marker of null called `leaveLay()`, and any document
// whose path would not resolve was silently switched off — while its
// stylesheet stayed on the page. the toggle then read "off" and turned it
// back on, which is why turning lay off never put the fonts back.

import { applyStylesheet, hasStylesheet, removeStylesheet, toCss } from './render/css.js';
import { clearBreakMarks } from './render/break-marks.js';
import { closePreview } from './render/preview.js';
import { DEFAULT_PROFILE } from './profile/defaults.js';
import type { Profile } from './profile/profile.js';
import type { Watcher } from './host/watcher.js';

let active = false;
let profile: Profile = DEFAULT_PROFILE;
let watcher: Watcher | null = null;
let watching: string | null = null;

export const currentProfile = (): Profile => profile;
export const isLay = (): boolean => active;

/** swapping the profile restyles in place when lay is already on. */
export async function setProfile(next: Profile): Promise<void> {
  profile = next;
  if (active) applyStylesheet(await toCss(profile));
}

/** main hands the save pipeline in once. lay decides when it runs — the
 *  off-state polls nothing at all. */
export function useWatcher(next: Watcher | null): void {
  watcher?.stop();
  watcher = next;
  watching = null;
}

export function watchFile(path: string | null): void {
  if (path === watching) return;
  watching = path;
  if (path && active) watcher?.start(path);
  else watcher?.stop();
}

/** idempotent, so a poll or a re-registration can call it freely. */
export async function enterLay(path?: string | null): Promise<void> {
  // the dom is the truth, not `active`: dev-loading the plugin again reruns
  // this module from scratch while the previous sheet stays in the head
  if (!active || !hasStylesheet()) {
    applyStylesheet(await toCss(profile));
    active = true;
  }
  if (path !== undefined) watchFile(path);
}

/** only ever called from an explicit user toggle. unconditional, because
 *  `active` starts false after a reload while the sheet it describes may
 *  still be on the page, and turning lay off has to actually put the fonts
 *  back. */
export function leaveLay(): void {
  removeStylesheet();
  clearBreakMarks();
  closePreview();
  active = false;
  watching = null;
  watcher?.stop();
}

/** reconcile against what the file says. adoption only — a document whose
 *  marker says lay turns lay on; a marker we could not read, or a file we
 *  could not reach, changes nothing. */
export async function adopt(marker: string | null, path: string | null): Promise<void> {
  if (marker && !active) await enterLay(path);
  else if (active) watchFile(path);
}
