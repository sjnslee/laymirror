// the state machine over off / lay, plus the profile in force.
// off means off: no stylesheet, nothing injected, nothing watched.

import { applyStylesheet, hasStylesheet, removeStylesheet, toCss } from './render/css.js';
import { DEFAULT_LAY } from './profile/defaults.js';
import type { Profile } from './profile/profile.js';

let active: string | null = null;
let profile: Profile = DEFAULT_LAY;

export const currentProfile = (): Profile => profile;
export const activeProfile = (): string | null => active;

/** swapping the profile restyles in place when lay is already on. */
export function setProfile(next: Profile): void {
  profile = next;
  if (active !== null) {
    applyStylesheet(toCss(profile));
    active = next.id;
  }
}

/** idempotent, so a poll can call it freely. */
export function enterLay(profileId: string): void {
  // the dom is the truth, not `active`: dev-loading the plugin again reruns
  // this module from scratch while the previous sheet stays in the head
  if (active === profileId && hasStylesheet()) return;
  applyStylesheet(toCss(profile));
  active = profileId;
}

/** unconditional, for the same reason: `active` starts null after a reload
 *  but the sheet it describes may still be on the page, and turning lay off
 *  has to actually put the fonts back. */
export function leaveLay(): void {
  removeStylesheet();
  active = null;
}

/** marker present -> lay, absent -> off. */
export function syncTo(marker: string | null): void {
  if (marker) enterLay(marker);
  else leaveLay();
}
