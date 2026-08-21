// the state machine over off / lay. off means off: no stylesheet, nothing
// injected, nothing watched.

import { applyStylesheet, removeStylesheet, toCss } from './render/css.js';
import { DEFAULT_LAY } from './profile/defaults.js';
import type { Profile } from './profile/profile.js';

let active: string | null = null;

/** profiles are inlined for now; template ingest replaces this lookup. */
function profileFor(id: string): Profile {
  return { ...DEFAULT_LAY, id };
}

export function activeProfile(): string | null {
  return active;
}

/** idempotent: re-applying the same profile is a no-op, so a poll can call
 *  this freely. */
export function enterLay(profileId: string): void {
  if (active === profileId) return;
  applyStylesheet(toCss(profileFor(profileId)));
  active = profileId;
}

export function leaveLay(): void {
  if (active === null) return;
  removeStylesheet();
  active = null;
}

/** marker present -> lay, absent -> off. */
export function syncTo(marker: string | null): void {
  if (marker) enterLay(marker);
  else leaveLay();
}
