// what laymirror remembers between sessions.
//
// profiles are kept in a library rather than one at a time, because two
// schools' documents can be open at once and a single slot would make the
// second one silently wear the first one's template. a document keeps the
// profile it was marked with; a document that has never had one adopts
// whichever was used last.

import type { PageBreak } from './docx/breaks.js';
import { isProfile, type Profile } from './profile/profile.js';
import type { PluginApi } from './host/plugin-api.js';

const PROFILES = 'profiles';
const LAST_PROFILE = 'lastProfile';
const DOCS = 'docs';

export interface DocState {
  profileId: string | null;
  breaks: PageBreak[];
}

const EMPTY: DocState = { profileId: null, breaks: [] };

const asRecord = (value: unknown): Record<string, unknown> =>
  value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};

const asBreaks = (value: unknown): PageBreak[] =>
  Array.isArray(value)
    ? value.filter(
        (mark): mark is PageBreak =>
          !!mark &&
          typeof mark === 'object' &&
          typeof (mark as PageBreak).headingId === 'string' &&
          Number.isInteger((mark as PageBreak).offset),
      )
    : [];

export interface Store {
  profile(id: string | null): Profile | null;
  profiles(): Profile[];
  setProfile(profile: Profile): void;
  lastProfileId(): string | null;
  doc(key: string | null): DocState;
  setDoc(key: string, patch: Partial<DocState>): void;
}

export function store(api: PluginApi): Store {
  const profiles = (): Record<string, unknown> => asRecord(api.storage.get(PROFILES));
  const docs = (): Record<string, unknown> => asRecord(api.storage.get(DOCS));

  return {
    profile(id) {
      if (!id) return null;
      const found = profiles()[id];
      return isProfile(found) ? found : null;
    },

    profiles() {
      return Object.values(profiles()).filter(isProfile);
    },

    setProfile(profile) {
      api.storage.set(PROFILES, { ...profiles(), [profile.id]: profile });
      api.storage.set(LAST_PROFILE, profile.id);
    },

    lastProfileId() {
      const id = api.storage.get(LAST_PROFILE);
      return typeof id === 'string' ? id : null;
    },

    doc(key) {
      if (!key) return EMPTY;
      const state = asRecord(docs()[key]);
      return {
        profileId: typeof state['profileId'] === 'string' ? state['profileId'] : null,
        breaks: asBreaks(state['breaks']),
      };
    },

    setDoc(key, patch) {
      const all = docs();
      const current = asRecord(all[key]);
      api.storage.set(DOCS, { ...all, [key]: { ...current, ...patch } });
    },
  };
}
