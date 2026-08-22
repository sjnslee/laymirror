import { describe, expect, it, beforeEach } from 'vitest';
import { store } from '../src/store.js';
import { DEFAULT_PROFILE } from '../src/profile/defaults.js';
import type { PluginApi } from '../src/host/plugin-api.js';
import type { Profile } from '../src/profile/profile.js';

/** cardmirror's storage is a json bag per plugin. */
function fakeApi(): PluginApi {
  const bag: Record<string, unknown> = {};
  return {
    storage: {
      get: (key: string) => bag[key],
      set: (key: string, value: unknown) => {
        bag[key] = value;
      },
    },
  } as unknown as PluginApi;
}

const profile = (id: string, name = id): Profile => ({ ...DEFAULT_PROFILE, id, name });

let api: PluginApi;
beforeEach(() => {
  api = fakeApi();
});

describe('profiles', () => {
  it('keeps a library rather than one slot', () => {
    const bag = store(api);
    bag.setProfile(profile('template:a.docx'));
    bag.setProfile(profile('template:b.docx'));
    expect(bag.profiles().map((p) => p.id).sort()).toEqual([
      'template:a.docx',
      'template:b.docx',
    ]);
  });

  it('remembers the one used last', () => {
    const bag = store(api);
    bag.setProfile(profile('template:a.docx'));
    bag.setProfile(profile('template:b.docx'));
    expect(bag.lastProfileId()).toBe('template:b.docx');
  });

  it('is null for an unknown id', () => {
    expect(store(api).profile('template:missing.docx')).toBeNull();
  });

  it('survives a bag holding junk', () => {
    api.storage.set('profiles', 'not an object');
    expect(store(api).profiles()).toEqual([]);
  });
});

describe('documents', () => {
  it('starts empty', () => {
    expect(store(api).doc('1ac.docx')).toEqual({ profileId: null, breaks: [] });
  });

  it('keeps a profile per document', () => {
    const bag = store(api);
    bag.setDoc('1ac.docx', { profileId: 'template:a.docx' });
    bag.setDoc('1nc.docx', { profileId: 'template:b.docx' });
    expect(bag.doc('1ac.docx').profileId).toBe('template:a.docx');
    expect(bag.doc('1nc.docx').profileId).toBe('template:b.docx');
  });

  it('patches without dropping what it is not given', () => {
    const bag = store(api);
    bag.setDoc('1ac.docx', { profileId: 'template:a.docx' });
    bag.setDoc('1ac.docx', { breaks: [{ headingId: 'h1', offset: 2 }] });
    expect(bag.doc('1ac.docx')).toEqual({
      profileId: 'template:a.docx',
      breaks: [{ headingId: 'h1', offset: 2 }],
    });
  });

  it('is empty for a document with no name', () => {
    expect(store(api).doc(null)).toEqual({ profileId: null, breaks: [] });
  });

  // storage is json a previous version wrote, so a malformed mark must be
  // dropped rather than reaching the injector
  it('discards breaks that are not breaks', () => {
    api.storage.set('docs', {
      '1ac.docx': { breaks: [{ headingId: 'h1', offset: 1 }, { headingId: 5 }, null, 'x'] },
    });
    expect(store(api).doc('1ac.docx').breaks).toEqual([{ headingId: 'h1', offset: 1 }]);
  });
});
