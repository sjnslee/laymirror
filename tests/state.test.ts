// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import { decode, encode, store } from '../src/state.js';
import type { PluginApi } from '../src/host/plugin-api.js';

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

const docx = (byte: number) => new Uint8Array([byte, 2, 3]);

let api: PluginApi;
beforeEach(() => {
  api = fakeApi();
});

describe('base64', () => {
  it('round-trips bytes a json bag cannot hold', () => {
    const bytes = new Uint8Array([0, 255, 13, 10, 200]);
    expect(decode(encode(bytes))).toEqual(bytes);
  });

  // one argument per byte overflows the stack, and a template with a crest is
  // big enough to reach it
  it('survives a template-sized file', () => {
    const big = new Uint8Array(300_000).map((_, i) => i % 256);
    expect(decode(encode(big))).toEqual(big);
  });
});

describe('templates', () => {
  it('keeps a library rather than one slot', () => {
    const bag = store(api);
    bag.addTemplate({ id: 'template:a.docx', name: 'a.docx', docx: docx(1) });
    bag.addTemplate({ id: 'template:b.docx', name: 'b.docx', docx: docx(2) });
    expect(bag.templates().map((t) => t.id)).toEqual(['template:a.docx', 'template:b.docx']);
  });

  it('gives the bytes back unchanged', () => {
    const bag = store(api);
    bag.addTemplate({ id: 'a', name: 'a.docx', docx: docx(9) });
    expect(bag.template('a')!.docx).toEqual(docx(9));
  });

  it('remembers the one used last', () => {
    const bag = store(api);
    bag.addTemplate({ id: 'a', name: 'a.docx', docx: docx(1) });
    bag.addTemplate({ id: 'b', name: 'b.docx', docx: docx(2) });
    expect(bag.lastTemplateId()).toBe('b');
  });

  it('is null for a template that was never loaded', () => {
    expect(store(api).template('gone')).toBeNull();
  });
});

describe('documents', () => {
  it('is off and untemplated until told otherwise', () => {
    expect(store(api).doc('1ac.docx')).toEqual({ templateId: null, values: {}, on: false });
  });

  it('patches one field without losing the others', () => {
    const bag = store(api);
    bag.setDoc('1ac.docx', { templateId: 'a', on: true });
    bag.setDoc('1ac.docx', { on: false });
    expect(bag.doc('1ac.docx').templateId).toBe('a');
  });

  // two schools' documents can be open at once, and a single slot would make
  // the second wear the first one's format
  it('keeps documents apart', () => {
    const bag = store(api);
    bag.setDoc('ours.docx', { templateId: 'a' });
    bag.setDoc('theirs.docx', { templateId: 'b' });
    expect(bag.doc('ours.docx').templateId).toBe('a');
  });

  it('has nothing to say about a document with no name', () => {
    expect(store(api).doc(null).on).toBe(false);
  });
});

describe('header values', () => {
  it("gives back what the document was given", () => {
    const bag = store(api);
    bag.setValues('1ac.docx', 'a', { code: 'BCP 26-27' });
    expect(bag.valuesFor('1ac.docx', 'a')).toEqual({ code: 'BCP 26-27' });
  });

  // a team code and a cutter's name are the same all season, so the next
  // document off the same template starts where the last one ended
  it('seeds a fresh document from the last one off the same template', () => {
    const bag = store(api);
    bag.setValues('1ac.docx', 'a', { code: 'BCP 26-27', title: 'Aff' });
    expect(bag.valuesFor('2ac.docx', 'a')).toEqual({ code: 'BCP 26-27', title: 'Aff' });
  });

  it("lets a document override what it inherited", () => {
    const bag = store(api);
    bag.setValues('1ac.docx', 'a', { code: 'BCP 26-27', title: 'Aff' });
    bag.setValues('2ac.docx', 'a', { title: 'Neg' });
    expect(bag.valuesFor('2ac.docx', 'a')).toEqual({ code: 'BCP 26-27', title: 'Neg' });
  });

  it('does not leak across templates', () => {
    const bag = store(api);
    bag.setValues('1ac.docx', 'a', { code: 'BCP 26-27' });
    expect(bag.valuesFor('theirs.docx', 'b')).toEqual({});
  });
});
