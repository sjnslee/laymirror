// what laymirror remembers between sessions.
//
// templates are a library rather than one slot: two documents off two templates
// can be open at once. one that has never had a template takes the last loaded.
//
// the storage bag is json in localStorage, so a template travels as base64 and
// the cap below turns a silent quota failure into a message.

import type { PluginApi } from './host/plugin-api.js';
import type { Located } from './host/paths.js';
import type { Values } from './docx/fields.js';
import type { Template } from './template/template.js';

const TEMPLATES = 'templates';
const LAST_TEMPLATE = 'lastTemplate';
const DEFAULTS = 'defaults';
const DOCS = 'docs';
const LOCATED = 'located';

/** every kept template together, about a tenth of a chromium origin's
 *  localStorage once base64. the bag shares that origin with cardmirror's own
 *  settings and recent files, whose writes fail silently when it is full. */
export const TEMPLATE_LIMIT = 1_000_000;

/** keyed by the document's full path, so two files that share a name never
 *  share a switch, a template or a header. */
export interface DocState {
  templateId: string | null;
  /** what the user typed into this document's header fields. */
  values: Values;
  on: boolean;
}

const EMPTY: DocState = { templateId: null, values: {}, on: false };

const asRecord = (value: unknown): Record<string, unknown> =>
  value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};

const asValues = (value: unknown): Values => {
  const out: Values = {};
  for (const [key, held] of Object.entries(asRecord(value))) {
    if (typeof held === 'string') out[key] = held;
  }
  return out;
};

export function encode(bytes: Uint8Array): string {
  let binary = '';
  // one argument per byte overflows the stack on a real template
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

export function decode(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export interface Store {
  template(id: string | null): Template | null;
  /** name and path without the bytes: decoding a template runs to megabytes of
   *  base64, and most callers only want to say which one is loaded. */
  templateInfo(id: string | null): { name: string; path: string | null } | null;
  /** false when the library would go over `TEMPLATE_LIMIT`, even after
   *  dropping templates no document uses. */
  addTemplate(template: Template): boolean;
  lastTemplateId(): string | null;
  doc(key: string | null): DocState;
  setDoc(key: string, patch: Partial<DocState>): void;
  /** the values a document should show: what it was given, over whatever was
   *  last used with the same template. */
  valuesFor(key: string | null, templateId: string | null): Values;
  setValues(key: string, templateId: string | null, values: Values): void;
  /** the path last pointed at, or found, for a filename. */
  located(filename: string): Located | null;
  setLocated(filename: string, located: Located): void;
}

export function store(api: PluginApi): Store {
  const templates = (): Record<string, unknown> => asRecord(api.storage.get(TEMPLATES));
  const docs = (): Record<string, unknown> => asRecord(api.storage.get(DOCS));
  const defaults = (): Record<string, unknown> => asRecord(api.storage.get(DEFAULTS));

  const held = (id: string): { name: string; path: string | null; docx: string } | null => {
    const record = asRecord(templates()[id]);
    if (typeof record['name'] !== 'string' || typeof record['docx'] !== 'string') return null;
    return {
      name: record['name'],
      path: typeof record['path'] === 'string' ? record['path'] : null,
      docx: record['docx'],
    };
  };

  return {
    template(id) {
      if (!id) return null;
      const record = held(id);
      if (!record) return null;
      try {
        return { id, name: record.name, path: record.path, docx: decode(record.docx) };
      } catch {
        return null;
      }
    },

    templateInfo(id) {
      if (!id) return null;
      const record = held(id);
      return record ? { name: record.name, path: record.path } : null;
    },

    addTemplate(template) {
      const docx = encode(template.docx);
      const inUse = new Set(
        Object.values(docs()).map((state) => asRecord(state)['templateId']),
      );
      // a template no document points at is only taking room
      const kept = Object.entries(templates()).filter(
        ([id]) => id !== template.id && inUse.has(id),
      );
      // base64 is four characters for every three bytes
      const size = (base64: unknown): number =>
        typeof base64 === 'string' ? Math.floor((base64.length * 3) / 4) : 0;
      const total = kept.reduce((sum, [, record]) => sum + size(asRecord(record)['docx']), 0);
      if (total + size(docx) > TEMPLATE_LIMIT) return false;

      api.storage.set(TEMPLATES, {
        ...Object.fromEntries(kept),
        [template.id]: { name: template.name, path: template.path, docx },
      });
      api.storage.set(LAST_TEMPLATE, template.id);
      return true;
    },

    lastTemplateId() {
      const id = api.storage.get(LAST_TEMPLATE);
      return typeof id === 'string' ? id : null;
    },

    doc(key) {
      if (!key) return EMPTY;
      const state = asRecord(docs()[key]);
      return {
        templateId: typeof state['templateId'] === 'string' ? state['templateId'] : null,
        values: asValues(state['values']),
        on: state['on'] === true,
      };
    },

    setDoc(key, patch) {
      const all = docs();
      api.storage.set(DOCS, { ...all, [key]: { ...asRecord(all[key]), ...patch } });
    },

    valuesFor(key, templateId) {
      const shared = templateId ? asValues(defaults()[templateId]) : {};
      return { ...shared, ...this.doc(key).values };
    },

    setValues(key, templateId, values) {
      this.setDoc(key, { values });
      // the next document off this template starts where this one ended.
      // replaced, not merged: `values` is every field on screen, and a cleared
      // field is absent from it, so a merge would bring its old text back
      if (!templateId) return;
      api.storage.set(DEFAULTS, { ...defaults(), [templateId]: values });
    },

    located(filename) {
      const record = asRecord(asRecord(api.storage.get(LOCATED))[filename]);
      return typeof record['path'] === 'string' && typeof record['at'] === 'number'
        ? { path: record['path'], at: record['at'] }
        : null;
    },

    setLocated(filename, located) {
      api.storage.set(LOCATED, { ...asRecord(api.storage.get(LOCATED)), [filename]: located });
    },
  };
}
