// what laymirror remembers between sessions.
//
// templates are a library rather than one slot: two documents off two templates
// can be open at once. a document keeps the template it was marked with; one
// that has never had a template adopts whichever was used last.
//
// cardmirror's storage bag is json in localStorage, so the file travels as
// base64 and the cap below is what turns a silent quota failure into a message.

import type { PluginApi } from './host/plugin-api.js';
import type { Values } from './docx/fields.js';
import type { Template } from './template/template.js';

const TEMPLATES = 'templates';
const LAST_TEMPLATE = 'lastTemplate';
const DEFAULTS = 'defaults';
const DOCS = 'docs';

/** roughly a fifth of a chromium origin's localStorage */
export const TEMPLATE_LIMIT = 2_000_000;

export interface DocState {
  templateId: string | null;
  /** what the user typed into this document's header fields. */
  values: Values;
  on: boolean;
  /** where this document was last found on disk. cardmirror's history is the
   *  first answer; this is what stands in when it has no entry to give. */
  path: string | null;
}

const EMPTY: DocState = { templateId: null, values: {}, on: false, path: null };

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
  addTemplate(template: Template): void;
  lastTemplateId(): string | null;
  doc(key: string | null): DocState;
  setDoc(key: string, patch: Partial<DocState>): void;
  /** the values a document should show: what it was given, over whatever was
   *  last used with the same template. */
  valuesFor(key: string | null, templateId: string | null): Values;
  setValues(key: string, templateId: string | null, values: Values): void;
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

    addTemplate(template) {
      api.storage.set(TEMPLATES, {
        ...templates(),
        [template.id]: {
          name: template.name,
          path: template.path,
          docx: encode(template.docx),
        },
      });
      api.storage.set(LAST_TEMPLATE, template.id);
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
        path: typeof state['path'] === 'string' ? state['path'] : null,
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
      // the next document off the same template starts where this one ended.
      // merged rather than replaced, so setting one field keeps the others.
      if (!templateId) return;
      const shared = asValues(defaults()[templateId]);
      api.storage.set(DEFAULTS, { ...defaults(), [templateId]: { ...shared, ...values } });
    },
  };
}
