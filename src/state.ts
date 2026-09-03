// what laymirror remembers between sessions.
//
// templates are kept as a library rather than one at a time, because two
// schools' documents can be open at once and a single slot would make the
// second one silently wear the first one's format. a document keeps the
// template it was marked with; a document that has never had one adopts
// whichever was used last.
//
// the template is stored as the file — cardmirror hands a plugin a json
// storage bag backed by localStorage, so the bytes travel as base64. a school
// template is tens of kilobytes; one carrying a crest is a few hundred. the
// cap is here so a user who picks a 20mb file is told why it was refused
// rather than watching the bag silently fail to write.

import type { PluginApi } from './host/plugin-api.js';
import type { Values } from './docx/fields.js';
import type { Template } from './template/template.js';

const TEMPLATES = 'templates';
const LAST_TEMPLATE = 'lastTemplate';
const DEFAULTS = 'defaults';
const DOCS = 'docs';

/** roughly a fifth of a chromium origin's localStorage, and far more than any
 *  real template needs. */
export const TEMPLATE_LIMIT = 2_000_000;

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
      // the next document off the same template starts where this one ended:
      // a team code and a cutter's name are the same all season. merged rather
      // than replaced, so setting one field does not forget the others.
      if (!templateId) return;
      const shared = asValues(defaults()[templateId]);
      api.storage.set(DEFAULTS, { ...defaults(), [templateId]: { ...shared, ...values } });
    },
  };
}
