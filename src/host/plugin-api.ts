// cardmirror's sanctioned plugin surface (api v1). shapes read off the
// shipped renderer bundle's registration validator.

import { storageKey } from './cardmirror.js';

export interface DocInfo {
  docId: string;
  docTitle: string;
}

/** handed to a command's run(). only the members laymirror uses are typed. */
export interface PluginApi {
  appVersion: string;
  docInfo(): DocInfo | null;
  showToast(message: string): void;
  /** persisted per plugin, json-serialised into localStorage. */
  storage: {
    get(key: string): unknown;
    set(key: string, value: unknown): void;
  };
  /** declared settings only — get plus a change subscription, not a record. */
  settings: {
    get(key: string): unknown;
    onChanged(handler: (key: string) => void): () => void;
  };
}

export interface Command {
  /** must start with `<pluginId>.` or registration is rejected. */
  id: string;
  label: string;
  keywords?: string[];
  /** auto-binds during the keymap rebuild, but only when the chord is free. */
  defaultKey?: string | string[];
  run(api: PluginApi): void | Promise<void>;
}

export interface PluginDefinition {
  id: string;
  name: string;
  apiVersion: 1;
  commands: Command[];
}

type Register = (def: PluginDefinition) => void;

/** the same api, minus the parts only cardmirror can provide.
 *
 *  the real object is built at registration and handed only to a command's
 *  `run()`, so this stands in until one runs: it reads and writes the identical
 *  storage bag, and its toast goes to the console because there is no reaching
 *  cardmirror's. a document is watched from the moment laymirror loads, and the
 *  first command upgrades the session to the real thing. */
export function bootApi(pluginId: string): PluginApi {
  const key = storageKey(pluginId);
  const bag = (): Record<string, unknown> => {
    try {
      const held: unknown = JSON.parse(localStorage.getItem(key) || '{}');
      return held && typeof held === 'object' ? (held as Record<string, unknown>) : {};
    } catch {
      return {};
    }
  };

  return {
    appVersion: 'unknown',
    docInfo: () => null,
    showToast: (message) => console.log(`[laymirror] ${message}`),
    storage: {
      get: (name) => bag()[name],
      set: (name, value) => {
        const next = { ...bag(), [name]: value };
        try {
          localStorage.setItem(key, JSON.stringify(next));
        } catch {
          // a storage bag that will not write is the template being too large;
          // the caller's cap is what reports that
        }
      },
    },
    settings: { get: () => undefined, onChanged: () => () => {} },
  };
}

export function register(def: PluginDefinition): boolean {
  const w = window as unknown as { __registerCardMirrorPlugin?: Register };
  if (typeof w.__registerCardMirrorPlugin !== 'function') return false;
  w.__registerCardMirrorPlugin(def);
  return true;
}
