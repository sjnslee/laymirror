// cardmirror's sanctioned plugin surface (api v1). shapes read off the
// shipped renderer bundle's registration validator.

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

export function register(def: PluginDefinition): boolean {
  const w = window as unknown as { __registerCardMirrorPlugin?: Register };
  if (typeof w.__registerCardMirrorPlugin !== 'function') return false;
  w.__registerCardMirrorPlugin(def);
  return true;
}
