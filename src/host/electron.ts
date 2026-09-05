// the file calls the save pipeline stands on. signatures read off the shipped
// preload/main at 1.3.0.
//
// `readFileAtPath` is scoped to paths the user has put in play, and serves only
// .cmir/.docx — so a .dotx or macro-enabled .docm can only come in through
// `openFile`, which puts the picker up and grants read scope on the way out.
// writes are unscoped.

export interface FileStat {
  mtimeMs: number;
  size: number;
}

export interface ReadFile {
  name: string;
  bytes: Uint8Array;
  handle: string;
  format: 'cmir' | 'docx';
}

export interface PickedFile {
  name: string;
  bytes: Uint8Array;
  handle: string;
}

export interface PickOptions {
  filters?: { name: string; extensions: string[] }[];
}

interface ElectronApi {
  statFile(path: string): Promise<FileStat | null>;
  openFile?(opts?: PickOptions): Promise<PickedFile | null>;
  readFileAtPath(path: string): Promise<ReadFile | null>;
  writeFileAtPath(
    path: string,
    bytes: Uint8Array,
    opts?: { failIfExists?: boolean },
  ): Promise<'collision' | undefined>;
}

function api(): ElectronApi | null {
  const w = window as unknown as { electronAPI?: ElectronApi };
  return w.electronAPI ?? null;
}

export function hasFileApi(): boolean {
  const a = api();
  return (
    !!a &&
    typeof a.statFile === 'function' &&
    typeof a.readFileAtPath === 'function' &&
    typeof a.writeFileAtPath === 'function'
  );
}

export async function statFile(path: string): Promise<FileStat | null> {
  return (await api()?.statFile(path)) ?? null;
}

export async function readFile(path: string): Promise<ReadFile | null> {
  return (await api()?.readFileAtPath(path)) ?? null;
}

/** resolves 'collision' rather than throwing when failIfExists hits one */
export async function writeFile(
  path: string,
  bytes: Uint8Array,
  opts?: { failIfExists?: boolean },
): Promise<'collision' | undefined> {
  const a = api();
  if (!a) throw new Error('electronAPI unavailable — desktop only');
  return a.writeFileAtPath(path, bytes, opts);
}

/** the os picker. null when the user cancelled, or the host has no picker. */
export async function openFile(
  filters: { name: string; extensions: string[] }[],
): Promise<PickedFile | null> {
  const picked = await api()?.openFile?.({ filters });
  return picked && picked.bytes ? picked : null;
}

/** `.docm` is on the list because a lay template usually ships with macros */
export const WORD_FILES = [
  { name: 'Word document or template', extensions: ['docx', 'docm', 'dotx', 'dotm'] },
];

/** pointing at the open document: cardmirror only ever has a .docx open. */
export const DOCX_FILES = [{ name: 'Word document', extensions: ['docx'] }];
