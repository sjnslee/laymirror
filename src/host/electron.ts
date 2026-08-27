// the three file calls the save pipeline stands on. signatures read off the
// shipped preload/main (1.3.0) and all three exercised in the phase 0 spike.
//
// `readFileAtPath` is scoped — main only serves paths the user has put in play
// this session or a past one — and only serves .cmir/.docx, so a school's .dotx
// or macro-enabled .docm cannot come in that way. `openFile` can: it puts the
// picker in front of the user, reads whatever they chose, and grants the path
// read scope on the way out. writes are unscoped.

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

/** resolves 'collision' rather than throwing when the target exists and the
 *  caller asked for failIfExists. */
export async function writeFile(
  path: string,
  bytes: Uint8Array,
  opts?: { failIfExists?: boolean },
): Promise<'collision' | undefined> {
  const a = api();
  if (!a) throw new Error('electronAPI unavailable — desktop only');
  return a.writeFileAtPath(path, bytes, opts);
}

/** put the os picker in front of the user and read what they chose. null when
 *  they cancelled, or when the host has no picker at all. */
export async function openFile(
  filters: { name: string; extensions: string[] }[],
): Promise<PickedFile | null> {
  const picked = await api()?.openFile?.({ filters });
  return picked && picked.bytes ? picked : null;
}

/** the formats a school hands out a template in. `.docm` is on the list
 *  because a lay template usually ships with the squad's macros attached. */
export const WORD_FILES = [
  { name: 'Word document or template', extensions: ['docx', 'docm', 'dotx', 'dotm'] },
];
