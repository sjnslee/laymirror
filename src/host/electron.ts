// the file calls the save pipeline stands on, read off the preload at 1.8.0 and
// 1.10.0.
//
// `readFileAtPath` is scoped to paths the user has put in play and serves only
// .cmir/.docx, so a .dotx or .docm arrives through `openFile` — the os picker,
// which grants read scope on the way out.
//
// writes go through `saveExisting`, cardmirror's own in-place save. since 1.8.0,
// the manifest's floor, it refuses a path this window does not have open, and a file changed on disk since the
// window last read or wrote it, and it moves the window's baseline to the bytes
// written, so cardmirror's next save does not read laymirror's write as someone
// else's and keep both as a conflicted copy.

export interface FileStat {
  mtimeMs: number;
  size: number;
}

export interface ReadFile {
  bytes: Uint8Array;
}

export interface PickedFile {
  name: string;
  bytes: Uint8Array;
  handle: string;
}

interface ElectronApi {
  statFile(path: string): Promise<FileStat | null>;
  openFile?(opts: { filters: Filters }): Promise<PickedFile | null>;
  readFileAtPath(path: string): Promise<ReadFile | null>;
  saveExisting(path: string, bytes: Uint8Array): Promise<void>;
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
    typeof a.saveExisting === 'function'
  );
}

export async function statFile(path: string): Promise<FileStat | null> {
  return (await api()?.statFile(path)) ?? null;
}

export async function readFile(path: string): Promise<ReadFile | null> {
  return (await api()?.readFileAtPath(path)) ?? null;
}

export async function saveExisting(path: string, bytes: Uint8Array): Promise<void> {
  const a = api();
  if (!a) throw new Error('electronAPI unavailable — desktop only');
  await a.saveExisting(path, bytes);
}

type Filters = { name: string; extensions: string[] }[];

/** the os picker. null when the user cancelled, or the host has no picker. */
export async function openFile(filters: Filters): Promise<PickedFile | null> {
  const picked = await api()?.openFile?.({ filters });
  return picked && picked.bytes ? picked : null;
}

/** `.docm` is on the list because a lay template usually ships with macros */
export const WORD_FILES: Filters = [
  { name: 'Word document or template', extensions: ['docx', 'docm', 'dotx', 'dotm'] },
];

/** pointing at the open document: cardmirror only ever has a .docx open. */
export const DOCX_FILES: Filters = [{ name: 'Word document', extensions: ['docx'] }];
