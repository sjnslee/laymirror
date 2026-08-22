// the three file calls the save pipeline stands on. signatures read off the
// shipped preload/main (1.3.0) and all three exercised in the phase 0 spike.
//
// read is scoped (main only serves paths the user put in play) and only
// serves .cmir/.docx — anything else reads as null, which is why a school's
// .dotx template cannot come in this way. write is unscoped.

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

interface ElectronApi {
  statFile(path: string): Promise<FileStat | null>;
  pickFile?(opts?: {
    defaultPath?: string;
    title?: string;
    filters?: { name: string; extensions: string[] }[];
  }): Promise<unknown>;
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

/** ask the user which file, for when we cannot work out the path ourselves.
 *  cardmirror gives a plugin no reliable way to ask which document is open —
 *  `docInfo()` is null until a doc id is minted, and the recent-files list is
 *  a history — so rather than fail, laymirror asks. */
export async function pickDocx(title: string): Promise<string | null> {
  const picked = await api()?.pickFile?.({
    title,
    filters: [{ name: 'Word document', extensions: ['docx'] }],
  });
  if (typeof picked === 'string') return picked;
  const record = picked as { handle?: string; path?: string; filePath?: string } | null;
  return record?.handle ?? record?.path ?? record?.filePath ?? null;
}
