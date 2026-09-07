// jsdom ships no `localStorage` here, and the recent-files history is the only
// way to turn a filename into a path.

export function stubStorage(): void {
  const bag = new Map<string, string>();
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: {
      getItem: (key: string) => bag.get(key) ?? null,
      setItem: (key: string, value: string) => void bag.set(key, String(value)),
      removeItem: (key: string) => void bag.delete(key),
      clear: () => bag.clear(),
      key: (i: number) => [...bag.keys()][i] ?? null,
      get length() {
        return bag.size;
      },
    },
  });
}
