// esbuild inlines these as data urls, so an import is the file's contents.
declare module '*.woff2' {
  const dataUrl: string;
  export default dataUrl;
}
