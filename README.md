# laymirror

lay debate plugin for [cardmirror](https://github.com/ant981228/cardmirror).

toggleable lay mode: the school's own styles, header, footer and margins put
back on every save, plus a page view of the document as word will print it.

**status: rebuilt around docx-preview. not yet carried through a live round.**

## why it exists

cardmirror's exporter rebuilds the `.docx` from scratch on every save — its
own `styles.xml`, one hardcoded letter section with 1in margins, and no
header, footer or theme at all (`src/export/exporter.ts:111`). so a lay
document loses the school's format every time it is saved.

laymirror puts it back.

## how it works

a lay document is marked by a custom document property inside the `.docx`.
without the marker the plugin is inert.

loading a school template snapshots its identity — `header1.xml`,
`footer1.xml`, the section, `theme1.xml`, `fontTable.xml` and `styles.xml` —
verbatim. every save then asks one question:

- **does the file carry a header reference?** only word writes one, so the
  file is word's and its header is authoritative. laymirror re-adopts it and
  rewrites nothing. a team code you typed in word survives forever.
- **does it not?** cardmirror just stripped it. laymirror restores the
  snapshot and remaps cardmirror's exported style ids onto the ones the
  template defines — `Heading4` → `Tag`, `Style13ptBold` → `Cite`.

styles are carried, never rebuilt, so `smallCaps`, thick underlines and
borders survive by construction rather than by being parsed.

page view renders the real file with
[docx-preview](https://github.com/VolodymyrBaydalka/docxjs), which resolves
`basedOn` chains and theme fonts the way word does. pages break, in order of
preference, where word said they broke (`w:lastRenderedPageBreak`), where you
asked, or — only when neither exists — where a measuring pass over the
rendered layout puts them. the last case says so on screen.

manual page breaks are held outside the document, anchored to cardmirror's
stable heading ids, and injected as real `<w:br w:type="page"/>` on save.
cardmirror's model cannot carry a break: its importer turns one into a plain
newline and its exporter writes it back as a line break.

## caveats

- pagination is exact only for a file word has saved. otherwise it is
  chromium laying out the right styles at the right size — close, and capable
  of drifting a line over a long document. `PAGE` and `NUMPAGES` stay live
  fields, so the printed file is right regardless.
- page view shows the file on disk. it refuses over unsaved edits rather than
  saving on your behalf.
- `open in word` is in the panel, and is the only perfectly faithful preview.
- a plugin cannot place a ribbon button. laymirror ships default shortcuts
  instead.

## install

desktop only

```
npm install
npm run build
```

in cardmirror: settings → plugins → developer → "load plugin from
file…" → `plugin.js`.

default shortcuts:

| | |
| --- | --- |
| `⌘⌥L` | open laymirror |
| `⌘⌥P` | page view |
| `⌘⌥↩` | insert or remove a page break |

for a ribbon button: settings → ribbon → "+ add button" and pick a laymirror
command. cardmirror allows up to 10 custom buttons.

## development

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```
