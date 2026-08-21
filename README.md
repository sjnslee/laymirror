# laymirror

lay debate plugin for [cardmirror](https://github.com/ant981228/cardmirror).

toggleable mode for lay debate with format scanning, font changes, pagination,
print utilities, and more

**status: all seven build phases are in** — marker, work-view typography,
template ingest, save pipeline, paginator, page view, print, font
substitution. not yet carried through a live round.

## details

lay document is marked by a custom document property under the `.docx`.
plugin is dormant without the marker.

laymirror reads template's fonts, sizes, weights, indents, spacing,
page setup, header and footer from `styles.xml` and `sectPr`.

every save of a marked document is read back and rewritten: the template's
styles, its header and footer, its page setup, its attached template, and a
real page break wherever one was asked for.

page view lays the document out on the template's page and prints it. work
view draws word's dotted rule where each page will break.

## caveats

- the paginator is ours, not word's, so a long document can drift a page.
  `PAGE` and `NUMPAGES` stay live fields, so the printed file is right
  regardless — the drift is confined to the preview.
- page view's header is built from the title, authors and team code typed
  into the panel. the printed `.docx` carries the template's own header.
- a manual page break is the text `[page break]` alone on a line, because
  cardmirror's model has nowhere else to keep one.

## install

desktop only

```
npm install
npm run build
```

in cardmirror: settings → plugins → developer → "load plugin from
file…" → `plugin.js`.

commands triggered by keyboard shortcut (settings → keyboard
shortcuts → plugins), or by adding to ribbon with "+ add".

suggested: add **laymirror: open** to ribbon. everything else — page view,
break marks, insert break, print — is a button inside it.

## development

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```
