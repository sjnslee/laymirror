# laymirror

lay debate plugin for [cardmirror](https://github.com/ant981228/cardmirror).

toggleable mode for lay debate with format scanning, font changes, pagination,
print utilities, and more

**status: in progress.** marker, work-view typography, template
ingest and the styles/section writers finished; the save pipeline,
pagination, page view and print in progress.

## details 

lay document is marked by a custom document property under the `.docx`. 
plugin is dormant without the marker.

laymirror reads template's fonts, sizes, weights, indents, spacing, 
page setup, header and footer from `styles.xml` and `sectPr`. 

## install

desktop only

```
npm install
npm run build
```

in cardmirror: settings → plugins → developer → "load plugin from
file…" → `plugin.js`.

commands triggered by keyboard shortcut (settings → keyboard
shortcuts → plugins), or by adding to ribbon with "+ add.

suggested: add **laymirror: open** to ribbon.

## development

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```


