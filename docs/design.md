# Design

## Cardmirror

Cardmirror's exporter rebuilds the docx on every save with a fresh
`styles.xml` and one letter section at 1in margins.

Plugin api: `docInfo()`, `showToast()`, json storage bag, declared settings,
commands with key chords.

## Host internals

`src/host/`:

| What | Where |
| --- | --- |
| Open document's path | `pmd-recent-files` in localStorage, matched against the filename chip |
| File read and write | `window.electronAPI`, writing through `saveExisting` |
| Save detection | polling `statFile` |
| Storage bag before any command runs | `localStorage['plugin:laymirror']` |
| Whether laymirror is still switched on | `enabled` in `localStorage['pmd-plugins']` |

`readFileAtPath` takes `.cmir` and `.docx`. Templates load through `openFile`,
the os picker.

Read against cardmirror 1.8.0 and 1.10.0, in `src/host/cardmirror.ts`,
`src/host/electron.ts` and `src/template/styles.ts`.

Api v1 has no unload hook, and a bundle cardmirror has already run cannot be
unloaded, so the enabled flag is what lets laymirror stop its own timers when
the user switches it off.

## Pipeline

```
cardmirror saves ──► watcher sees the mtime move ──► read the file
                                                        │
       template ──► blueprint ──► apply ◄───────────────┘
                                    │
                                    ├─ remap cardmirror's style ids onto the template's
                                    ├─ restore styles, theme, fonts, numbering, header, footer
                                    ├─ restore the section: page size, margins, header refs
                                    ├─ fill the header fields
                                    └─ write the file back ──► resync the watcher
```

## Template

Template file: base64 in the storage bag, cut to the parts below plus what reads them. Parts are
copied byte for byte:

| Part | Sets |
| --- | --- |
| `styles.xml` | styles |
| `theme1.xml`, `fontTable.xml` | fonts |
| `numbering.xml` | lists |
| `header*.xml`, `footer*.xml` | header and footer, plus rels and image parts |
| body `sectPr` | page size, margins, header refs |

Header fields are read from the template, never from the saved file.

A page break lives on the style, in `styles.xml`.
