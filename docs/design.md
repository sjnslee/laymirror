# design

## constraints

the plugin api is `docInfo()`, `showToast()`, a json storage bag, declared
settings, and commands with key chords. it cannot reach the document, hook a
save, add a ribbon button or add a settings page. the api object is passed only
to a command's `run()`, too late to start a watcher from.

cardmirror's exporter rebuilds the package on every save: its own `styles.xml`,
one hardcoded letter section at 1in margins, relationships for styles and
settings. no header, no footer, no theme. laymirror puts those back after each
save.

`openExternal` takes `http(s)` and `mailto`, and `shell.openPath` is called on
fixed directories, so no ipc channel opens a file in word.

## host internals

`src/host/` holds what comes from the renderer rather than the api:

| what | where |
| --- | --- |
| open document's path | `pmd-recent-files` in localStorage, matched against the filename chip |
| file read and write | `window.electronAPI` |
| save detection | polling `statFile` |
| storage bag before any command runs | `localStorage['plugin:laymirror']` |

`pmd-recent-files` gets no entry for a document cardmirror hands to a spawned
window, which is every open after the first. an unplaceable filename is asked
for through the picker and the answer kept against it.

`readFileAtPath` is scoped to `.cmir` and `.docx`, so a template loads through
`openFile`, the os picker, which reads any extension and grants read scope on
the path. `writeFileAtPath` is unscoped and updates cardmirror's changed-on-disk
baseline, so a rewrite does not make the next save complain.

## pipeline

```
cardmirror saves ──► watcher sees the mtime move ──► read the file
                                                        │
       template ──► blueprint ──► apply ◄───────────────┘
                                    │
                                    ├─ remap cardmirror's style ids onto the template's
                                    ├─ restore styles, theme, fonts, numbering, header, footer
                                    ├─ restore the section: page size, margins, header refs
                                    ├─ fill the header fields
                                    └─ write it back ──► resync the watcher
```

## template

stored as the file itself, base64 in the storage bag; everything else derives
from it on demand. parts are copied verbatim, never parsed and re-emitted.

carried: `styles.xml`, `theme1.xml`, `fontTable.xml`, `numbering.xml`, every
header and footer, whatever those relate to (`.rels`, images), the body
`sectPr`, and the attached template's basename. the theme and font table also
need relationships, which cardmirror never writes, or `asciiTheme="minorHAnsi"`
resolves to nothing.

the template wins on every apply. the header is fixed; the few words that change
are typed into the panel.

## header fields

editable stretches are found in the template's own header:

- marked: placeholders wrapped in `U+200B`. nothing else in that paragraph is
  offered.
- inferred: no marks anywhere, so each stretch of plain text between tabs and
  word fields is offered.

values are written between the marks, so a field survives into the next read.
marks pair in document order; an unpaired one drops its field instead of
shifting the rest.

a word field's decoration (" page ", " of ") reads as plain text and is skipped.

a value lands whole in the first run it covers and later runs are emptied rather
than removed, since the run carries the small caps. word splits a placeholder on
its own revision ids, so `26-27` arrives as `2` then `6-27`.

discovery runs against the pristine template, so a field keeps its identity
after its value is replaced.

## page breaks

cardmirror's importer turns `<w:br w:type="page"/>` into `\n` and its exporter
writes `\n` back as `<w:br/>`; `<w:pageBreakBefore/>` in a paragraph's `pPr` is
dropped by both. only the style survives, and a lay template puts
`w:pageBreakBefore` on heading 1, which travels inside `styles.xml`.

## feedback

no part of an apply is visible in the editor, so the panel reports the last
write and a failure says why. header values are held as typed, so a plain ⌘S
writes what is on screen.

## versioned internals

`src/host/cardmirror.ts` and `src/template/styles.ts` are stamped with the
cardmirror version they were read against (1.3.0). export style ids (`Heading4`
for a tag, `Style13ptBold` for a cite mark) come from its exporter; the
native/legacy import split from its parse worker; `pmd-recent-files` and
`#doc-name-chip-text` map a filename to a path; `plugin:<id>` is the storage
bag.
