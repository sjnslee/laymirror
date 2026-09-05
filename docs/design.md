# laymirror

a lay file is a document: a header on every page, the template's fonts, one
speech per page. laymirror puts that back on the file cardmirror writes.

## what cardmirror leaves us

a plugin is one classic script in the renderer's main world. the sanctioned api
is `docInfo()`, `showToast()`, a json storage bag, declared settings, and
commands with default key chords. it cannot touch the document, hook a save, add
a ribbon button or add a settings page. the api object is only ever handed to a
command's `run()`, so it arrives too late to start watching with.

everything else comes from the renderer, and is kept in `src/host/`:

| what | where |
| --- | --- |
| the open document's path | `pmd-recent-files` in localStorage, matched against the filename chip |
| reading and writing the file | `window.electronAPI` |
| knowing a save happened | polling `statFile` |
| the storage bag, before any command has run | `localStorage['plugin:laymirror']` |

the history is not complete: cardmirror writes no entry for a document it hands
to a window it spawned — every open after the first, and every finder
double-click. a `.docx` the history cannot place is asked for through the picker
and the answer kept against the filename.

`readFileAtPath` is scoped by the main process to `.cmir` and `.docx`, so a
template arrives through `openFile` — the os picker — which reads any extension
and grants the path read scope. `writeFileAtPath` is unscoped and updates
cardmirror's changed-on-disk baseline, so a file laymirror rewrites does not
make the next cardmirror save complain.

## the one fact everything follows from

cardmirror's exporter **rebuilds the package from scratch on every save**: its
own `styles.xml`, one hardcoded letter section with 1in margins, relationships
for styles and settings and nothing else. no header, footer or theme.

```
cardmirror saves ──► watcher sees the mtime move ──► read the file
                                                        │
       template ──► blueprint ──► apply ◄───────────────┘
                                    │
                                    ├─ remap cardmirror's style ids onto the template's
                                    ├─ restore styles, theme, fonts, numbering, header, footer
                                    ├─ restore the section: page size, margins, header refs
                                    ├─ fill the header fields with what the user typed
                                    └─ write it back ──► resync the watcher
```

## the template is the truth

a template is stored as **the file itself**, base64 in the storage bag, and
everything laymirror knows is derived from it on demand.

parts are carried **verbatim**, never parsed into a model and re-emitted.

what travels: `styles.xml`, `theme1.xml`, `fontTable.xml`, `numbering.xml`,
every header and footer, everything those relate to (their own `.rels`, their
images), the body `sectPr`, and the attached template's basename. word resolves
the theme and font table through relationships cardmirror never writes, so those
are added — `asciiTheme="minorHAnsi"` would otherwise resolve to nothing.

the template is authoritative on every apply. the header is fixed; the two or
three words inside it that change are typed into laymirror's panel.

## header fields

laymirror finds the editable stretches inside the template's own header rather
than building one:

- **marked** — the template wraps each placeholder in `U+200B`. the author has
  said what is editable, so nothing else in that paragraph is offered.
- **inferred** — no marks anywhere, so every stretch of plain text between tabs
  and word fields is offered instead.

a value is written **between** the marks, never over them, so the field survives
into the next read. marks are paired in document order, and one left without a
partner drops its field rather than shifting every field after it.

a word field's own decoration (" page ", " of ") reads as plain text but is
excluded; the number between them is a result word recomputes.

a value lands whole in the first run it covers and the rest are emptied rather
than removed, because the run carries the small caps. word splits a placeholder
across runs on its own revision ids — `26-27` arrives as `2` then `6-27` — so
writing into each run separately would double the value.

discovery always runs against the pristine template, so a field keeps its
identity after its value has been replaced.

## page breaks are the template's job

cardmirror's model cannot hold one. its importer turns `<w:br w:type="page"/>`
into a bare `\n` and its exporter writes every `\n` back as `<w:br/>`;
`<w:pageBreakBefore/>` in a paragraph's `pPr` is dropped by both.

what survives is the **style**. a lay template puts `w:pageBreakBefore` on
heading 1, so every pocket starts a page, and it travels inside the `styles.xml`
laymirror carries.

## no page view

there is no page rendering, and no "open in word": `openExternal` accepts
`http(s)` and `mailto` only, and the `shell.openPath` calls in the main process
take fixed directories. no ipc channel hands an arbitrary path to the os.

## saying what happened

none of laymirror's work shows up in cardmirror — the header, the fonts and the
page setup are on the file, not on screen. so the panel says what the last write
did and when, a failed write says why in a status line of laymirror's own, and
header values are held as typed rather than on pressing apply, so a plain ⌘S
writes what is on screen.

## what breaks when cardmirror changes

every undocumented internal is in `src/host/cardmirror.ts` and
`src/template/styles.ts`, stamped with the version it was read against (1.3.0),
so an upgrade breaks a test rather than a round.

the sharpest edges: the export style ids (`Heading4` for a tag, `Style13ptBold`
for a cite mark) come from its exporter; the native/legacy import split comes
from its parse worker; `pmd-recent-files` and `#doc-name-chip-text` are how a
filename becomes a path; and `plugin:<id>` is where the storage bag lives.
