# laymirror

lay debate is slower, printed, and judged by a parent. so a lay file is a
*document*: it has the school's header on every page, the school's fonts, one
speech per page, and it has to come off the printer looking like the school's
other files. cardmirror is built for the other kind of debate, where none of
that is true.

laymirror is the difference.

## what cardmirror leaves us

a plugin is one classic script run in the renderer's main world. the sanctioned
api is small — `docInfo()`, `showToast()`, a json storage bag, declared
settings, and commands with default key chords. it cannot touch the document,
cannot hook a save, cannot add a ribbon button and cannot add a settings page.

everything else laymirror needs comes from the renderer it is running inside,
and is kept in `src/host/`:

| what | where | why it is not the api |
| --- | --- | --- |
| the open document's path | `pmd-recent-files` in localStorage, matched against the filename chip | `docInfo()` is null until cardmirror mints a doc id, which a word-authored `.docx` never has |
| reading and writing the file | `window.electronAPI` | there is no file access in the plugin api |
| knowing a save happened | polling `statFile` | there is no save hook |

`readFileAtPath` is scoped by the main process and serves only `.cmir` and
`.docx`, so a template arrives through `openFile` — the os picker — which reads
any extension and grants the path read scope on the way out. `writeFileAtPath`
is unscoped, and it updates cardmirror's own changed-on-disk baseline, so a
file laymirror rewrites does not make the next cardmirror save complain.

## the one fact everything follows from

cardmirror's exporter **rebuilds the package from scratch on every save**. its
own `styles.xml`, one hardcoded letter section with 1in margins, relationships
for styles and settings and nothing else. no header. no footer. no theme.

so laymirror is not a formatter. it is the thing that puts back what the
exporter has just thrown away, every time the exporter runs.

```
cardmirror saves ──► watcher sees the mtime move ──► read the file
                                                        │
       template ──► blueprint ──► apply ◄────────────────┘
                                    │
                                    ├─ remap cardmirror's style ids onto the school's
                                    ├─ restore styles, theme, fonts, numbering, header, footer
                                    ├─ restore the section: page size, margins, header refs
                                    ├─ fill the header fields with what the user typed
                                    └─ write it back ──► resync the watcher
```

## the template is the truth

a template is stored as **the file itself**, base64 in the plugin's storage bag,
and everything laymirror knows is derived from it on demand. an earlier design
stored a digested profile, which meant every new thing laymirror learned to read
— numbering, a crest in the header, a page break in a style — needed the user to
load their template again.

parts are carried **verbatim**. an earlier design parsed the template into a
model and re-emitted it, which silently dropped every property nobody remembered
to parse: `smallCaps`, thick underlines, borders. bytes cannot forget, and they
carry a school's crest as readily as its fonts.

what travels: `styles.xml`, `theme1.xml`, `fontTable.xml`, `numbering.xml`,
every header and footer, everything those headers and footers relate to (their
own `.rels`, their images), the body `sectPr`, and the attached template's
basename. word resolves the theme and font table through relationships that
cardmirror never writes, so those are added — a theme with no relationship is a
part word never reads, and `asciiTheme="minorHAnsi"` would resolve to nothing.

it is also authoritative every time. an earlier version asked whether word had
written the file and adopted its header if so, which made the header something
you edited in word and laymirror preserved. that is backwards for a squad: the
school's header is fixed, and the two or three words inside it that change are
typed into laymirror's panel.

## header fields

a school header is fixed except for a team code, a year, a file title and a
cutter's name. laymirror finds those inside the school's own header rather than
building one:

- **marked** — the template wraps a placeholder in a zero-width character.
  explicit, so it wins.
- **inferred** — no marks anywhere, so every stretch of plain text between tabs
  and word fields is offered. on a real lay header that is exactly
  `BCP 26-27`, `File Title` and `Name`.

a word field's own decoration is excluded: " page " and " of " read as plain
text but are not the user's to edit, and the number between them is a result
word recomputes. a value lands whole in the first run it covers and the rest are
emptied rather than removed, because the run carries the small caps.

discovery always runs against the pristine template, so a field keeps its
identity after its value has been replaced, and a value can be typed over.

## page breaks

cardmirror's model cannot hold one. its importer turns `<w:br w:type="page"/>`
into a bare `\n` inside a text node and its exporter writes every `\n` back as
`<w:br/>`, so the type is gone. `<w:pageBreakBefore/>` in a paragraph's own
`pPr` is dropped by both.

what survives is the **style**. a lay template says `w:pageBreakBefore` on
heading 1, so every pocket starts a page — and because it is a property of the
style, it travels inside the `styles.xml` laymirror carries and word and
docx-preview both honour it. nobody types a page break in a lay file; the
template does it.

so the editor draws two things:

- a css rule above every block type whose mapped style breaks before it, read
  out of the template rather than remembered. free, and it never fights
  prosemirror for the dom.
- an overlay line wherever the text itself carries a literal form feed — a
  document that came from somewhere else. a css rule cannot select one
  character, so those are measured with ranges and drawn in a layer above the
  editor. nothing is observed until such a character actually turns up.

## page view and print

`docx-preview` renders the real package — the school's `styles.xml`, theme,
header and footer, at the template's real page size and margins — and chromium
lays it out. pages break where word said they broke
(`w:lastRenderedPageBreak`), or where the document breaks itself, or, only when
neither exists, where a fill pass over the rendered layout puts them. the bar
says which.

this is also the print pipeline. a plugin cannot hand a file to word:
cardmirror's `openExternal` accepts `http(s)` and `mailto` and refuses
everything else, so there is no `ms-word:` or `file://` escape. electron's print
dialog carries "save as pdf", which is the pdf export as well.

## what breaks when cardmirror changes

every undocumented internal is in `src/host/cardmirror.ts` and
`src/template/styles.ts`, stamped with the version it was read against (1.3.0),
so an upgrade breaks a test rather than a round.

the sharpest edges: the block classes (`h1.pmd-pocket` and friends) come from
cardmirror's schema `toDOM`; the export style ids (`Heading4` for a tag,
`Style13ptBold` for a cite mark) come from its exporter; the native/legacy
import split — which decides whether a style comes back as a tag or as an
ordinary paragraph — comes from its parse worker.
