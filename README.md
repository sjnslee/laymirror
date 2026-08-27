# laymirror

lay debate plugin for [cardmirror](https://github.com/ant981228/cardmirror).

lay debate is printed and judged by a parent, so a lay file has to come off the
printer looking like the school's other files. cardmirror rebuilds the `.docx`
from scratch on every save and keeps none of that. laymirror puts it back.

**status: prototype. not yet carried through a live round.**

## what it does

**wears the school's template.** load your school's `.docx`, `.docm`, `.dotx` or
`.dotm` once. every save afterwards gets its styles, theme, fonts, numbering,
header, footer, page size and margins put back — carried across byte for byte,
so smallCaps, thick underlines, borders and a crest in the header all survive by
construction rather than by being parsed.

**fills in the header.** laymirror never builds a header. it finds the two or
three words inside the school's own that change — the team code, the year, the
file title, the cutter's name — and offers them as fields with an apply button.
`PAGE` and `NUMPAGES` stay live fields, so page numbers stay word's to compute.

if your template marks its placeholders by wrapping them in a zero-width
character, laymirror uses those. otherwise it offers every stretch of plain text
between the header's tabs and word fields, which on a real lay header is exactly
the right three things.

**shows where the pages end.** a lay template starts a new page before every
heading 1, and cardmirror shows the document as one unbroken column. laymirror
draws a rule where the page breaks, read out of the template's own styles. it
also marks a literal page-break character, which cardmirror imports as text and
draws nothing for.

**prints.** page view renders the real file with
[docx-preview](https://github.com/VolodymyrBaydalka/docxjs) at the template's
real page size, and prints it. electron's print dialog carries "save as pdf".

## install

desktop only.

```
npm install
npm run build
```

in cardmirror: settings → plugins → developer → "load plugin from
file…" → `plugin.js`.

| | |
| --- | --- |
| `⌘⇧L` | open laymirror |
| `⌘⇧P` | page view |

a plugin cannot place a ribbon button, so laymirror ships shortcuts instead. for
a button: settings → ribbon → "+ add button" and pick a laymirror command.

## using it

1. `⌘⇧L`
2. **load…** and pick your school's template
3. **turn on**
4. type the header fields, **apply**

turning it on writes the school's format onto the file straight away, and every
cardmirror save from then on gets it again. a marker travels inside the `.docx`,
so a file a teammate marked arrives already lay.

## caveats

- pagination in page view is exact for a file word has saved and for one the
  template breaks itself. failing both it is chromium laying out the right
  styles at the right size — close, and capable of drifting a line over a long
  document. the bar says which.
- page view shows the file on disk, not unsaved edits.
- laymirror finds the open document through cardmirror's recent-files history.
  two open files with the same name are ambiguous, and it says so rather than
  rewriting the wrong one.
- `⌘⇧L` binds only if cardmirror has not already taken the chord; `⌘⌥L` is the
  fallback. every command is in the palette either way.

## development

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```

`tests/plugin.test.ts` drives the whole thing — registration, panel, template
load, toggle, apply — the way cardmirror does. drop a school template at
`local/lay-template.docm` and `tests/donor.test.ts` runs the pipeline against a
real one; `local/` is gitignored, because a school's template is theirs.

[docs/design.md](docs/design.md) is why any of it is shaped this way.
