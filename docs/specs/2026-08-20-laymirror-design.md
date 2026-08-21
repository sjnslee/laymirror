# laymirror — design

a cardmirror plugin for lay debate documents.

## problem

cardmirror is built for tech debate. its formatting is verbatim's:
calibri, tiny card bodies, one look for every document, no pagination,
no headers or footers. lay debate needs the opposite — the document is a
printed artifact handed to a parent judge, so it needs readable serif
type, real pages, a header carrying team code / title / authors, a
footer with page numbers, and it needs to open in microsoft word looking
exactly like the school's template.

lay conventions differ per school, so the formatting must be
configurable rather than hardcoded.

## goals

- a `.docx` that opens natively in word as a lay document — right
  styles, right fonts, header, footer, live page numbers, correct
  attached template
- ctrl+s produces that file. no separate export step
- two views inside cardmirror: a **work view** (word's draft view) and a
  **page view** (word's print layout)
- per-school formatting, defined by pointing at that school's template
- fully inert when the document isn't a lay document

## non-goals

- `.cmir`. lay documents are docx-only
- web or lite editions. cardmirror plugins are desktop-only anyway
- byte-exact word pagination (see *fidelity* below)
- changing how tech documents look or export

---

## host constraints

verified against cardmirror 1.3.0. these shape everything below.

| finding | consequence |
| --- | --- |
| plugin api v1 is commands + settings + storage only. no css hook, no prosemirror hook, no export hook, **no lifecycle events** | every integration point is unofficial. quarantined in `src/host/` |
| `plugin.js` runs in the renderer main world via `webFrame.executeJavaScript` | full dom + `localStorage` + `window.electronAPI` access |
| ctrl+s is a **native electron menu accelerator** dispatched from main; `window.electronAPI` is contextBridge-exposed under `contextIsolation: true, sandbox: true` | save cannot be intercepted. we react to it instead |
| `electronAPI.statFile(path)` → `{mtimeMs, size}`, unscoped | save detection by polling |
| `electronAPI.readFileAtPath` is scoped, but a file the user opened this session is granted — and grants persist across sessions in an LRU journal; `writeFileAtPath` is **entirely unscoped** | read + rewrite the saved docx in place |
| `readFileAtPath` serves **only `.cmir` and `.docx`** — every other extension reads as `null` | a school's `.dotx` cannot be ingested through it; the template picker needs `pickFile`/`openFile`, or v1 takes `.docx` donors only |
| `cmirDocId` is absent until cardmirror itself saves the file | a lay `.docx` that word owns has no id to match on, so path resolution cannot rely on it alone |
| there is no command palette; `defaultKey` auto-binds only when the chord is free | how the six commands are actually invoked is an open question, not a given |
| display already runs on css custom properties — `--pmd-size-*`, `--pmd-body-font`, `--pmd-color-*` — plus stable classes `.pmd-tag`, `.pmd-card-body`, `.pmd-cite-para`, `.pmd-undertag`, `.pmd-analytic` under `#editor` / `.pmd-pane-editor` | restyling is an injected stylesheet. `!important` beats core's inline writes and survives the user touching appearance settings |
| only one `bodyFont` for the whole document | per-type fonts are ours to add |
| exporter emits one hardcoded `<w:sectPr>` (letter, 1" margins, **no header/footer refs**) and never writes `header1.xml` / `footer1.xml` | pagination parts are entirely ours |
| exporter maps `tag`→`Heading4`, and emits card bodies and cites with **no pstyle** | style ids must be rewritten to the template's |
| the importer has **two paths**. it takes the *native* one — marks matched by `styleId` — only when the document's styles contain all of `Style13ptBold`/`Style 13 pt Bold`, `StyleUnderline`/`Style Underline`, and `Emphasis`. otherwise it falls back to the *legacy* one, which matches paragraphs by lowercased `w:name` and character styles against a small fixed table | writing real lay style names is round-trip safe for `Tag`, `Cite`, `card` and `Underline`, which are all in the legacy table — but the donor lacks two of the three sentinel styles, so **cite marks are silently lost** unless we also emit `StyleUnderline` and `Emphasis`. they cost nothing in word |
| heading depth comes from `w:outlineLvl`, mapped `1→pocket, 2→hat, 3→block, 4→tag` | heading names don't matter; outline levels do |
| `<w:br w:type="page"/>` imports as a plain line break and exports as one | **cardmirror destroys manual page breaks.** we own them |
| canonical `Heading1/2/3` all carry `pageBreakBefore`; the lay template puts it only on `Heading1` | page-break-before is a per-type property |
| `word/settings.xml` `<w:attachedTemplate>` is how cardmirror makes verbatim recognize its exports | same lever points word at the lay `.dotx` |
| autosave only writes `.cmir` — docx is skipped as too expensive | lay saves are always explicit |
| github-install release assets cap at 5 MiB each; the "load plugin from file" dev path has **no cap** | distribution is the constrained path, local development isn't |
| `openExternal` rejects anything but `http(s):` and `mailto:` | we cannot hand the docx to word to print. our print path is the only one |

## the format model

**a school's template docx is the profile.** it already carries
everything: `styles.xml` has font, size, weight, casing, underline,
color, alignment, indent, spacing and `pageBreakBefore` per style;
`sectPr` has page size and margins; `header1.xml` / `footer1.xml` have
the header and footer with their `PAGE` / `NUMPAGES` fields.

so laymirror does not define a profile format. it ingests a template and
stores two things:

```
profile = {
  templateBytes      // the donor docx, stored with the profile
  mapping            // cardmirror type -> template styleId
  overrides          // per-type tweaks the template doesn't cover
  attachedTemplate   // basename, e.g. "Lay Cut Cards.dotx"
  fontSubstitutions  // family -> local fallback
}
```

the mapping exists because the vocabularies differ:

| cardmirror type | exports as | a school'"'"'s lay template |
| --- | --- | --- |
| pocket | `Heading1` | `Heading1` (serif, large, small caps, centered, page-break-before) |
| hat | `Heading2` | `Heading2` |
| block | `Heading3` | `Heading3` |
| tag | `Heading4` | **`Tag`** (body serif, bold) |
| cite_paragraph | *(none)* | **`Cite`** (tag + thick underline) |
| card_body | *(none)* | **`card`** (body size, indented both sides) |
| underline_mark | `StyleUnderline` | **`Underline`** |
| cite_mark | `Style13ptBold` | `Style13ptBold` ✓ |
| analytic / undertag | `Analytic` / `Undertag` | *(absent — override)* |

one profile is rendered two ways, from the same object:

- **`toCss(profile)`** → an injected stylesheet for the editor
- **`toOoxml(profile)`** → `styles.xml`, `header1.xml`, `footer1.xml`,
  `sectPr`, `settings.xml`

these must not drift. one profile, two renderers, shared tests.

## states

- **off** — not a lay document. the plugin does nothing at all: no
  stylesheet, no watcher, no rewriting.
- **lay / work view** — lay typography over cardmirror's normal
  continuous layout. no page chrome, but page breaks are *indicated*.
  this is word's draft view, which is what cardmirror already
  approximates.
- **lay / page view** — pages, margins, header, footer, page numbers.
  read-only in v1. this is what prints.

page view is read-only for v1 by decision; editable pages stay open as a
later change. work view is where editing happens.

### activation

the marker lives in the file, as a custom document property beside
cardmirror's own `cmirDocId`:

```xml
<property name="layMirrorProfile"><vt:lpwstr>ashford-lay</vt:lpwstr></property>
```

it travels to teammates, survives a word round-trip (cardmirror's own
comment says this is verified), and our save rewrite re-stamps it every
time. toggling lay on/off is a command.

### resolving the open document's path

`api.docInfo()` gives `docId` + filename; `localStorage['pmd-recent-files']`
gives `{handle: absolutePath, filename, format}`. filename matching is
ambiguous when two open documents share a name, so we confirm by reading
the candidate's `cmirDocId` from `docProps/custom.xml` and comparing to
`docInfo().docId`. exact, and the result is cached.

## save pipeline

ctrl+s can't be hooked, so we react:

1. watcher polls `statFile(path)` on the active lay document
2. mtime bump ⇒ a save landed
3. `readFileAtPath` → unzip → rewrite → zip → `writeFileAtPath`

the rewrite:

- replace `word/styles.xml` with the template's, plus overrides
- rewrite `document.xml` style references to the template's ids, and add
  `<w:pStyle>` to card bodies and cite paragraphs (cardmirror emits none)
- restore manual page breaks from our markers
- add `header1.xml` / `footer1.xml`, their relationships, and
  `[Content_Types].xml` overrides
- replace `sectPr` with page size, margins, and header/footer references
- point `<w:attachedTemplate>` at the lay `.dotx`, **basename only** —
  the donor template hardcodes an absolute path containing a real
  person's home directory, and word only basename-matches
- re-stamp the `layMirrorProfile` marker

a command triggers the same rewrite on demand, for when waiting on the
poll isn't wanted.

this is a watcher standing in for a missing api. it is contained to
`src/host/watcher.ts` and would collapse to a few lines if cardmirror
ever ships a save hook.

## pagination

`docx-preview` (apache-2.0, ~975 KB — fits the asset cap) was the
obvious candidate for page view, but it only breaks pages at markers
already in the file: manual breaks, `w:lastRenderedPageBreak`, and page
setting changes. word writes those hints; cardmirror never does. so it
would render one enormous page.

**we write the paginator either way.** given that, v1 skips the
dependency: page view is a read-only clone of the editor dom, measured
and split into fixed-height page boxes, with header and footer chrome
drawn around each. one styling path serves both views, so they cannot
disagree, and print needs no browser re-pagination because content is
already placed into page boxes.

the paginator honors page size and margins, per-type
`page-break-before`, `keep-with-next` / `keep-lines`, widow and orphan
control, and manual breaks. work view draws word's dotted page-break
indicators from the same computation.

### fidelity

exact word pagination is not achievable at reasonable cost. superdoc
does clear that bar — its layout painter uses word's own metrics rather
than the browser's — but it cannot be used here.

size and agpl are not the reason (an earlier draft of this spec said
they were; both were wrong). superdoc's browser bundle is 1.2 MB, well
under the cap. the reason is that superdoc 2's fidelity lives in
`@superdoc/docx-engine`, a **separate proprietary package** that the
superdoc tarball deliberately excludes. its license permits use only as
a dependency of superdoc, and §3.1(d) forbids redistributing it or
making it available to a third party. shipping it inside a `plugin.js`
release asset is exactly that. a commercial license from harbour
enterprises would clear it; nothing else does.

we write our own paginator from the ooxml spec and word's documented
layout behavior. we do not inspect, benchmark, or reimplement the
superdoc engine.

what we get: page breaks matching word in the large majority of cases,
drifting on long documents. two things bound the damage.

**the printed docx is always correct.** word does its own layout and
`PAGE` / `NUMPAGES` are live fields, so a judge's copy is right even
when our preview was a page off. the risk is confined to our preview and
our own print button.

**fonts are the main drift source.** the template needs palatino
linotype and garamond. cardmirror bundles metric-compatible substitutes
for calibri, cambria, times new roman, arial, georgia, verdana and
others — but not those two. we reuse cardmirror's `font-detect.ts`
approach to detect missing families, warn with the names, and let the
user pick a per-family fallback in settings.

## print

print the page view. page boxes already match the target page size, so
`@page` only needs matching size and zero margin — the browser does no
re-pagination. header and footer repeat because they are part of each
page box, and page numbers are ours to fill.

## settings panel

cardmirror's declared-settings api renders only boolean / text / number
/ select, which can't express the mapping table. the plugin opens its
own modal instead:

- **template** — pick a school's `.docx` / `.dotx`; shows the styles found
- **mapping** — each cardmirror text type, with a dropdown of the
  template's styles and a live preview
- **overrides** — font, size, weight, casing, underline, color,
  alignment, indent, spacing, page-break-before, for types the template
  doesn't cover
- **header / footer** — title, authors, team code; recovered by parsing
  the existing `header1.xml` on open rather than retyped
- **fonts** — missing families and their fallbacks

## repo layout

```
src/
  main.ts                 registration, commands, wiring
  profile/
    profile.ts            types, defaults
    read-template.ts      template docx -> profile
    mapping.ts            type <-> styleId
  render/
    css.ts                profile -> stylesheet
    paginate.ts           measure + split
    page-view.ts          read-only paged surface
    draft-marks.ts        page-break indicators
    print.ts
  docx/
    zip.ts                fflate wrapper
    rewrite.ts            apply profile to a saved docx
    styles.ts             styles.xml
    headers.ts            header1/footer1 + rels + content types
    sect.ts               sectPr
    marker.ts             lay marker in docProps/custom.xml
  host/                   *** every undocumented cardmirror internal ***
    cardmirror.ts         dom classes, css vars, localStorage keys
    watcher.ts            statFile polling
    paths.ts              docId -> path
  ui/
    settings-panel.ts
    fonts.ts
tests/
cardmirror-plugin.json
build.mjs                 esbuild -> plugin.js
```

`src/host/` is the quarantine. every assumption about cardmirror's
internals lives there, documented with the version it was verified
against, so a v2 plugin api or a cardmirror refactor is a contained fix.

## build order

**0. spike — done, passed.** all three calls are reachable from a
dev-loaded plugin, and a rewritten docx opens cleanly in word. see
*phase 0 results* in the implementation plan.

1. skeleton — plugin registers, commands appear, lay on/off marker
   round-trips through `docProps/custom.xml`
2. lay css from the built-in generic profile. work view looks right
3. template ingest, mapping, settings panel
4. **save pipeline** — watcher + rewrite. the core deliverable: ctrl+s
   yields a native lay docx that word opens correctly
5. paginator, page view, draft-view break indicators, manual page
   breaks — these are one piece of work, since the indicators and the
   breaks fall out of the same computation
6. print
7. font detection and substitution

phase 4 is the one that has to work. phases 1–3 are prerequisites for
testing it honestly. all seven are v1.

## risks

| risk | mitigation |
| --- | --- |
| plugin api v2 sandboxes plugins and closes the renderer access this depends on | `src/host/` quarantine; open an upstream issue asking for a save hook and a style hook |
| watcher races another writer, or fires mid-save on a partial file | verify zip integrity before rewriting; skip and retry on a bad read |
| pagination drifts from word | fields are live so the printed file is right; warn on missing fonts |
| distribution — the installer allowlist currently holds one repo | dev-load path and the `__plugins('community-on')` console unlock work today; ask the maintainer for a listing when it's worth sharing |
| cardmirror eats manual page breaks | we carry them ourselves through the rewrite |

## open questions

- should tables get lay treatment in v1, or inherit whatever the
  template's `TableNormal` says?
- multi-pane: page view in one pane while another edits — allowed, or
  one page view per window?
- does anyone need two lay profiles active at once (two schools' files
  open together)? the marker is per-file, so it costs little to allow.
