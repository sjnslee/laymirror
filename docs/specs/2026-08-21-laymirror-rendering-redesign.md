# laymirror — rendering redesign

supersedes the rendering, style and header/footer sections of
`2026-08-20-laymirror-design.md`. the marker model, activation model,
zip layer and save-watcher survive unchanged.

## why this revision exists

the first build was written against the shipped renderer bundle, not
against cardmirror's source, and against assumptions about the school
template rather than the template itself. reading both turned up four
faults, each of which alone would have made page view unusable.

| fault | evidence |
| --- | --- |
| page view measured virtualized dom | `src/editor/style.css:2771,3153` — `.pmd-card`, `.pmd-pocket`, `.pmd-hat`, `.pmd-block` all carry `content-visibility: auto`. our offscreen stage (`left:-10000px; visibility:hidden`) guaranteed every block was skipped from layout, so every height we read was the `contain-intrinsic-height` placeholder (200px / 40px), never real geometry |
| a card was treated as one block | `src/schema/nodes.ts:641` — `card` is a `div.pmd-card` *containing* tag / cite / undertag / body. `.ProseMirror`'s children are card wrappers, not paragraphs |
| lay state was derived from a file read | opening the panel ran `readState()` → `resolveDocPath()`; a failed resolve returned `marker: null`, and `syncTo(null)` called `leaveLay()`. so opening the panel turned lay off in state while its stylesheet stayed on screen — and the toggle then turned it back *on* |
| styles were reconstructed instead of carried | we parsed the donor into an intermediate model and re-emitted it, dropping every property we forgot to parse and never resolving `basedOn` chains or `majorHAnsi` |

the first two explain "lines intersect", "line spacing is horribly
off", "half the text doesn't appear". the third explains "fonts stay lay
after turning lay off" and "it says lay is not on for some documents".
the fourth explains "the hat isn't close to the right font size" and
"cite has no underline".

## corrected host findings

verified against cardmirror source at `ant981228/cardmirror@main`, not
the shipped bundle.

| finding | consequence |
| --- | --- |
| `src/export/exporter.ts:111` hardcodes `<w:sectPr><w:pgSz .../><w:pgMar w:top="1440" .../></w:sectPr>` and emits no `header1.xml`, no `footer1.xml`, no `headerReference`, no theme | **every cardmirror save destroys the school's header, footer, custom styles, theme and margins.** laymirror's job is restoration, not decoration |
| `src/editor/plugin-api.ts:166` — `docInfo()` returns null unless a `docId` has been minted | unusable for identifying a word-authored `.docx`. name the document from the chip instead |
| `src/editor/plugin-registry.ts` — a command may declare `defaultKey` | laymirror declares none. that is why there is no way to reach it |
| `src/editor/settings.ts:328` — `MAX_RIBBON_CUSTOM_BUTTONS = 10` | plugins cannot place a ribbon button. users add one in settings; plugin command ids appear in that list. document it, ship a `defaultKey`, stop pretending otherwise |
| the preload exposes 187 methods, none of them shell, exec or `printToPDF`; `saveExisting(handle, bytes)` takes bytes the *renderer* supplies, and the exporter is not exposed to plugins | we cannot serialize live editor state. page view renders the last saved file |
| `openExternal` exists | "open in word" is one call, and is the only perfectly faithful preview there will ever be |

## what the template actually is

`~/code/src/template.docx`, read part by part.

- page: `12240 × 15840` twips, margins `720 / 720 / 1008 / 720`,
  header `720`, footer `720`. that is 0.5in sides and top, 0.7in bottom
- `Normal` is **palatino linotype 10pt** — garamond appears only in the
  header and footer
- `Heading1` (pocket): times new roman, bold, smallCaps, 20pt, centred,
  `pageBreakBefore`
- `Heading2` (hat): `majorHAnsi` → calibri, bold, `#4F81BD`, 13pt
- the school's own `Tag`, `Cite` (`basedOn` Tag, thick underline) and
  `UnderlineBold` character style
- header: two paragraphs — team code ⇥ speech name, then a literal
  `Team Code` ⇥ `Page N of M` — garamond bold smallCaps 12pt, using
  `w:ptab` positional tabs. the rule is the bottom border of an empty
  floating text box anchored behind the text
- footer: centred underscore run + `Page N of M`, thick underline
- **page numbers appear in both the header and the footer**, and there
  are no authors anywhere

the previous design synthesised a footer, invented an `{{authors}}`
token, and wrote 1in margins.

## architecture

```
                   ┌─ marked lay ─────────────────────────────┐
                   │                                          │
  ctrl+s ──▶ cardmirror writes a bare .docx                    │
                   │                                          │
             watcher sees the write                           │
                   │                                          │
             does the file have a headerReference?            │
                   ├── yes → word wrote it → re-snapshot ─────┤
                   └── no  → cardmirror stripped it           │
                              │                               │
                        restore snapshot + remap style ids ───┘
                              │
                              ▼
                    a real word document on disk
                              │
              ┌───────────────┴───────────────┐
              │                               │
      docx-preview (page view)          open in word
      real styles / theme / header      perfectly faithful
```

## state

the rule that broke lay toggling: **screen state and file state are
independent, and a failed file read may never change what is on
screen.**

- whether lay is on lives in memory, mirrored into `api.storage`. it is
  the only authority for what the stylesheet does
- the in-file marker is reconciliation, not truth. opening the panel
  reads it to *adopt* lay when the file says lay and memory does not —
  it can never turn lay off
- `leaveLay()` runs only from an explicit user toggle
- a document laymirror cannot resolve to a path degrades to "lay on,
  file not tracked", never to "lay off". the toggle keeps working on
  screen either way

## save pipeline — snapshot and restore

marking a document lay snapshots its **document identity** from disk,
verbatim, into plugin storage keyed by the document:

- `word/header1.xml`, `word/footer1.xml` (and any other
  header*/footer* parts), plus their `[Content_Types].xml` overrides
  and `document.xml.rels` entries
- the `w:sectPr` element
- `word/theme/theme1.xml`, `word/fontTable.xml`, `word/styles.xml`
- `settings.xml`'s `attachedTemplate` relationship

on every save the watcher asks one question, and only one:

- **the file has a `headerReference`** — only word writes one, so word
  last touched this file. it is authoritative. **re-snapshot, rewrite
  nothing.**
- **it does not** — cardmirror just stripped it. **restore the
  snapshot**, then remap the style ids cardmirror exported onto the ids
  the school's `styles.xml` defines.

this is self-healing and needs no ui. a team code typed into the header
in word survives every subsequent cardmirror save, because the next
save re-adopts it rather than overwriting it.

a document with no snapshot yet falls back to the profile's template.

### style id remapping

cardmirror's exporter emits the verbatim set (`src/ooxml/styles.ts`):
`Heading1`–`Heading4`, `Style13ptBold`, `StyleUnderline`, `Emphasis`,
`Undertag`, `Analytic`. the profile maps each onto a style id the
school's `styles.xml` actually defines — for this template
`Heading4`→`Tag`, `Style13ptBold`→`Cite`, `StyleUnderline`→
`UnderlineBold`, the rest identity.

cite paragraphs and card bodies leave cardmirror with no `pStyle` at
all, and `card` is flattened away on export, so they are classified
from the run marks they carry. that logic is correct today and is kept.

### styles are carried, never rebuilt

`buildStylesXml` is deleted. the school's `styles.xml` and `theme1.xml`
travel byte for byte. word then sees exactly the school's styles rather
than our reconstruction of them, and every property we never thought to
parse — `smallCaps`, `kern`, `u w:val="thick"`, `pBdr`, `jc` — survives
by construction.

## rendering

`docx-preview` (apache-2.0, 75 kb minified, one dependency) renders the
on-disk file. verified headlessly against the real template:

```
pages: 2
page style: width=612pt minHeight=792pt pad=36pt/36pt/50.4pt/36pt
header: "BCP 26-27  1AC   Team Code  Page 1 of 1"
footer: "______…______Page 1 of 1"
```

612 × 792pt is 8.5 × 11in and the padding is the template's exact
margins, 0.7in bottom included. the emitted css resolves everything the
old code did not:

```css
.docx { --docx-majorHAnsi-font: Calibri; --docx-accent1-color: #4F81BD; }
.docx p, p.docx_normal span { font-family: 'Palatino Linotype'; font-size: 10.00pt }
p.docx_heading1 span { font-family: 'Times New Roman'; font-weight: bold;
                       font-variant: small-caps; font-size: 20.00pt }
p.docx_heading2 span { font-family: var(--docx-majorHAnsi-font);
                       font-weight: bold; color: #4F81BD; font-size: 13.00pt }
p.docx_cite span    { font-family: 'Palatino Linotype'; font-weight: bold;
                       text-decoration: underline; font-size: 10.00pt }
```

`Cite` resolves through two levels of `basedOn`, `Heading4` resolves a
`themeShade`. this is the whole reason for the dependency.

`docx-preview` emits one `<section>` per page at the real page size
with the real margins, and re-renders the document's own header and
footer into every page (`createPageElement`, `renderHeaderFooter`). we
therefore need no page boxes, no running-header machinery and no
`paged.js` — layering paged.js on top would re-chunk a dom that is
already page-structured, and the two would fight.

### where page breaks come from

`docx-preview` does not paginate; it breaks where the file says to.
three sources, in priority order:

1. **`w:lastRenderedPageBreak`** — per ECMA-376 §17.3.3.13, the
   position that ended a page "when this document was last saved by an
   application which paginates its content". word writes these on every
   save. rendering with `ignoreLastRenderedPageBreak: false` replays
   **word's own pagination**, not an emulation of it. this is the exact
   path, and it is available for any file word has touched
2. **`w:br w:type="page"`** — manual breaks the user inserted
3. **neither** — cardmirror has just saved and stripped them. fall back
   to a measured fill: render the flow once, walk the *real* rendered
   dom (docx-preview's output has no `content-visibility`, so its
   geometry is truthful), insert breaks where content overflows the
   content box, re-render. iterate to stability

path 3 is approximate and page view says so in one line of chrome:
*approximate — save from word for exact pages*. paths 1 and 2 say
nothing, because they are exact.

the old `paginate.ts` predicted heights from a model. the replacement
observes overflow in real layout, and only ever runs when word's own
answer is unavailable.

### fidelity, stated honestly

- word last saved it → **exact**, because we replay word's breaks
- otherwise → chromium's line breaking, with correct fonts, sizes,
  styles and margins. close, and capable of drifting a line over a long
  document. nothing free closes that gap
- always available → **open in word**, one `openExternal` call

`Palatino Linotype` is not installed on macos (`Palatino` is), so a
substitution ui stays, now scoped to fonts the *template* asks for.
cardmirror bundles metric-compatible faces for arial, calibri, cambria,
times new roman and georgia. the embedded eb garamond is deleted: it
was added for body text that is palatino.

## work view

the editor keeps a css stylesheet, but stops carrying a second style
engine. the template's css is generated **by docx-preview** —
`parseAsync` + `renderDocument` over the template alone — and the
selectors are rewritten from `.docx_tag` onto cardmirror's `.pmd-tag`
and friends. one style resolver serves both views, so work view and
page view can no longer disagree.

## ui and access

- every command declares a `defaultKey`: `Mod-Alt-L` panel,
  `Mod-Alt-P` page view
- durable configuration moves to cardmirror's own plugin-settings
  modal, as the ebb plugin does. the panel keeps only what the modal
  cannot express: template upload, per-document metadata, page view,
  open in word
- the readme documents adding a ribbon button, and names the cap of 10
- template upload replaces the file input's "no file chosen" with the
  loaded filename

## deleted

`render/paginate.ts`, `render/measure.ts`, `render/page-view.ts`,
`render/print.ts`, `render/draft-marks.ts`, `render/embedded-fonts.ts`,
`docx/styles.ts`, `ui/fonts.ts`, the synthesis half of
`docx/headers.ts`, and their tests — roughly 1,200 lines and 80 kb of
embedded font.

kept as-is: `docx/zip.ts`, `docx/marker.ts`, `docx/xml.ts`,
`docx/sect.ts`, `host/electron.ts`, `host/watcher.ts`,
`host/plugin-api.ts`, `profile/mapping.ts`.

rewritten:

| file | becomes |
| --- | --- |
| `profile/read-template.ts` | extracts the identity snapshot verbatim and lists the style ids the template defines; parses nothing else |
| `profile/profile.ts` | a snapshot plus a style-id map, not a bag of parsed typography |
| `render/css.ts` | rewrites docx-preview's generated selectors onto `.pmd-*`; resolves no styles itself |
| `docx/rewrite.ts` | restore-or-adopt, per the save pipeline above |
| `docx/headers.ts` | snapshot carry only; synthesis and `{{token}}` substitution deleted |
| `lay.ts`, `main.ts` | the state rule above |
| `ui/settings-panel.ts` | template upload, per-document metadata, page view, open in word — everything else moves to the host settings modal |

## bundle

| | |
| --- | --- |
| today | 210 kb |
| − embedded eb garamond | −80 kb |
| + docx-preview + jszip | +175 kb |
| **projected** | **≈305 kb** against a 5 mib cap |

apache-2.0 is compatible with laymirror's polyform noncommercial
licence. `THIRD-PARTY-NOTICES.md` gains docx-preview's attribution.

## risks

| risk | mitigation |
| --- | --- |
| `lastRenderedPageBreak` is reportedly not always emitted (OfficeDev/office-js#1332) | it is an optimisation, not a dependency — path 3 is always there |
| docx-preview may not render the header's floating text box, so the rule may be missing on screen | the rule is in the *file*, so word and print are unaffected. if it is missing on screen, draw the `pBdr` in css |
| docx-preview is unverified in a real browser here — the headless probe hit `getBBox`, a jsdom gap | first implementation task is a render smoke test inside cardmirror |
| the snapshot grows plugin storage | store one snapshot per document, capped, evicted with the marker |
| a school template with no header at all | fall back to carrying nothing; do not synthesise |

## open questions

- should page view offer to save when the editor is dirty, or refuse
  and say so? refusing is simpler and never surprises
- manual page breaks still ride through cardmirror as the literal text
  `[page break]`, which is visible in work view. acceptable, or hide it
- do we keep a per-document profile, or one profile per install
