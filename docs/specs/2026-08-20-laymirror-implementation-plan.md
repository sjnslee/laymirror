# laymirror — implementation plan

companion to `2026-08-20-laymirror-design.md`. that document says what
and why; this one says how.

## shape of the thing

one esbuild bundle, `plugin.js`, plus `cardmirror-plugin.json`.
typescript source, vitest tests, no framework. the plugin registers
commands and settings through cardmirror's public api and does
everything else through renderer access that is unofficial by
necessity — quarantined in `src/host/`.

```
open lay docx
   ↓ marker in docProps/custom.xml
detect ─→ load profile ─→ inject stylesheet ─→ work view
   ↓                                              ↓ toggle
watch statFile                                paginate ─→ page view ─→ print
   ↓ mtime bump
read ─→ unzip ─→ rewrite ─→ zip ─→ write
```

two rules the architecture exists to enforce:

1. **one profile, two renderers.** `toCss` and `toOoxml` read the same
   object. if they can drift, the screen lies about the file.
2. **all cardmirror internals in one directory.** every css variable,
   dom class, localstorage key, and undocumented reach lives in
   `src/host/`, stamped with the version it was verified against.

---

## modules

### `src/profile/`

**`profile.ts`** — types, no logic.

```ts
export type BlockType =
  | 'pocket' | 'hat' | 'block' | 'tag' | 'analytic'
  | 'undertag' | 'cite_paragraph' | 'card_body' | 'paragraph';

export type RunType =
  | 'underline_mark' | 'emphasis_mark' | 'cite_mark'
  | 'analytic_mark' | 'undertag_mark';

export interface TypeSpec {
  /** styleId written into document.xml. */
  styleId: string;
  /** w:name — what cardmirror's legacy remapper reads back on import.
   *  changing this silently breaks round-trip; see mapping.ts. */
  styleName: string;
  font?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: 'none' | 'single' | 'thick' | 'double';
  smallCaps?: boolean;
  color?: string;              // hex6, no '#'
  align?: 'left' | 'center' | 'right' | 'justify';
  indentLeftDxa?: number;
  indentRightDxa?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacing?: { rule: 'auto' | 'exact' | 'atLeast'; value: number };
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  outlineLevel?: number | null;
}

export interface PageSetup {
  widthTwips: number;
  heightTwips: number;
  margin: { top: number; right: number; bottom: number; left: number;
            header: number; footer: number };
}

export interface Profile {
  id: string;
  name: string;
  types: Record<BlockType | RunType, TypeSpec>;
  page: PageSetup;
  /** raw header1/footer1 xml from the donor, tokens unresolved. */
  headerXml: string | null;
  footerXml: string | null;
  /** basename only, e.g. "Lay Cut Cards.dotx". never an absolute path. */
  attachedTemplate: string | null;
  /** the donor's styles.xml, used as the base we merge into. */
  donorStylesXml: string;
  fontFallbacks: Record<string, string>;
}
```

**`read-template.ts`** — `readTemplate(bytes): Promise<Profile>`.
unzips a donor `.docx`/`.dotx`, parses `styles.xml` into `TypeSpec`s
(resolving `basedOn` chains), reads `sectPr` for `PageSetup`, lifts
`header1.xml`/`footer1.xml` verbatim, reads `attachedTemplate` and
**reduces it to a basename** — donors carry an absolute
path with a real person's home directory in it, and word only
basename-matches.

**`mapping.ts`** — the cardmirror↔template vocabulary bridge, and the
one place that knows round-trip safety:

```ts
/** style names cardmirror's legacy remapper recognizes on import.
 *  a mapping whose styleName is outside this set exports fine but
 *  reimports as a plain paragraph. validated, surfaced in the ui. */
export const ROUND_TRIP_SAFE_NAMES: Record<BlockType, readonly string[]>;
export function validateMapping(p: Profile): MappingWarning[];
```

**`defaults.ts`** — a deliberately generic lay profile, no school'"'"'s, 
inlined so the plugin works before any template is chosen.

### `src/render/`

**`css.ts`** — `toCss(profile): string`. emits one stylesheet scoped to
the editor selector, every declaration `!important` so it survives
cardmirror rewriting its inline custom properties when the user touches
appearance settings. sets `--pmd-size-*` where a variable exists and
writes direct rules where one doesn't (per-type `font-family`, which
cardmirror has no variable for).

**`paginate.ts`** — the measuring engine. pure, dom-in/boxes-out:

```ts
export interface PageBox { blocks: HTMLElement[]; heightPx: number; }
export interface PaginateResult { pages: PageBox[]; breaks: BreakMark[]; }
export function paginate(
  blocks: readonly HTMLElement[],
  page: PageSetup,
  opts: { widowLines: number; orphanLines: number },
): PaginateResult;
```

honors `pageBreakBefore`, `keepNext`, `keepLines`, widow/orphan counts,
and manual breaks. written from the ooxml spec and word's documented
pagination behavior — **not** derived from any proprietary engine.

**`page-view.ts`** — read-only paged surface. clones the editor dom
(cheap: `cloneNode` of already-rendered output), runs `paginate`, wraps
each `PageBox` in a page element with header and footer chrome, fills
`PAGE`/`NUMPAGES` values. mounts as an overlay above the editor.
must neutralize `#editor { zoom: var(--editor-zoom) }` before measuring
or every box is off by the zoom factor.

**`draft-marks.ts`** — work view's page-break indicators, drawn from the
same `paginate` output as widgets between blocks. word's dotted rule,
labelled for manual breaks.

**`print.ts`** — `@page` sized to `PageSetup` with zero margin, page
boxes already laid out, so the browser re-paginates nothing.

### `src/docx/`

**`zip.ts`** — thin `fflate` wrapper. `unzip(bytes): Parts`,
`zip(parts): Uint8Array`. cardmirror uses fflate too; we bundle our own
rather than reach for theirs.

**`styles.ts`** — `buildStylesXml(profile): string`. starts from
`donorStylesXml` so anything we don't model survives, then overwrites
the mapped styles from their `TypeSpec`s.

**`headers.ts`** — `buildHeaderFooter(profile, meta)`. token-substitutes
title / authors / team code into the donor's header xml, and returns the
parts plus the `document.xml.rels` relationships and
`[Content_Types].xml` overrides they need.

**`sect.ts`** — `buildSectPr(profile, rels)`. page size, margins,
header/footer references.

**`marker.ts`** — read and write `layMirrorProfile` in
`docProps/custom.xml`, merging beside cardmirror's `cmirDocId` rather
than replacing the part (cardmirror's own `writeDocId` shows the shape).

**`rewrite.ts`** — the pipeline, and the highest-value module:

```ts
export interface DocMeta { title: string; authors: string; teamCode: string; }

export async function rewriteDocx(
  bytes: Uint8Array, profile: Profile, meta: DocMeta,
): Promise<Uint8Array>;
```

in order: verify it is a real zip with `word/document.xml` (a partial
read mid-save must abort, not corrupt) → replace `styles.xml` →
rewrite `document.xml` style references and **add `pStyle` to card
bodies and cite paragraphs**, which cardmirror emits bare → restore
manual page breaks from markers → add header/footer parts, rels,
content-types → replace `sectPr` → point `attachedTemplate` at the lay
`.dotx` → stamp the marker.

### `src/host/` — the quarantine

**`cardmirror.ts`** — every assumption, in one file, each with the
version it was verified against (1.3.0):

```ts
export const VERIFIED_AGAINST = '1.3.0';
export const EDITOR_SELECTOR = '#editor, .pmd-pane-editor';
export const CLASS = {
  tag: 'pmd-tag', cardBody: 'pmd-card-body', citePara: 'pmd-cite-para',
  undertag: 'pmd-undertag', analytic: 'pmd-analytic', card: 'pmd-card',
} as const;
export const LS = { recents: 'pmd-recent-files' } as const;

/** the open document as prosemirror json, via the editor dom's
 *  `pmViewDesc` back-reference. undocumented; verified against
 *  prosemirror-view 1.37 (`viewdesc.ts` sets `dom.pmViewDesc = this`).
 *  null when the shape is not what we expect — never throws. */
export function readDocNode(): unknown | null;
```

everything else imports from here. nothing else touches a cardmirror
internal.

**`paths.ts`** — resolve the focused document's absolute path.
`api.docInfo()` gives `docId` + filename; `pmd-recent-files` gives
`{handle, filename, format}`. filename matching is ambiguous when two
open docs share a name, so confirm by reading the candidate's
`cmirDocId` from `docProps/custom.xml` and comparing. cached.

**`watcher.ts`** — save detection.

```ts
export interface Watcher { start(path: string): void; stop(): void; }
export function watchSaves(onSaved: (path: string) => void): Watcher;
```

polls `electronAPI.statFile`. backs off when the window is unfocused.
serializes with its own writes so the rewrite's `writeFileAtPath`
doesn't retrigger itself.

### `src/ui/`

**`settings-panel.ts`** — own modal, because cardmirror's declared
settings render only boolean/text/number/select and cannot express a
mapping table. sections: template picker, mapping table with live
preview, per-type overrides, header/footer metadata, fonts.

**`fonts.ts`** — metric-probe detection in the style of cardmirror's
`font-detect.ts`, plus per-family fallback selection.

### `src/main.ts`

registration, commands, and the state machine over `off` /
`lay+work` / `lay+page`.

commands: toggle lay, toggle page view, open settings, apply formatting
now, print, insert page break.

---

## data flow

**open.** cardmirror opens a docx. we notice a new focused doc by
polling `api.docInfo()` — there are no lifecycle events — resolve its
path, read `layMirrorProfile`. absent ⇒ stay off entirely. present ⇒
load the profile, recover header metadata by parsing the file's existing
`header1.xml`, inject the stylesheet.

**save.** user hits ctrl+s. cardmirror writes its own docx. watcher sees
the mtime bump, reads, rewrites, writes back. the write is flagged so
the watcher ignores its own mtime bump.

**page view.** clone editor dom → neutralize zoom → `paginate` → wrap in
page boxes with chrome → overlay. work view gets break widgets from the
same result.

**print.** print the page view.

---

## testing

vitest + jsdom for logic, a real headless browser only if measurement
tests need it.

**fixtures.** a synthetic donor is the primary fixture. add a
synthetic cardmirror export (produced once by the real exporter, checked
in) as the rewrite input.

**unit.** `readTemplate` against the donor: asserts palatino 10pt
Normal, times new roman 20pt bold smallcaps centered `Heading1` with
`pageBreakBefore`, `Tag` bold, `Cite` thick-underlined, `card` indented
288 dxa. snapshot `toOoxml` and `toCss`.

**round-trip — the test that matters most.** take a cardmirror export,
`rewriteDocx` it, then feed the result back through **cardmirror's real
importer**, imported from a pinned git checkout as a dev dependency.
assert tags come back as `tag`, cites as `cite_paragraph`, card bodies
as `card_body`. this is the claim the whole design rests on — that
writing genuine lay style names stays round-trip safe because the legacy
remapper recognizes them. it should be a test, not a belief.

**paginator.** fixture documents with hand-computed break points.
property test: every block appears exactly once across pages, in order —
no loss, no duplication. explicit cases for `pageBreakBefore`, `keepNext`
runs, and widow/orphan.

**canary.** pin a cardmirror checkout and assert every constant in
`host/cardmirror.ts` still resolves against it — the css variables
exist in `style.css`, the classes are emitted, the localstorage keys are
still written, `dom.pmViewDesc` is still set by the pinned
prosemirror-view. a cardmirror upgrade then breaks a test instead of
breaking the plugin in a round.

**not automated.** the electron integration — `statFile`,
`readFileAtPath`, `writeFileAtPath` from a loaded plugin — cannot be
tested without driving electron. that is what phase 0 is for, done by
hand, once.

---

## build order

**phase 0 — spike.** dev-load a stub `plugin.js` that calls `statFile`,
`readFileAtPath`, and `writeFileAtPath`, rewrites one byte of a docx,
and confirms word still opens it. everything downstream assumes these
three calls work from renderer plugin code, and that was established by
reading cardmirror's source, not by running it. an hour, before
anything else. if it fails the whole save pipeline needs redesigning and
better to know now.

1. **skeleton** — registers, commands appear, `layMirrorProfile`
   round-trips through `docProps/custom.xml`. off-state is genuinely
   inert.
2. **work view** — `toCss` from the built-in generic defaults.
3. **template ingest** — `readTemplate`, `mapping`, settings panel.
4. **save pipeline** — watcher plus `rewriteDocx`. the deliverable:
   ctrl+s yields a docx word opens as a lay document.
5. **pagination** — `paginate`, page view, draft break marks, manual
   page breaks. one piece of work; they share the computation.
6. **print.**
7. **fonts** — detection and substitution.

phase 4 is the one that has to work. 1–3 exist to make it testable.

---

## risks

| risk | severity | mitigation |
| --- | --- | --- |
| phase 0 fails — the electron calls aren't reachable from a plugin | fatal to the design | spike first. fallback is an explicit "export lay docx" command using `saveAs`, which loses transparent ctrl+s |
| plugin api v2 sandboxes plugins and closes renderer access | high, not imminent — v2 is signalled, not shipped | `src/host/` quarantine keeps it a contained fix. open an upstream issue asking for save and style hooks |
| watcher reads a partially-written file | medium | validate zip integrity before rewriting; retry once; never write back on a failed parse |
| watcher retriggers on its own write | medium | flag self-writes by expected mtime/size |
| pagination drifts from word | medium, bounded | `PAGE`/`NUMPAGES` are live fields, so the printed file is right regardless. drift is confined to preview |
| missing palatino / garamond widens drift | medium | detect, warn by name, offer per-family fallback |
| cardmirror destroys manual page breaks | medium | we carry them through the rewrite ourselves |
| a mapping's `styleName` falls outside the legacy vocabulary and silently breaks reimport | medium | `validateMapping` + the round-trip test |
| distribution — installer allowlist holds one repo, 5 MiB asset cap | low | dev-load has no cap and works today; ask the maintainer for a listing when sharing matters |
| licensing | low, if respected | no superdoc engine, no inspection of it, no reimplementation of it. our paginator comes from the ooxml spec and word's documented behavior |

## open questions

carried from the design doc, none blocking:

- tables in v1, or inherit the donor's `TableNormal`?
- page view in one pane while another edits, or one per window?
- two lay profiles active at once, for two schools' files open together?
