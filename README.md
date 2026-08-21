# laymirror

lay debate documents for [cardmirror](https://github.com/ant981228/cardmirror).

cardmirror is built for tech debate: verbatim formatting, dense evidence,
no pagination. lay debate is the opposite — the document gets printed and
handed to a parent judge, so it needs readable serif type, real pages, a
header with the team code and title, a footer with page numbers, and it
has to open in word looking like the school's own template.

**status: in progress.** the marker, work-view typography, template
ingest and the styles/section writers are done. the save pipeline,
pagination, page view and print are not.

## how it works

a lay document is marked by a custom document property that travels
inside the `.docx`, so a teammate who opens the file gets the lay
formatting without configuring anything. when the marker is absent the
plugin does nothing at all — no stylesheet, no watching, no rewriting.

**a school's template docx is the profile.** laymirror defines no format
of its own: point it at a template and it reads the fonts, sizes,
weights, indents, spacing, page setup, header and footer straight out of
`styles.xml` and `sectPr`. the built-in profile is deliberately generic
and is not any school's.

## install

desktop only — plugins run in cardmirror's renderer, which the web and
lite editions don't have.

```
npm install
npm run build
```

then in cardmirror: settings → plugins → developer → "load plugin from
file…" → pick `plugin.js`.

commands are invoked by keyboard shortcut (settings → keyboard
shortcuts → plugins), or by putting one on the ribbon with "+ add
button". **laymirror: open** is the one worth a ribbon slot — everything
else lives in that panel.

## development

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```

`src/host/` is a quarantine: every assumption about cardmirror's
internals lives there, stamped with the version it was verified against
(1.3.0), so an upgrade breaks a test instead of breaking the plugin
mid-round.

tests run against a synthetic donor template built in `tests/fixture.ts`.
no real school template is in the repo. drop one at `local/donor.docx`
(gitignored) and the optional checks in `tests/donor.test.ts` will run
against it too.

## license

[PolyForm Noncommercial 1.0.0](LICENSE), matching cardmirror's own
license. parts of `src/profile/mapping.ts` are derived from cardmirror's
published source in order to interoperate with it. not affiliated with
or endorsed by the cardmirror project.
