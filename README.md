# laymirror

lay debate plugin for [cardmirror](https://github.com/ant981228/cardmirror).

cardmirror's exporter rebuilds the `.docx` from scratch on every save and keeps
no header, footer or theme. laymirror puts the school's template back on —
styles, fonts, page setup and header — every time it runs.

**status: prototype**


## install

desktop only.

### from a github release

this is the one that survives a restart.

1. publish a github **release** with `cardmirror-plugin.json` and `plugin.js`
   attached as assets. cardmirror reads the *latest* release and nothing else —
   files sitting in the repo are never fetched.
2. cardmirror's installer only accepts repos on a curated list. open the dev
   console (command palette → "open dev console") and unlock it:
   ```js
   __plugins('community-on')
   ```
3. settings → plugins → install from github → `<owner>/laymirror` (or the
   full github url).
4. **turn the row's toggle on.** an installed plugin is disabled until you do,
   and a disabled plugin is silent rather than broken-looking.

after that it loads on every launch. cardmirror checks the repo's latest
release for updates, so shipping a new version is a release with a bumped
`version` in the manifest.

### from a local build

```
npm install
npm run build
```

settings → plugins → developer → "load plugin from file…" → `plugin.js`.
this one is session-only: cardmirror forgets it on restart.


## use

⌘⇧L opens the panel.

1. **load…** — pick the school's template (`.docx`, `.docm`, `.dotx`, `.dotm`)
2. **turn on** — this document is now a lay file
3. type into the header fields
4. save

every save rewrites the file on disk with the template's styles, fonts, theme,
numbering, page size, margins and header, and the header text you typed.

the fields laymirror offers are the ones the template marks. wrap a placeholder
in a zero-width space (`U+200B`) in word and it becomes editable; everything
outside the marks is the school's and is left alone. a template that marks
nothing falls back to offering each stretch of plain text between tabs and page
fields.

cardmirror keeps showing its own formatting — it has no header and no page
view. open the file in word to see the result.


## dev

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```

`docs/design.md` is why it works the way it does.
