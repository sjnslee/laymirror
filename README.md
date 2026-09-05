# laymirror

lay debate plugin for [cardmirror](https://github.com/ant981228/cardmirror).

**status: prototype**


## install

desktop only.

cardmirror installs a plugin from a github release's assets, and only from
repositories on its curated list — which this one is not on yet. until then,
build it and load it by hand:

```
npm install
npm run build
```

in cardmirror: settings → plugins → developer → "load plugin from
file…" → `plugin.js`.


## dev

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```
