# laymirror

lay debate plugin for [cardmirror](https://github.com/ant981228/cardmirror).

**status: prototype**


## install

desktop only.

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
