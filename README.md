# Laymirror

Debate plugin for [Cardmirror](https://github.com/ant981228/cardmirror).
Automatically formats output docx file based on uploaded custom template

- file specifics e.g. header, header text, fonts, page setup, etc. can be adjusted in the plugin interface
  Applied to docx file on every save

Made for switching between different formats, particularly with printed lay debate evidence

**Status: functional prototype**

## Install

Desktop only.

Cardmirror installs a plugin from a github release's assets, but only from
repositories on its curated list. As a non whitelisted plugin, installation
requires building and loading separately:

```
npm install
npm run build
```

in cardmirror: settings → plugins → developer → "load plugin from
file…" → `plugin.js`.

## Usage

cmd+shift+l to open interface
Keybind can be changed through Cardmirror

## Templates

Guidelines for creation of templates and how they are processed by the plugin are included here: [docs/making-a-template.md](docs/making-a-template.md)

## dev

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```
