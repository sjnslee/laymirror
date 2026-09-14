Formatting plugin for [CardMirror](https://github.com/ant981228/cardmirror).
Automatically formats output docx file based on user's uploaded custom template

- Header text can be adjusted in the plugin interface. Fonts, page setup, styles and theme come from the template
- Applied to docx file on every save

Built for the purpose of switching between different format guidelines and implementing printable functionality,
particularly with printed lay debate evidence

## Install

Desktop only.

CardMirror installs a plugin from a github release's assets, but only from
repositories on its curated list. As a non whitelisted plugin, installation
requires building and loading separately:

```
npm install
npm run build
```

in cardmirror: settings → plugins → developer → "load plugin from
file…" → `plugin.js`.

## Usage

cmd+shift+l to open interface (cmd+alt+l if CardMirror has taken the first)

Keybind can be changed through CardMirror

Also in the command palette:

- `laymirror: open`
- `laymirror: turn lay formatting on or off`
- `laymirror: apply the template now`
- `laymirror: point at the open document on disk`
- `laymirror: diagnostics`

## Templates

Guidelines for creation of templates and how they are processed by the plugin are included here: [docs/template-guidelines.md](docs/template-guidelines.md)

An example of a template for lay debate is included here: [docs/template-example.docx](docs/template-example.docx)

## Dev

```
npm test          # vitest
npm run typecheck
npm run build     # esbuild -> plugin.js
```
