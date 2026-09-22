Formatting plugin for [CardMirror](https://github.com/ant981228/cardmirror).
Automatically formats output docx file based on an uploaded custom template

- Header text can be adjusted in the plugin interface, and fonts, page setup, styles and theme come from the template
- Applied to docx file on save

Made for switching between different format guidelines and implementing printable functionality,
particularly with printed lay debate evidence

## Install

As a non whitelisted plugin, installation needs building and loading separately:

```
npm install
npm run build
```

in cardmirror: settings -> plugins -> developer -> "load plugin from
file…" -> `plugin.js`.

## Usage

cmd+shift+l opens interface (alternatively cmd+alt+l)

Other features are similarly integrated into the command palette without a set keybind:

- `laymirror: open`
- `laymirror: turn lay formatting on or off`
- `laymirror: apply the template now`
- `laymirror: point at the open document on disk`
- `laymirror: diagnostics`

## Templates

Guidelines for creating templates and how they are processed by the plugin are included here: [docs/template-guidelines.md](docs/template-guidelines.md)

An example of a template for lay debate is included here: [docs/template-example.docx](docs/template-example.docx)

## Dev

```
npm test
npm run typecheck
npm run build
```
