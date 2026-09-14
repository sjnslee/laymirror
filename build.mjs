// cardmirror loads a plugin by running one file in the renderer's main
// world, so the bundle has to be a self-contained classic script.
import { build } from 'esbuild';
import { readFileSync } from 'node:fs';

const watch = process.argv.includes('--watch');

// fflate ships no `/*!` banner, so `legalComments: 'inline'` finds nothing to
// keep and the bundle would carry no notice at all. MIT wants one alongside
// every copy, and the bundle is the only file an installing user receives.
const notice = [
  '/*!',
  ' * laymirror bundles fflate (https://github.com/101arrowz/fflate).',
  ...readFileSync('node_modules/fflate/LICENSE', 'utf8')
    .trimEnd()
    .split('\n')
    .map((line) => ` * ${line}`.trimEnd()),
  ' */',
].join('\n');

const options = {
  entryPoints: ['src/main.ts'],
  outfile: 'plugin.js',
  bundle: true,
  format: 'iife',
  platform: 'browser',
  target: 'chrome120', // electron 42
  sourcemap: false,
  legalComments: 'inline',
  banner: { js: notice },
};

if (watch) {
  const ctx = await (await import('esbuild')).context(options);
  await ctx.watch();
  console.log('watching');
} else {
  const result = await build(options);
  if (result.errors.length) process.exit(1);
  console.log('built plugin.js');
}
