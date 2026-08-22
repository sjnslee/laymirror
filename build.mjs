// cardmirror loads a plugin by running one file in the renderer's main
// world, so the bundle has to be a self-contained classic script.
import { build } from 'esbuild';

const watch = process.argv.includes('--watch');

const options = {
  entryPoints: ['src/main.ts'],
  outfile: 'plugin.js',
  bundle: true,
  format: 'iife',
  platform: 'browser',
  target: 'chrome120', // electron 42
  sourcemap: false,
  // docx-preview is apache-2.0 and its licence requires the notice to travel
  // with the code, so licence comments stay in the bundle
  legalComments: 'inline',
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
