import { build } from 'esbuild';
import { copyFile } from 'node:fs/promises';

await build({
  entryPoints: ['scripts/certificate-source.cjs'],
  outfile: 'portable/certificate.cjs',
  bundle: true, platform: 'node', target: 'node22', minify: true, legalComments: 'inline'
});
await copyFile('node_modules/node-forge/LICENSE', 'portable/forge-LICENSE.txt');
