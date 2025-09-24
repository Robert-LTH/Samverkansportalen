#!/usr/bin/env node
import { context, build as esbuild } from 'esbuild';
import { rm } from 'fs/promises';
import path from 'path';

const args = new Set(process.argv.slice(2));
const watch = args.has('--watch') || args.has('-w');
const serve = args.has('--serve');
const minify = args.has('--minify') || process.env.NODE_ENV === 'production';
const outdir = path.resolve(process.cwd(), 'dist');

const sharedOptions = {
  entryPoints: [path.resolve(process.cwd(), 'src/index.ts')],
  bundle: true,
  sourcemap: true,
  format: 'esm',
  platform: 'browser',
  target: ['es2017'],
  splitting: true,
  outdir,
  metafile: true,
  logLevel: 'info',
  minify,
  define: {
    'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV ?? (minify ? 'production' : 'development'))
  },
  external: [
    '@microsoft/sp-core-library',
    '@microsoft/sp-lodash-subset',
    '@microsoft/sp-office-ui-fabric-core',
    '@microsoft/sp-property-pane',
    '@microsoft/sp-webpart-base',
    '@pnp/sp',
    'office-ui-fabric-react',
    'office-ui-fabric-react/lib/Styling',
    'react',
    'react-dom',
    'ImprovementPortalWebPartStrings'
  ]
};

async function ensureCleanOutput() {
  await rm(outdir, { recursive: true, force: true });
}

async function run() {
  await ensureCleanOutput();

  if (watch || serve) {
    const ctx = await context(sharedOptions);
    await ctx.watch();

    console.log('Watching for changes...');

    if (serve) {
      const port = process.env.PORT ? Number(process.env.PORT) : 4321;
      const host = process.env.HOST ?? '0.0.0.0';
      const { host: serverHost, port: serverPort } = await ctx.serve({
        servedir: outdir,
        host,
        port
      });
      console.log(`Serving build output at http://${serverHost}:${serverPort}`);
      await new Promise(() => {});
    }
  } else {
    const result = await esbuild(sharedOptions);
    const warnings = result.warnings ?? [];
    if (warnings.length > 0) {
      console.warn(`Build completed with ${warnings.length} warning${warnings.length === 1 ? '' : 's'}.`);
    }
  }
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
