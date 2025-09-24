#!/usr/bin/env node
import { rm } from 'fs/promises';
import path from 'path';

const folders = ['dist', 'lib', 'temp'];

async function removeFolder(folder) {
  const target = path.resolve(process.cwd(), folder);
  await rm(target, { recursive: true, force: true });
  console.log(`Removed ${folder}`);
}

async function run() {
  await Promise.all(folders.map(removeFolder));
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
