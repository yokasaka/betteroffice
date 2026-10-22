import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { $ } from 'bun';

const [version, out] = process.argv.slice(2);
if (!version || !out) throw new Error('usage: bun dotta/scripts/pack-npm.ts <version> <out dir>');

const packages = ['packages/pptx-i18n', 'packages/pptx', 'packages/pptx-react'];
const destination = resolve(out);
await mkdir(destination, { recursive: true });

for (const dir of packages) {
  const path = `${dir}/package.json`;
  const manifest = JSON.parse(await readFile(path, 'utf8'));
  manifest.version = version;
  for (const name of Object.keys(manifest.dependencies ?? {})) {
    if (name.startsWith('@betteroffice/pptx')) manifest.dependencies[name] = version;
  }
  await writeFile(path, `${JSON.stringify(manifest, null, 2)}\n`);
}

await $`bun run build:pptx`;

for (const dir of packages) {
  await $`bun pm pack --destination ${destination}`.cwd(dir);
}
