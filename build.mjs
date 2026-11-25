import { mkdir, readFile, writeFile } from 'node:fs/promises';

const files = {
  '/': ['index.html', 'text/html; charset=utf-8'],
  '/style.css': ['style.css', 'text/css; charset=utf-8'],
  '/app.js': ['app.js', 'text/javascript; charset=utf-8'],
};
const assets = Object.fromEntries(await Promise.all(Object.entries(files).map(async ([url, [name, type]]) =>
  [url, {type, text: await readFile(new URL('./web/' + name, import.meta.url), 'utf8')}]
)));
const entry = await readFile(new URL('./worker/index.mjs', import.meta.url), 'utf8');
await mkdir(new URL('./dist/server/', import.meta.url), {recursive: true});
await writeFile(new URL('./dist/server/index.js', import.meta.url), 'const assets = ' + JSON.stringify(assets) + ';\n' + entry);
