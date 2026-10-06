// Builds www/ for the Android app from ../ItemChecker.html.
// The libraries are downloaded once into www/lib so the app works without internet.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const WWW_DIR = join(__dirname, 'www');
const WWW_LIB = join(WWW_DIR, 'lib');
const SRC_HTML = join(__dirname, '..', 'ItemChecker.html');
const OUT_HTML = join(WWW_DIR, 'index.html');

const LIBS = {
  'tailwind.js': 'https://cdn.tailwindcss.com/3.4.17',
  'sql-wasm.js': 'https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/sql-wasm.js',
  'sql-wasm.wasm': 'https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/sql-wasm.wasm',
  'xlsx.full.min.js': 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
};

mkdirSync(WWW_LIB, { recursive: true });
for (const [file, url] of Object.entries(LIBS)) {
  const path = join(WWW_LIB, file);
  if (existsSync(path)) continue;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Download failed (${res.status}): ${url}`);
  writeFileSync(path, Buffer.from(await res.arrayBuffer()));
  console.log('downloaded', file);
}

let html = readFileSync(SRC_HTML, 'utf8');
const swap = (from, to) => {
  if (!html.includes(from)) throw new Error(`ItemChecker.html changed; not found: ${from}`);
  html = html.replace(from, to);
};
swap('<script src="https://cdn.tailwindcss.com"></script>', '<script src="lib/tailwind.js"></script>');
swap('<script src="https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/sql-wasm.js"></script>', '<script src="lib/sql-wasm.js"></script>');
swap('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>', '<script src="lib/xlsx.full.min.js"></script>');
swap('<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.min.js"></script>\n', '');
swap("const SQLJS_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/';", "const SQLJS_CDN = 'lib/';");
// This build IS the phone app, whether or not the Capacitor bridge has loaded yet.
swap('const PHONE_APP = !!window.Capacitor;', 'const PHONE_APP = true;');
// Android has no file type for .db, so a filtered picker greys those files out. Show every file; Import checks the type.
swap(' accept=".csv,.xlsx,.xls,.db,.sqlite,.sqlite3"', '');
swap('Check your internet connection and reload the page.', 'Close the app and open it again.');

writeFileSync(OUT_HTML, html);
console.log('www/index.html ready');
