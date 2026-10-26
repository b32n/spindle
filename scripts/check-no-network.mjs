// Source guard: the libraries must not talk to the network on their own.
//   node scripts/check-no-network.mjs          # scan packages/*/src, exit 1 on a violation
//   node scripts/check-no-network.mjs --self-test
//
// Flags, in non-test source files of packages/*/src:
//   1. network APIs: fetch(), XMLHttpRequest, navigator.sendBeacon, EventSource, importScripts,
//      new WebSocket(), service workers
//   2. hard-coded http(s):// URLs to hosts that are not on the allow-list below
// A deliberate, documented exception carries `// network-allow: <reason>` on the same line
// (or the line above). The WebSocket transport package is opt-in infrastructure whose URL the
// host supplies, so it is exempt from rule 1 for `new WebSocket(`.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');

/** Hosts that may appear in source: XML namespaces, spec URLs, placeholders, and the default font mirror. */
export const ALLOWED_HOSTS = [
  'www.w3.org',
  'schemas.openxmlformats.org',
  'schemas.microsoft.com',
  'purl.org',
  'example.com',
  'example.org',
  'localhost',
  '127.0.0.1',
  'fonts.b32n.io', // default font mirror (documentation/fonts.md); opt out with fontSource {kind:'local'}
];

const API_RULES = [
  { re: /\bfetch\s*\(/, what: 'fetch()' },
  { re: /\bXMLHttpRequest\b/, what: 'XMLHttpRequest' },
  { re: /\bsendBeacon\b/, what: 'navigator.sendBeacon' },
  { re: /\bEventSource\b/, what: 'EventSource' },
  { re: /\bimportScripts\s*\(/, what: 'importScripts' },
  { re: /\bnavigator\.serviceWorker\b/, what: 'service worker' },
  { re: /\bnew\s+WebSocket\s*\(/, what: 'new WebSocket()', exemptPackages: ['transport-websocket'] },
];

const URL_RE = /https?:\/\/([A-Za-z0-9.-]+)/g;

function isTestFile(file) {
  return /\.(test|spec)\.(ts|tsx)$/.test(file) || /[\\/]__tests__[\\/]/.test(file);
}

function walk(dir, out = []) {
  for (const name of readdirSync(dir)) {
    const p = path.join(dir, name);
    if (statSync(p).isDirectory()) walk(p, out);
    else if (/\.(ts|tsx)$/.test(name) && !isTestFile(p)) out.push(p);
  }
  return out;
}

/** Violations in one file's text. `pkg` is the package folder name. */
export function scanSource(text, pkg) {
  const found = [];
  const lines = text.split('\n');
  let inBlockComment = false;
  lines.forEach((raw, i) => {
    let line = raw;
    // Skip comments: doc/prose URLs are not requests.
    if (inBlockComment) {
      if (line.includes('*/')) { inBlockComment = false; line = line.slice(line.indexOf('*/') + 2); } else return;
    }
    const trimmed = line.trim();
    if (trimmed.startsWith('//') || trimmed.startsWith('*')) return;
    if (trimmed.startsWith('/*')) {
      if (!trimmed.includes('*/')) inBlockComment = true;
      return;
    }
    const code = line.replace(/\s\/\/.*$/, (m) => (m.includes('network-allow') ? m : ''));
    const allowed = /network-allow:/.test(raw) || (i > 0 && /network-allow:/.test(lines[i - 1]));
    if (allowed) return;
    for (const rule of API_RULES) {
      if (rule.exemptPackages?.includes(pkg)) continue;
      if (rule.re.test(code)) found.push({ line: i + 1, what: rule.what, text: raw.trim() });
    }
    for (const m of code.matchAll(URL_RE)) {
      const host = m[1].replace(/\.$/, '');
      if (!ALLOWED_HOSTS.some((h) => host === h || host.endsWith(`.${h}`))) {
        found.push({ line: i + 1, what: `URL to ${host}`, text: raw.trim() });
      }
    }
  });
  return found;
}

function selfTest() {
  const cases = [
    ['const r = await fetch(url);', 'x', 1],
    ['const x = new XMLHttpRequest();', 'x', 1],
    ["const s = new WebSocket(url);", 'transport-websocket', 0],
    ["const s = new WebSocket(url);", 'sheets-core', 1],
    ["const u = 'https://cdn.evil.test/a.png';", 'x', 1],
    ["const ns = 'http://www.w3.org/2000/svg';", 'x', 0],
    ["// see https://github.com/foo/bar", 'x', 0],
    ["/* https://github.com/foo/bar */", 'x', 0],
    ["const u = 'https://ok.test/a'; // network-allow: documented opt-in", 'x', 0],
    ["// network-allow: documented\nfetch(url);", 'x', 0],
    ["placeholder=\"https://example.com/image.png\"", 'x', 0],
    ["const f = 'https://fonts.b32n.io/css2';", 'x', 0],
    ["const f = 'https://fonts.googleapis.com/css2';", 'x', 1],
  ];
  let bad = 0;
  for (const [src, pkg, expected] of cases) {
    const got = scanSource(src, pkg).length;
    if (got !== expected) { console.error(`self-test FAILED (${got} != ${expected}): ${JSON.stringify(src)} [${pkg}]`); bad++; }
  }
  if (bad) process.exit(1);
  console.log(`self-test ok (${cases.length} cases)`);
}

function main() {
  if (process.argv.includes('--self-test')) return selfTest();
  const pkgsDir = path.join(root, 'packages');
  let violations = 0;
  let files = 0;
  for (const pkg of readdirSync(pkgsDir)) {
    const src = path.join(pkgsDir, pkg, 'src');
    try { statSync(src); } catch { continue; }
    for (const file of walk(src)) {
      files++;
      for (const v of scanSource(readFileSync(file, 'utf8'), pkg)) {
        violations++;
        console.error(`${path.relative(root, file)}:${v.line}: ${v.what}\n    ${v.text}`);
      }
    }
  }
  if (violations) {
    console.error(`\n${violations} network use(s) in library source. Spindle makes no third-party requests on its own;` +
      ` add \`// network-allow: <reason>\` only for a documented, opt-in exception.`);
    process.exit(1);
  }
  console.log(`no-network check ok (${files} source files)`);
}

if (import.meta.url === `file://${process.argv[1]}`) main();
