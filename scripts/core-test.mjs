import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const required = [
  'index.html',
  'package.json',
  'src/main.jsx',
  'src/styles.css',
  'worker.js',
  'wrangler.toml'
];
let failed = false;
for (const file of required) {
  if (!fs.existsSync(path.join(root, file))) {
    console.error(`FAIL: missing ${file}`);
    failed = true;
  } else {
    console.log(`PASS: ${file}`);
  }
}

const main = fs.readFileSync(path.join(root, 'src/main.jsx'), 'utf8');
const worker = fs.readFileSync(path.join(root, 'worker.js'), 'utf8');

const checks = [
  ['frontend API uses same-origin proxy', /const API_URL\s*=\s*["']\/api\/office-task["']/.test(main)],
  ['worker exposes office-task POST route', /url\.pathname\s*===\s*["']\/api\/office-task["']/.test(worker)],
  ['worker exposes health route', /url\.pathname\s*===\s*["']\/api\/office-task\/health["']/.test(worker)],
  ['worker uses POST upstream', /method:\s*["']POST["']/.test(worker)],
  ['worker returns JSON', /Content-Type["']\s*:\s*["']application\/json/.test(worker)],
];
for (const [name, ok] of checks) {
  console.log(`${ok ? 'PASS' : 'FAIL'}: ${name}`);
  if (!ok) failed = true;
}

if (failed) process.exit(1);
console.log('ALL CORE TESTS PASSED');
