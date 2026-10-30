/**
 * Student Movement Tracker — Syntax & Structural Verification Script (scripts/syntax-check.js)
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');

const JS_AND_GS_FILES = [
  'google-apps-script/Code.gs',
  'js/config.js',
  'js/utils.js',
  'js/api.js',
  'js/teacher-dashboard.js',
  'js/principal-dashboard.js',
  'js/leading-teacher-dashboard.js',
  'js/admin-dashboard.js',
  'server.js',
  'scripts/seed-test-spreadsheet.js',
  'scripts/create-test-sheet.js',
  'tests/gas-mock.js',
  'tests/run-tests.js'
];

const HTML_FILES = [
  'index.html',
  'principal-dashboard.html',
  'leading-teacher-dashboard.html',
  'admin-dashboard.html'
];

let passed = 0;
let failed = 0;

console.log('============================================================');
console.log('Running Syntax & Structural Verification Checks');
console.log('============================================================');

for (const relPath of JS_AND_GS_FILES) {
  const absPath = path.join(ROOT, relPath);
  try {
    if (!fs.existsSync(absPath)) {
      throw new Error('File does not exist');
    }
    const source = fs.readFileSync(absPath, 'utf8');
    new vm.Script(source, { filename: relPath });
    console.log(`[PASS] Syntax OK: ${relPath}`);
    passed++;
  } catch (err) {
    console.error(`[FAIL] Syntax Error in ${relPath}: ${err.message}`);
    failed++;
  }
}

for (const htmlRel of HTML_FILES) {
  const absPath = path.join(ROOT, htmlRel);
  try {
    const content = fs.readFileSync(absPath, 'utf8');
    if (!content.includes('<!DOCTYPE html>') || !content.includes('</html>')) {
      throw new Error('Missing DOCTYPE or closing html tag');
    }
    if (!content.includes('js/config.js') || !content.includes('js/utils.js') || !content.includes('js/api.js')) {
      throw new Error('Missing required shared JS script includes');
    }
    console.log(`[PASS] HTML Structure OK: ${htmlRel}`);
    passed++;
  } catch (err) {
    console.error(`[FAIL] HTML Check Failed in ${htmlRel}: ${err.message}`);
    failed++;
  }
}

console.log('------------------------------------------------------------');
console.log(`Syntax & Structural Summary: ${passed} passed, ${failed} failed`);
console.log('============================================================');

if (failed > 0) {
  process.exit(1);
}
