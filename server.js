/**
 * Student Movement Tracker — Local Development & Test Server (server.js)
 *
 * Responsibilities:
 * - Serves all four web dashboards locally:
 *   * http://localhost:3000/index.html (Teacher Dashboard)
 *   * http://localhost:3000/principal-dashboard.html (Principal Dashboard)
 *   * http://localhost:3000/leading-teacher-dashboard.html (Leading Teacher Dashboard)
 *   * http://localhost:3000/admin-dashboard.html (Admin Dashboard)
 * - Hosts an isolated local mock backend endpoint (/api and /exec) backed by
 *   a dedicated, separate test spreadsheet (tests/gas-mock.js + google-apps-script/Code.gs)
 *   populated with realistic test data without touching any real school data.
 * - Exports standalone test spreadsheet CSV files to test-spreadsheet/
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { createGasEnvironment } = require('./tests/gas-mock.js');
const Code = require('./google-apps-script/Code.gs');
const { seedEnvironment, exportCsvFiles, TEST_CREDENTIALS } = require('./scripts/seed-test-spreadsheet.js');

const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
const ROOT_DIR = __dirname;
const TEST_SHEET_DIR = path.join(ROOT_DIR, 'test-spreadsheet');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.csv': 'text/csv; charset=utf-8'
};

function createServer(options = {}) {
  // Initialize isolated test spreadsheet runtime
  const gasEnv = options.gasEnv || createGasEnvironment();
  seedEnvironment(Code, gasEnv);

  try {
    exportCsvFiles(gasEnv, TEST_SHEET_DIR);
  } catch (e) {
    // Non-fatal if folder creation restricted
  }

  return http.createServer((req, res) => {
    const reqUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    let pathname = decodeURIComponent(reqUrl.pathname || '/');

    // Handle Local Mock Apps Script API (/api or /exec)
    if (pathname === '/api' || pathname === '/exec') {
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');

      if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
      }

      const params = Object.fromEntries(reqUrl.searchParams.entries());

      if (req.method === 'GET') {
        try {
          const out = Code.doGet({ parameter: params });
          const raw = out.getContent();
          res.writeHead(200, {
            'Content-Type': 'application/json; charset=utf-8',
            'X-Content-Type-Options': 'nosniff'
          });
          res.end(raw);
        } catch (err) {
          res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ success: false, error: err.message, code: 'INTERNAL_ERROR' }));
        }
        return;
      }

      if (req.method === 'POST') {
        let body = '';
        req.on('data', (chunk) => { body += chunk; });
        req.on('end', () => {
          try {
            const out = Code.doPost({
              parameter: params,
              postData: { contents: body }
            });
            const raw = out.getContent();
            res.writeHead(200, {
              'Content-Type': 'application/json; charset=utf-8',
              'X-Content-Type-Options': 'nosniff'
            });
            res.end(raw);
          } catch (err) {
            res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ success: false, error: err.message, code: 'INTERNAL_ERROR' }));
          }
        });
        return;
      }
    }

    if (pathname === '/') {
      pathname = '/index.html';
    }

    const safePath = path.normalize(pathname).replace(/^(\.\.[/\\])+/, '');
    const fullPath = path.join(ROOT_DIR, safePath);

    if (!fullPath.startsWith(ROOT_DIR)) {
      res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('403 Forbidden');
      return;
    }

    fs.stat(fullPath, (err, stats) => {
      if (err || !stats.isFile()) {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('404 Not Found: ' + pathname);
        return;
      }

      const ext = path.extname(fullPath).toLowerCase();
      const contentType = MIME_TYPES[ext] || 'application/octet-stream';
      res.writeHead(200, {
        'Content-Type': contentType,
        'X-Content-Type-Options': 'nosniff'
      });
      fs.createReadStream(fullPath).pipe(res);
    });
  });
}

if (require.main === module) {
  const server = createServer();
  server.listen(PORT, () => {
    console.log('============================================================');
    console.log('Student Movement Tracker — Local Development Server');
    console.log('Connected to: Isolated Test Spreadsheet (Zero Real Data Access)');
    console.log('============================================================');
    console.log(`Teacher Dashboard:         http://localhost:${PORT}/index.html`);
    console.log(`Principal Dashboard:       http://localhost:${PORT}/principal-dashboard.html`);
    console.log(`Leading Teacher Dashboard: http://localhost:${PORT}/leading-teacher-dashboard.html`);
    console.log(`Admin Dashboard:           http://localhost:${PORT}/admin-dashboard.html`);
    console.log(`Test Backend API:          http://localhost:${PORT}/api`);
    console.log('============================================================');
    console.log('Test Passcodes:');
    console.log(`  Admin Passcode:          ${TEST_CREDENTIALS.adminPasscode}`);
    console.log(`  Principal Passcode:      ${TEST_CREDENTIALS.principalPasscode}`);
    console.log(`  Leading Teacher Passcode:${TEST_CREDENTIALS.leadingTeacherPasscode}`);
    console.log(`  Teacher App PIN:         ${TEST_CREDENTIALS.teacherPin}`);
    console.log('============================================================');
  });
}

module.exports = { createServer, TEST_CREDENTIALS };
