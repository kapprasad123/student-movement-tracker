/**
 * Student Movement Tracker — Local Development Server (server.js)
 *
 * Serves the four dashboards locally using Node.js built-in modules:
 * - http://localhost:3000/index.html (Teacher Dashboard)
 * - http://localhost:3000/principal-dashboard.html (Principal Dashboard)
 * - http://localhost:3000/leading-teacher-dashboard.html (Leading Teacher Dashboard)
 * - http://localhost:3000/admin-dashboard.html (Admin Dashboard)
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
const ROOT_DIR = __dirname;

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon'
};

function createServer() {
  return http.createServer((req, res) => {
    const parsedUrl = url.parse(req.url, true);
    let pathname = decodeURIComponent(parsedUrl.pathname || '/');
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
    console.log('============================================================');
    console.log(`Teacher Dashboard:         http://localhost:${PORT}/index.html`);
    console.log(`Principal Dashboard:       http://localhost:${PORT}/principal-dashboard.html`);
    console.log(`Leading Teacher Dashboard: http://localhost:${PORT}/leading-teacher-dashboard.html`);
    console.log(`Admin Dashboard:           http://localhost:${PORT}/admin-dashboard.html`);
    console.log('============================================================');
  });
}

module.exports = { createServer };
