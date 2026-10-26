// Local dev server -- added so this project can be exercised locally without
// `netlify dev` (needed to test report-hub calling this app's functions
// directly). Static root is "." (per netlify.toml's publish = "."), not a
// public/ subfolder, unlike the sibling report apps. Usage: node server.js

const http = require('http');
const fs = require('fs');
const path = require('path');

const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  fs.readFileSync(envPath, 'utf8').split('\n').forEach(line => {
    const [k, ...v] = line.split('=');
    if (k && v.length && !k.trim().startsWith('#')) process.env[k.trim()] = v.join('=').trim();
  });
}

const PORT = process.env.PORT || 8791;
// Loopback only by default: this server proxies the Netlify functions with live
// credentials and has no auth, so it must not be reachable from the LAN.
// Set HOST=0.0.0.0 only if you deliberately want that.
const HOST = process.env.HOST || '127.0.0.1';
// Reject non-loopback Host headers (blocks DNS-rebinding: a malicious page
// resolving its own hostname to 127.0.0.1 to read our responses).
const LOOPBACK_HOST = /^(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$/i;
const ROOT = path.join(__dirname, 'public');   // mirrors netlify.toml: only public/ is served

const MIME = {
  '.html': 'text/html', '.css': 'text/css', '.js': 'application/javascript',
  '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.b64': 'text/plain',
};

const server = http.createServer(async (req, res) => {
  if (HOST === '127.0.0.1' && !LOOPBACK_HOST.test(req.headers.host || '')) {
    res.writeHead(403); res.end('Forbidden: unexpected Host header'); return;
  }
  let parsed;
  try { parsed = new URL(req.url, 'http://localhost'); } catch { res.writeHead(400); res.end('Bad request'); return; }
  const pathname = parsed.pathname;

  const fnMatch = pathname.match(/^\/\.netlify\/functions\/([A-Za-z0-9_-]+)(?:\/|$)/);
  if (fnMatch) {
    const fnName = fnMatch[1];
    const fnPath = path.join(__dirname, 'netlify', 'functions', `${fnName}.js`);
    if (!fs.existsSync(fnPath)) {
      res.writeHead(404); res.end(`Function not found: ${fnName}`); return;
    }
    try {
      delete require.cache[require.resolve(fnPath)];
      const fn = require(fnPath);

      let body = [];
      req.on('data', d => body.push(d));
      req.on('end', async () => {
        try {
          const event = {
            httpMethod: req.method,
            path: pathname,
            queryStringParameters: Object.fromEntries(parsed.searchParams),
            headers: req.headers,
            body: body.length ? Buffer.concat(body).toString('utf8') : null,
            isBase64Encoded: false,
          };
          const result = await fn.handler(event, {});
          const headers = { 'Content-Type': 'application/json', ...(result.headers || {}) };
          res.writeHead(result.statusCode || 200, headers);
          if (result.isBase64Encoded) {
            res.end(Buffer.from(result.body, 'base64'));
          } else {
            res.end(result.body || '');
          }
        } catch (err) {
          console.error(`[${fnName}] Error:`, err.message);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ message: err.message }));
        }
      });
    } catch (err) {
      console.error(`[${fnName}] Load error:`, err.message);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ message: err.message }));
    }
    return;
  }

  let rel;
  try { rel = decodeURIComponent(pathname); } catch { res.writeHead(400); res.end('Bad request'); return; }
  if (rel.includes('\0')) { res.writeHead(400); res.end('Bad request'); return; }
  let filePath = path.resolve(ROOT, '.' + (rel === '/' ? '/index.html' : rel));
  // Never serve anything outside the static root ("/../server.js" style traversal).
  if (filePath !== ROOT && !filePath.startsWith(ROOT + path.sep)) { res.writeHead(403); res.end('Forbidden'); return; }
  if (!path.extname(filePath)) filePath = path.join(filePath, 'index.html');
  // never serve anything outside public/ (blocks ../ tricks)
  if (path.relative(ROOT, filePath).startsWith('..') || path.isAbsolute(path.relative(ROOT, filePath))) { res.writeHead(403); res.end('Forbidden'); return; }

  fs.readFile(filePath, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    const ext = path.extname(filePath);
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(PORT, HOST, () => {
  console.log(`\n  MSP Report Builder (Intune) running at http://localhost:${PORT}\n`);
});
