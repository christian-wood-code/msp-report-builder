// Local dev server -- added so this project can be exercised locally without
// `netlify dev` (needed to test report-hub calling this app's functions
// directly). Static root is "." (per netlify.toml's publish = "."), not a
// public/ subfolder, unlike the sibling report apps. Usage: node server.js

const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  fs.readFileSync(envPath, 'utf8').split('\n').forEach(line => {
    const [k, ...v] = line.split('=');
    if (k && v.length && !k.trim().startsWith('#')) process.env[k.trim()] = v.join('=').trim();
  });
}

const PORT = process.env.PORT || 8791;
const ROOT = __dirname;

const MIME = {
  '.html': 'text/html', '.css': 'text/css', '.js': 'application/javascript',
  '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.b64': 'text/plain',
};

const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);
  const pathname = parsed.pathname;

  const fnMatch = pathname.match(/^\/.netlify\/functions\/([^/?]+)/);
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
            queryStringParameters: Object.fromEntries(
              Object.entries(parsed.query).map(([k, v]) => [k, String(v)])
            ),
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

  let filePath = path.join(ROOT, pathname === '/' ? 'index.html' : pathname);
  if (!path.extname(filePath)) filePath = path.join(filePath, 'index.html');

  fs.readFile(filePath, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    const ext = path.extname(filePath);
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`\n  MSP Report Builder (Intune) running at http://localhost:${PORT}\n`);
});
