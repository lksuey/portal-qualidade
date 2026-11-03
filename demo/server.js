'use strict';
/**
 * Local demo server (Node, no dependencies). Serves apps-script/Index.html and
 * exposes the public functions of Codigo.gs over /rpc, emulating google.script.run.
 * The library content is fictional and lives in Index.html. Binds to 127.0.0.1 only.
 *
 *   node demo/server.js [port]
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const { loadPortal, demoPassword, DEMO_LOGINS, ROOT } = require('./runtime');

const port = Number(process.argv[2] || process.env.PORT || 8766);
const rt = loadPortal();

const SHIM = `<script>
(function () {
  function runner(ok, fail) {
    return new Proxy({}, { get: function (_, name) {
      if (name === 'withSuccessHandler') return function (f) { return runner(f, fail); };
      if (name === 'withFailureHandler') return function (f) { return runner(ok, f); };
      return function () {
        var args = Array.prototype.slice.call(arguments);
        fetch('/rpc', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fn: name, args: args }) })
          .then(function (r) { return r.json(); })
          .then(function (r) { if (r.error) { if (fail) fail({ message: r.error }); } else if (ok) ok(r.result); })
          .catch(function (e) { if (fail) fail(e); });
      };
    } });
  }
  window.google = { script: { run: runner() } };
})();
</script>`;

function callPublic(fn, args) {
  // Same visibility rule as Apps Script: a trailing underscore means private.
  if (!/^[A-Za-z][A-Za-z0-9]*$/.test(fn) || typeof rt.context[fn] !== 'function') throw new Error('Unknown function: ' + fn);
  const result = rt.context[fn](...args);
  return JSON.parse(JSON.stringify(result === undefined ? null : result));
}

const server = http.createServer((req, res) => {
  if (req.method === 'GET' && (req.url === '/' || req.url.startsWith('/?'))) {
    const html = fs.readFileSync(path.join(ROOT, 'apps-script', 'Index.html'), 'utf8').replace('<head>', '<head>' + SHIM);
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    return res.end(html);
  }
  if (req.method === 'POST' && req.url === '/rpc') {
    let body = '';
    req.on('data', (c) => { body += c; if (body.length > 5e6) req.destroy(); });
    req.on('end', () => {
      let out;
      try {
        const { fn, args } = JSON.parse(body);
        out = { result: callPublic(fn, Array.isArray(args) ? args : []) };
      } catch (e) {
        out = { error: String(e && e.message || e) };
      }
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(out));
    });
    return;
  }
  res.writeHead(404); res.end('Not found');
});

server.listen(port, '127.0.0.1', () => {
  console.log('Portal da Qualidade demo running at http://127.0.0.1:' + port + '/');
  console.log('Demo accounts (synthetic): ' + DEMO_LOGINS.map((l) => l + ' / ' + demoPassword(l)).join('   '));
});
