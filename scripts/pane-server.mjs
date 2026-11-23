import https from 'node:https';
import { readFile } from 'node:fs/promises';

const files = new Map([
  ['index.html','text/html'], ['help.html','text/html'], ['shortcuts.json','application/json'], ['pane.css','text/css'],
  ...['excel','navigation','pane','expression','map','rows'].map(name => ['build/' + name + '.js','application/javascript']),
  ...[16,32,80].map(size => ['icon-' + size + '.png','image/png'])
]);

export function createPaneServer(options) {
  return https.createServer(options, async (request, response) => {
    const path = request.url?.split('?')[0].slice(1) || 'index.html';
    if (!files.has(path)) { response.writeHead(404).end(); return; }
    try {
      const body = await readFile(new URL('../public/' + path, import.meta.url));
      response.writeHead(200, { 'Content-Type': files.get(path), 'Cache-Control': 'no-store' }).end(body);
    } catch {
      response.writeHead(500).end('Build the add-in with npm run build.');
    }
  });
}
