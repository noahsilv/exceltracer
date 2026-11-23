import { execFileSync } from 'node:child_process';
import { X509Certificate } from 'node:crypto';
import { mkdir, readFile, writeFile, rename } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { once } from 'node:events';
import { createPaneServer } from './pane-server.mjs';
import issuer from '../portable/certificate.cjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const registry = 'HKCU\\Software\\Microsoft\\Office\\16.0\\Wef\\Developer';

function certId(pem) {
  const certificate = new X509Certificate(pem);
  if (certificate.subject !== 'CN=Formula Trace localhost' || certificate.ca ||
      certificate.subjectAltName !== 'DNS:localhost' || !certificate.verify(certificate.publicKey)) {
    throw Error('Unexpected localhost certificate identity.');
  }
  return certificate.fingerprint.replaceAll(':', '');
}

function certUtil(...args) {
  return execFileSync('certutil.exe', ['-user', ...args], {stdio: 'pipe'});
}

function trusted(id) {
  try { certUtil('-store', 'Root', id); return true; }
  catch (error) { if (error.code === 'ENOENT') throw error; return false; }
}

async function run() {
  if (process.platform !== 'win32') throw Error('Open this launcher on Windows with desktop Excel.');
  const data = join(process.env.LOCALAPPDATA, 'FormulaTrace');
  const publicFile = join(data, 'last-localhost.cer');
  const workbook = join(root, 'examples', 'Debt-Model.xlsx');
  const manifest = join(root, 'manifest.xml');
  const xml = await readFile(manifest, 'utf8');
  const id = xml.match(/<Id>([0-9a-f-]{36})<\/Id>/i)?.[1];
  if (!id) throw Error('Could not read the add-in ID from manifest.xml. Extract the complete ZIP.');
  const identity = issuer.createIdentity();
  const server = createPaneServer(identity);
  server.listen(3000, 'localhost');
  await once(server, 'listening');
  let added = false;
  const currentId = certId(identity.cert);
  const cleanup = () => {
    if (added) {
      try { certUtil('-delstore', 'Root', currentId); }
      catch { console.error('Windows could not remove the localhost certificate. Thumbprint: ' + currentId); }
      added = false;
    }
  };
  process.once('exit', cleanup);
  process.once('SIGINT', () => process.exit(0));
  try {
    await mkdir(data, {recursive:true});
    let prior;
    try { prior = await readFile(publicFile); } catch (e) { if (e.code !== 'ENOENT') throw e; }
    // Recover after the previous console was forcibly closed. Only the public cert is saved.
    if (prior) {
      const priorId = certId(prior);
      if (trusted(priorId)) certUtil('-delstore', 'Root', priorId);
    }
    const staged = publicFile + '.new';
    await writeFile(staged, identity.cert);
    await rename(staged, publicFile);
    console.log('Setting up the Formula Trace localhost certificate for this Windows account…');
    certUtil('-addstore', 'Root', publicFile);
    added = true;
    execFileSync('reg.exe', ['add', registry, '/v', id, '/t', 'REG_SZ', '/d', manifest, '/f']);
    console.log('Opening the sample in Excel. Keep this window open while tracing. Press Ctrl+C to stop.');
    execFileSync('powershell.exe', ['-NoProfile','-NonInteractive','-Command',
      "$ErrorActionPreference = 'Stop'; Start-Process -FilePath 'excel.exe' -ArgumentList ('\"' + $env:FORMULA_TRACE_BOOK + '\"')"
    ], {env:{...process.env, FORMULA_TRACE_BOOK:workbook}, stdio:'inherit'});
  } catch (error) {
    server.close();
    throw error;
  }
}
run().catch(error => {
  const detail = error.stderr?.toString().trim() || error.stdout?.toString().trim();
  console.error(error.code === 'EADDRINUSE'
    ? 'Port 3000 is already in use. Close the earlier Formula Trace window or free port 3000, then retry.'
    : error.message);
  if (detail) console.error(detail);
  process.exitCode = 1;
});
