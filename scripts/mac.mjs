import { execFileSync } from 'node:child_process';
import { X509Certificate } from 'node:crypto';
import { once } from 'node:events';
import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import issuer from '../portable/certificate.cjs';
import { createPaneServer } from './pane-server.mjs';

const project = fileURLToPath(new URL('../', import.meta.url));
const name = 'Formula Trace localhost';
function fingerprint(pem) {
  const cert = new X509Certificate(pem);
  if (cert.subject !== 'CN=' + name || cert.ca || cert.subjectAltName !== 'DNS:localhost' ||
      !cert.verify(cert.publicKey)) throw Error('Unrecognized previous Formula Trace certificate.');
  return cert.fingerprint256.replaceAll(':', '');
}

// Inject the native command boundary in tests. Key generation, HTTPS and sideload files are real.
export async function startMac({ userHome = homedir(), port = 3000, run = execFileSync } = {}) {
  const library = join(userHome, 'Library');
  const keychain = join(library, 'Keychains', 'login.keychain-db');
  const state = join(library, 'Application Support', 'Formula Trace');
  const receipt = join(state, 'localhost.cer');
  const catalog = join(library, 'Containers/com.microsoft.Excel/Data/Documents/wef');
  const manifest = await readFile(join(project, 'manifest.xml'), 'utf8');
  const id = /<Id>([0-9a-f-]{36})<\/Id>/i.exec(manifest)?.[1];
  if (!id) throw Error('Missing Formula Trace manifest ID.');
  const workbook = join(project, 'examples/Debt-Model-Mac.xlsx');
  const identity = issuer.createIdentity();
  const current = fingerprint(identity.cert);
  const security = (...args) => run('/usr/bin/security', args, { encoding:'utf8' });
  function removeCertificate(hash) {
    let listed;
    try { listed = security('find-certificate', '-a', '-Z', '-c', name, keychain); }
    catch (error) { if (error.status === 44) return; throw error; }
    if (listed.toUpperCase().includes(hash)) security('delete-certificate', '-t', '-Z', hash, keychain);
  }

  const server = createPaneServer(identity);
  server.listen(port, 'localhost');
  await once(server, 'listening');
  let attemptedTrust = false;
  const cleanup = () => {
    if (!attemptedTrust) return;
    removeCertificate(current);
    attemptedTrust = false;
  };
  try {
    await mkdir(state, {recursive:true, mode:0o700});
    let prior;
    try { prior = await readFile(receipt); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    if (prior) removeCertificate(fingerprint(prior));
    await writeFile(receipt + '.new', identity.cert, {mode:0o600});
    await rename(receipt + '.new', receipt);
    console.log('Allow the macOS Keychain request for the Formula Trace localhost certificate.');
    attemptedTrust = true;
    security('add-trusted-cert', '-r', 'trustRoot', '-p', 'ssl', '-s', 'localhost', '-k', keychain, receipt);
    await mkdir(catalog, {recursive:true});
    const installed = join(catalog, id + '.manifest.xml');
    await writeFile(installed + '.new', manifest);
    await rename(installed + '.new', installed);
    run('/usr/bin/open', ['-a', 'Microsoft Excel', workbook]);
    return {server, cleanup};
  } catch (error) {
    server.close();
    try { cleanup(); } catch (failure) { console.error('Certificate cleanup:', failure.message); }
    throw error;
  }
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    if (process.platform !== 'darwin') throw Error('Open Start Formula Trace.command on a Mac.');
    const {cleanup} = await startMac();
    console.log('Select Home → Add-ins → Formula Trace. Keep this window open. Press Control+C to stop.');
    process.once('exit', () => {
      try { cleanup(); } catch (error) { console.error('Certificate cleanup:', error.message); }
    });
    for (const signal of ['SIGINT','SIGTERM','SIGHUP']) process.once(signal, () => process.exit(0));
  } catch (error) {
    console.error(error.code === 'EADDRINUSE' ? 'Port 3000 is in use. Close the previous launcher and retry.' : error.message);
    process.exitCode = 1;
  }
}
