import test from 'node:test';
import assert from 'node:assert/strict';
import https from 'node:https';
import { execFileSync } from 'node:child_process';
import { X509Certificate } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startMac } from '../scripts/mac.mjs';

test('Mac launcher serves the pane, installs the manifest and opens the Mac workbook', async t => {
  const userHome = await mkdtemp(join(tmpdir(), 'formula trace '));
  t.after(() => rm(userHome, {recursive:true,force:true}));
  const commands = [];
  let ca, installedHash;
  function run(command,args) {
    commands.push([command,...args]);
    switch (args[0]) {
      case 'add-trusted-cert':
        ca = readFileSync(args.at(-1));
        installedHash = new X509Certificate(ca).fingerprint256.replaceAll(':','');
        assert.deepEqual(args.slice(1,-1), ['-r','trustRoot','-p','ssl','-s','localhost','-k',join(userHome,'Library/Keychains/login.keychain-db')]);
        break;
      case 'find-certificate': return 'SHA-256 hash: '+installedHash;
      case 'delete-certificate': assert.deepEqual(args.slice(1,4), ['-t','-Z',installedHash]); break;
    }
    return '';
  }
  const {server,cleanup} = await startMac({userHome,port:0,run});
  t.after(() => {server.close();cleanup();});
  const catalog = join(userHome,'Library/Containers/com.microsoft.Excel/Data/Documents/wef');
  assert.equal(await readFile(join(catalog,'030c881e-c1db-4256-b653-21a708cec84d.manifest.xml'),'utf8'),
               await readFile(new URL('../manifest.xml',import.meta.url),'utf8'));
  assert.deepEqual(commands.at(-1).slice(0,3),['/usr/bin/open','-a','Microsoft Excel']);
  assert.match(commands.at(-1)[3],/Debt-Model-Mac\.xlsx$/);
  if (process.platform === 'darwin') {
    const cert = join(userHome, 'Library/Application Support/Formula Trace/localhost.cer');
    execFileSync('/usr/bin/security', ['verify-cert','-c',cert,'-r',cert,'-p','ssl','-n','localhost','-L']);
  }
  for (const route of ['index.html','build/pane.js','shortcuts.json']) {
    const body = await new Promise((resolve,reject)=>{
      https.get({hostname:'localhost',port:server.address().port,path:'/'+route,ca}, res=>{
        const chunks=[];
        res.on('data',b=>chunks.push(b));
        res.on('end',()=>resolve(Buffer.concat(chunks)));
        res.on('error',reject);
      }).on('error',reject);
    });
    assert.deepEqual(body,await readFile(new URL('../public/'+route,import.meta.url)));
  }
  cleanup();
  assert.equal(commands.at(-1)[1],'delete-certificate');
});
