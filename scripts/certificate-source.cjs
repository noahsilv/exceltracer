// Bundled at release time so the Windows launcher can run without installing npm packages.
const { generateKeyPairSync, randomBytes } = require('node:crypto');
const pki = require('node-forge/lib/pki');
const sha256 = require('node-forge/lib/sha256');

exports.createIdentity = function createIdentity() {
  const pair = generateKeyPairSync('rsa', {
    modulusLength: 2048,
    publicKeyEncoding: {type: 'spki', format: 'pem'},
    privateKeyEncoding: {type: 'pkcs8', format: 'pem'}
  });
  const cert = pki.createCertificate();
  cert.publicKey = pki.publicKeyFromPem(pair.publicKey);
  cert.serialNumber = '01' + randomBytes(15).toString('hex');
  cert.validity.notBefore = new Date(Date.now() - 300000);
  cert.validity.notAfter = new Date(Date.now() + 90 * 86400000);
  const name = [{ name: 'commonName', value: 'Formula Trace localhost' }];
  cert.setSubject(name);
  cert.setIssuer(name);
  cert.setExtensions([
    {name: 'basicConstraints', critical: true, cA: false},
    {name: 'keyUsage', critical: true, digitalSignature: true, keyEncipherment: true},
    {name: 'extKeyUsage', serverAuth: true},
    {name: 'subjectAltName', altNames: [{type: 2, value: 'localhost'}]}
  ]);
  cert.sign(pki.privateKeyFromPem(pair.privateKey), sha256.create());
  return { key: pair.privateKey, cert: pki.certificateToPem(cert) };
};
