// Runs a first history initialization and stops forever just before the initial commit, after the original
// snapshot has been published and while the initialization lock is held. A test then kills this process.
const fsp = require('node:fs/promises');
const path = require('node:path');
const git = require('isomorphic-git');
const engine = require('@attentionform/astra-engine');

async function main() {
  const [document, xml, payloadBase64, readyFile] = process.argv.slice(2);
  const hang = async () => {
    await fsp.writeFile(readyFile, 'published');
    await new Promise(() => { setInterval(() => {}, 60_000); });
  };
  const target = git.default && git.default.commit ? git.default : git;
  target.commit = hang;
  await engine.initAstraRepo(path.resolve(document), xml, {}, Buffer.from(payloadBase64, 'base64'));
}

main().catch((error) => {
  process.stderr.write(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 1;
});
