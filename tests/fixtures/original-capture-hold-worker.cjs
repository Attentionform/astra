const fsp = require('node:fs/promises');
const path = require('node:path');
const engine = require('@attentionform/astra-engine');

async function main() {
  const [document, payloadBase64, readyFile] = process.argv.slice(2);
  const link = fsp.link.bind(fsp);
  fsp.link = async (from, to) => {
    if (path.basename(String(to)) === 'original.snapshot') {
      await fsp.writeFile(readyFile, String(from));
      await new Promise(() => { setInterval(() => {}, 60_000); });
    }
    return link(from, to);
  };
  await engine.captureOriginal(path.resolve(document), Buffer.from(payloadBase64, 'base64'));
}

main().catch((error) => {
  process.stderr.write(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 1;
});
