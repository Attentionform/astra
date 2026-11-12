const fsp = require('node:fs/promises');
const path = require('node:path');
const { initAstraRepo, originalSha256 } = require('@attentionform/astra-engine');

async function waitFor(file, timeoutMs) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try { await fsp.access(file); return; } catch { await new Promise((resolve) => setTimeout(resolve, 5)); }
  }
  throw new Error(`timed out waiting for ${file}`);
}

async function main() {
  const [document, xml, payload, id, barrier] = process.argv.slice(2);
  const bytes = Buffer.from(payload, 'base64');
  await fsp.writeFile(path.join(barrier, `${id}.ready`), 'ready');
  await waitFor(path.join(barrier, 'go'), 15_000);
  try {
    await initAstraRepo(document, xml, {}, bytes);
    process.stdout.write(JSON.stringify({ ok: true, id, sha256: originalSha256(bytes) }));
  } catch (error) {
    process.stdout.write(JSON.stringify({ ok: false, id, code: error && error.code, message: error instanceof Error ? error.message : String(error) }));
    process.exitCode = 1;
  }
}

main().catch((error) => {
  process.stderr.write(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 2;
});
