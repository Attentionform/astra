/**
 * Shows, or records, the SHA-256 fingerprints of the CLA texts.
 *
 *   npx tsx scripts/cla/hash.ts           print the fingerprints and whether they match the config
 *   npx tsx scripts/cla/hash.ts --write   record them in .github/cla/config.json for the current version
 *
 * Signature files name the fingerprint of the text their author accepted, so once anyone has signed a
 * version its text must not change. --write therefore refuses to replace a recorded fingerprint when
 * signature files exist: raise "version" in the config (and keep the old text) instead.
 */
import * as fs from 'node:fs';
import { pathToFileURL } from 'node:url';
import { claTextSha256, validateConfig } from './evaluate';

const CONFIG = '.github/cla/config.json';
const CORPORATE = 'CLA-CORPORATE.md';

function signatureFiles(dir: string): string[] {
  return fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => /^\d+\.json$/.test(f)) : [];
}

function main(): void {
  const raw = JSON.parse(fs.readFileSync(CONFIG, 'utf8'));
  const v = String(raw.version);
  const claNow = claTextSha256(fs.readFileSync(raw.claPath, 'utf8'));
  const corpNow = claTextSha256(fs.readFileSync(CORPORATE, 'utf8'));
  const rows: [string, string, string | undefined][] = [
    [raw.claPath, claNow, raw.claSha256?.[v]],
    [CORPORATE, corpNow, raw.corporateSha256?.[v]],
  ];
  for (const [file, now, pinned] of rows) console.log(`${file}  ${now}  ${pinned === now ? 'matches the config' : pinned ? 'DIFFERS from the config' : 'not in the config'}`);
  if (!process.argv.includes('--write')) return;

  const signed = signatureFiles(raw.signaturesDir).length > 0;
  for (const [file, now, pinned] of rows) {
    if (pinned && pinned !== now && signed) {
      throw new Error(`${file} changed but version ${v} has already been signed. Raise "version" in ${CONFIG} and record the new text as a new version.`);
    }
  }
  // Replace only the recorded values so the rest of the file keeps its formatting. A new version's entries
  // are added to the two maps by hand, together with the new version number.
  let text = fs.readFileSync(CONFIG, 'utf8');
  for (const [key, now] of [['claSha256', claNow], ['corporateSha256', corpNow]] as const) {
    const re = new RegExp(`("${key}":\\s*\\{[^}]*"${v}":\\s*")[0-9a-f]{64}(")`);
    if (!re.test(text)) throw new Error(`add a "${v}" entry to "${key}" in ${CONFIG} first`);
    text = text.replace(re, `$1${now}$2`);
  }
  validateConfig(JSON.parse(text));
  fs.writeFileSync(CONFIG, text);
  console.log(`recorded the fingerprints for version ${v} in ${CONFIG}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    main();
  } catch (e) {
    console.error(e instanceof Error ? e.message : e);
    process.exit(1);
  }
}
