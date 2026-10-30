/**
 * Creates your CLA signature file, ready to add to your pull request.
 *
 *   npm run cla:sign -- <github-login> "Your Full Name"           writes cla/signatures/<id>.json
 *   npm run cla:sign -- <github-login> "Your Full Name" --print   prints the file instead of writing it
 *
 * Reads your numeric id from the public GitHub API (the one thing that is easy to get wrong), fills in the
 * current CLA fingerprint and statement, and checks the result with the same rules the `cla` check uses
 * before it writes anything. It never overwrites a file that exists. Without npm: npx tsx scripts/cla/sign.ts ...
 */
import * as fs from 'node:fs';
import * as path from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseSignature, signaturePath, signatureTemplate, validateConfig } from './evaluate';

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const print = args.includes('--print');
  const [login, ...rest] = args.filter((a) => a !== '--print');
  const name = rest.join(' ').trim();
  if (!login || !name) {
    console.error('usage: npm run cla:sign -- <github-login> "Your Full Name" [--print]');
    process.exit(2);
  }
  if (!/^[A-Za-z0-9](?:[A-Za-z0-9-]{0,38})$/.test(login)) throw new Error('that is not a valid GitHub login');
  const config = validateConfig(JSON.parse(fs.readFileSync('.github/cla/config.json', 'utf8')));
  const res = await fetch(`https://api.github.com/users/${login}`, { headers: { Accept: 'application/vnd.github+json', 'User-Agent': 'astra-cla-sign' } });
  if (!res.ok) throw new Error(`GitHub answered ${res.status} for ${login}`);
  const user = (await res.json()) as { id?: number; login?: string };
  if (!Number.isInteger(user.id) || typeof user.login !== 'string') throw new Error('unexpected answer from GitHub');
  const id = user.id as number;
  const body = JSON.parse(signatureTemplate({ login: user.login, id }, config));
  body.name = name;
  const text = JSON.stringify(body, null, 2) + '\n';

  // The file the check would accept, or no file at all: the same parser, run here first.
  const checked = parseSignature(text, id, config);
  if (!checked.ok) throw new Error(`this file would not pass the cla check: ${checked.error}`);

  if (print) {
    console.error(`Save this as ${signaturePath(config, id)} and add it to your pull request:\n`);
    console.log(text);
    return;
  }
  const file = signaturePath(config, id);
  if (fs.existsSync(file)) throw new Error(`${file} already exists, so you have signed. Nothing was changed.`);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, text, { flag: 'wx' });
  console.log(`Wrote ${file}\n\nNext: git add ${file}, commit it with your own account (git commit -s), and push. The cla check turns green by itself.`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((e) => {
    console.error(e instanceof Error ? e.message : e);
    process.exit(1);
  });
}
