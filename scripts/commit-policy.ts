/**
 * Commit rules for pull requests.
 *
 *  1. The author email is a real, reachable-looking address. GitHub "hidden"
 *     (noreply) addresses and placeholder domains are refused.
 *  2. The commit carries `Signed-off-by: Name <email>` and the email matches
 *     the author email.
 *  3. The author is a person, not a bot or an AI tool's own identity. Using AI
 *     tools is fine; the human who commits is responsible for the result.
 *
 * Usage: npx tsx scripts/commit-policy.ts <base> <head>
 * Checks every non-merge commit in base..head. Exit 1 on any problem.
 */
import { execFileSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

export interface CommitInfo {
  sha: string;
  name: string;
  email: string;
  body: string;
}

const EMAIL = /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>.]{2,}$/;
const PLACEHOLDER_TLDS = new Set(['invalid', 'local', 'localhost', 'example', 'test', 'internal', 'lan', 'home', 'localdomain']);
const PLACEHOLDER_DOMAINS = new Set(['example.com', 'example.net', 'example.org', 'users.noreply.github.com', 'noreply.github.com', 'noreply.gitlab.com']);
const AUTOMATED_NAMES = new Set(['claude', 'claude code', 'copilot', 'github copilot', 'chatgpt', 'codex', 'gemini', 'devin', 'cursor agent', 'cursoragent']);

function isDependabot(c: CommitInfo): boolean {
  return c.name === 'dependabot[bot]' && /dependabot\[bot\]@users\.noreply\.github\.com$/i.test(c.email);
}

export function checkCommit(c: CommitInfo): string[] {
  if (isDependabot(c)) return [];
  const problems: string[] = [];
  const email = c.email.trim().toLowerCase();
  const domain = email.split('@')[1] ?? '';
  const local = email.split('@')[0] ?? '';

  if (!c.name.trim()) problems.push('the author name is empty');
  if (!EMAIL.test(email)) {
    problems.push(`the author email "${c.email}" is not a valid address`);
  } else {
    const tld = domain.split('.').pop() ?? '';
    if (domain.endsWith('.noreply.github.com') || PLACEHOLDER_DOMAINS.has(domain) || PLACEHOLDER_TLDS.has(tld) || /^no-?reply$/.test(local)) {
      problems.push(
        `the author email "${c.email}" is hidden or a placeholder; use your real email (git config user.email) and amend the commit`
      );
    }
  }
  if (/\[bot\]$/i.test(c.name) || AUTOMATED_NAMES.has(c.name.trim().toLowerCase())) {
    problems.push(`the author "${c.name}" is an automated or AI identity; commits must come from a person who takes responsibility for them`);
  }

  const signers = [...c.body.matchAll(/^Signed-off-by:\s*(.+?)\s*<([^>]+)>\s*$/gim)].map((m) => m[2].trim().toLowerCase());
  if (signers.length === 0) {
    problems.push('there is no "Signed-off-by:" line; commit with `git commit -s`');
  } else if (!signers.includes(email)) {
    problems.push(`no "Signed-off-by:" email matches the author email "${c.email}"`);
  }
  return problems;
}

export function parseLog(raw: string): CommitInfo[] {
  return raw
    .split('\x1e')
    .map((r) => r.replace(/^\n/, ''))
    .filter((r) => r.trim().length > 0)
    .map((r) => {
      const [sha, name, email, ...rest] = r.split('\x1f');
      return { sha: sha.trim(), name, email, body: rest.join('\x1f') };
    });
}

export function checkRange(base: string, head: string): { sha: string; problems: string[] }[] {
  const raw = execFileSync('git', ['log', '--no-merges', '--format=%H%x1f%an%x1f%ae%x1f%B%x1e', `${base}..${head}`], { encoding: 'utf8' });
  return parseLog(raw)
    .map((c) => ({ sha: c.sha, problems: checkCommit(c) }))
    .filter((r) => r.problems.length > 0);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const [base, head = 'HEAD'] = process.argv.slice(2);
  if (!base) {
    console.error('usage: npx tsx scripts/commit-policy.ts <base> [head]');
    process.exit(2);
  }
  const bad = checkRange(base, head);
  for (const r of bad) {
    console.error(`${r.sha.slice(0, 10)}:`);
    for (const p of r.problems) console.error(`  - ${p}`);
  }
  if (bad.length > 0) {
    console.error(`\n${bad.length} commit(s) break the commit rules. See CONTRIBUTING.md.`);
    process.exit(1);
  }
  console.log('commit rules: ok');
}
