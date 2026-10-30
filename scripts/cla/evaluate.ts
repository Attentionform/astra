/**
 * CLA decision logic. Pure functions only: no network, no file access.
 *
 * A contributor signs by adding one file, `cla/signatures/<github-id>.json`,
 * in their own pull request. The check verifies, from data the caller has
 * already fetched, that:
 *   - every person who wrote commits is covered (allowlist, a valid signature
 *     file on the base branch or added by this pull request, or a corporate
 *     agreement);
 *   - a signature file is added only by the account it names, in the exact
 *     format, with the exact statement, for the current CLA version;
 *   - the pull request touches no one else's signature and none of the
 *     owner-only paths.
 *
 * Fail closed: anything that cannot be decided is an error, never a pass.
 */

import { createHash } from 'node:crypto';

export interface ClaConfig {
  schema: number;
  version: number;
  acceptedVersions: number[];
  /** SHA-256 of each individual CLA version's text (see claTextSha256), keyed by version. Frozen once anyone has signed. */
  claSha256: Record<string, string>;
  /** Same for the corporate agreement. Checked by a test, not at run time. */
  corporateSha256: Record<string, string>;
  statement: string;
  minNameLength: number;
  statusContext: string;
  signaturesDir: string;
  corporateDir: string;
  claPath: string;
  botMarker: string;
  contact: string;
  allowlist: { logins: string[]; patterns: string[] };
  protectedPaths: string[];
}

export interface SignatureRecord {
  github_id: number;
  github_login: string;
  name: string;
  cla_version: number;
  /** SHA-256 of the CLA text the contributor accepted. */
  cla_sha256: string;
  statement: string;
}

export interface CorporateRecord {
  company: string;
  agreement_ref: string;
  cla_version: number;
  logins: string[];
  domains: string[];
}

/** A GitHub account tied to a commit. `login` and `id` are null when the commit email matches no account. */
export interface Person {
  login: string | null;
  id: number | null;
  type?: string;
}

export interface CommitInfo {
  sha: string;
  author: Person;
  committer: Person;
  authorEmail: string;
  parentCount: number;
  message: string;
}

export interface PrFile {
  path: string;
  status: string;
  previousPath?: string;
}

export interface EvalInput {
  commits: CommitInfo[];
  files: PrFile[];
  /** Text of each changed signature file as it is in the pull request (null when it cannot be read). */
  headTexts: Record<string, string | null>;
  /** For each changed signature path, the shas of the pull request's commits that touched it. */
  touchedBy: Record<string, string[]>;
  /** Text of `<signaturesDir>/<id>.json` on the base branch, by account id (null when absent). */
  baseTexts: Record<string, string | null>;
  corporate: CorporateRecord[];
}

export type CoverReason = 'allowlist' | 'signed' | 'corporate';

export interface PersonResult {
  login: string | null;
  id: number | null;
  covered: boolean;
  reason?: CoverReason;
  company?: string;
  commits: string[];
}

export interface EvaluationResult {
  state: 'success' | 'failure' | 'error';
  people: PersonResult[];
  uncovered: PersonResult[];
  unlinkedCommits: string[];
  violations: string[];
  coAuthors: string[];
  messages: string[];
}

export function statementFor(config: ClaConfig, version = config.version): string {
  return config.statement.replace('{version}', String(version));
}

export function signaturePath(config: ClaConfig, id: number): string {
  return `${config.signaturesDir}/${id}.json`;
}

// eslint-disable-next-line no-control-regex
const CONTROL = /[\u0000-\u001f\u007f]/;
const LOGIN = /^[A-Za-z0-9](?:[A-Za-z0-9-]{0,38})$/;
const DOMAIN = /^(?:[a-z0-9-]+\.)+[a-z]{2,}$/;
const SIGNATURE_KEYS = ['cla_sha256', 'cla_version', 'github_id', 'github_login', 'name', 'statement'];
const SHA256_HEX = /^[0-9a-f]{64}$/;

/** Fingerprint of an agreement's text: SHA-256 of the UTF-8 bytes with line endings normalised to LF. */
export function claTextSha256(text: string): string {
  return createHash('sha256').update(text.replace(/\r\n/g, '\n'), 'utf8').digest('hex');
}

/** Throws unless `text` (the CLA on the default branch) is exactly the text the config pins for the current version. */
export function verifyClaText(text: string | null, config: ClaConfig): void {
  if (text === null) throw new Error(`${config.claPath} could not be read from the default branch`);
  const pinned = config.claSha256[String(config.version)];
  if (claTextSha256(text) !== pinned) {
    throw new Error(`${config.claPath} on the default branch does not match the fingerprint pinned for version ${config.version}. Change the text only together with a new version number.`);
  }
}
const MAX_FILE_BYTES = 2000;
export const NAME_PLACEHOLDER = 'Your Full Name';

function globToRegex(pattern: string): RegExp {
  const escaped = pattern.replace(/[.+?^${}()|\\[\]]/g, '\\$&').replace(/\*/g, '.*');
  return new RegExp(`^${escaped}$`, 'i');
}

export function isAllowlisted(person: Person, config: ClaConfig): boolean {
  const login = person.login;
  if (!login) return false;
  const lower = login.toLowerCase();
  if (config.allowlist.logins.some((l) => l.toLowerCase() === lower)) return true;
  return config.allowlist.patterns.some((p) => globToRegex(p).test(login));
}

export function validateConfig(raw: unknown): ClaConfig {
  const c = raw as Partial<ClaConfig> | null;
  const bad = (why: string): never => {
    throw new Error(`invalid CLA config: ${why}`);
  };
  if (!c || typeof c !== 'object') return bad('not an object');
  if (c.schema !== 2) bad('unknown schema');
  if (!Number.isInteger(c.version) || (c.version as number) < 1) bad('version');
  if (!Array.isArray(c.acceptedVersions) || !c.acceptedVersions.every(Number.isInteger) || !c.acceptedVersions.includes(c.version as number)) bad('acceptedVersions must include version');
  const hashes = (m: unknown, versions: number[]): boolean =>
    !!m && typeof m === 'object' && versions.every((v) => typeof (m as Record<string, unknown>)[String(v)] === 'string' && SHA256_HEX.test((m as Record<string, string>)[String(v)]));
  if (!hashes(c.claSha256, c.acceptedVersions as number[])) bad('claSha256 must hold a SHA-256 for every accepted version');
  if (!hashes(c.corporateSha256, [c.version as number])) bad('corporateSha256 must hold a SHA-256 for the current version');
  if (typeof c.statement !== 'string' || !c.statement.includes('{version}')) bad('statement');
  for (const k of ['statusContext', 'signaturesDir', 'corporateDir', 'claPath', 'botMarker', 'contact'] as const) {
    if (typeof c[k] !== 'string' || !(c[k] as string)) bad(k);
  }
  if (!c.allowlist || !Array.isArray(c.allowlist.logins) || !Array.isArray(c.allowlist.patterns)) bad('allowlist');
  if (!Array.isArray(c.protectedPaths) || !c.protectedPaths.every((p) => typeof p === 'string' && p.length > 0)) bad('protectedPaths');
  if (!Number.isInteger(c.minNameLength)) bad('minNameLength');
  return c as ClaConfig;
}

export type Parsed<T> = { ok: true; value: T } | { ok: false; error: string };

/**
 * Strict parse of a signature file. `expectedId` is the id in the file name.
 * Exactly these keys, no others; the id must equal the file name; the
 * statement must be the exact sentence for the stated version.
 */
export function parseSignature(text: string | null, expectedId: number, config: ClaConfig): Parsed<SignatureRecord> {
  if (text === null) return { ok: false, error: 'the file could not be read' };
  if (Buffer.byteLength(text, 'utf8') > MAX_FILE_BYTES) return { ok: false, error: `the file is larger than ${MAX_FILE_BYTES} bytes` };
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    return { ok: false, error: 'the file is not valid JSON' };
  }
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return { ok: false, error: 'the file must hold one JSON object' };
  const o = raw as Record<string, unknown>;
  const keys = Object.keys(o).sort();
  if (keys.join(',') !== SIGNATURE_KEYS.join(',')) return { ok: false, error: `the file must have exactly the keys ${SIGNATURE_KEYS.join(', ')}` };
  if (!Number.isInteger(o.github_id) || (o.github_id as number) <= 0) return { ok: false, error: 'github_id must be a positive integer' };
  if (o.github_id !== expectedId) return { ok: false, error: `github_id ${String(o.github_id)} does not match the file name (${expectedId})` };
  if (typeof o.github_login !== 'string' || !LOGIN.test(o.github_login)) return { ok: false, error: 'github_login is not a valid GitHub login' };
  if (typeof o.name !== 'string') return { ok: false, error: 'name must be a string' };
  const name = o.name.trim();
  if (name !== o.name || name.length < config.minNameLength || name.length > 120 || CONTROL.test(name)) {
    return { ok: false, error: `name must be ${config.minNameLength} to 120 characters, with no leading or trailing spaces and no control characters` };
  }
  if (name === NAME_PLACEHOLDER) return { ok: false, error: 'put your own full name in "name" (it still has the placeholder)' };
  if (!Number.isInteger(o.cla_version) || (o.cla_version as number) < 1) return { ok: false, error: 'cla_version must be a positive integer' };
  const version = o.cla_version as number;
  if (o.statement !== statementFor(config, version)) return { ok: false, error: `statement must be exactly: ${statementFor(config, version)}` };
  const pinned = config.claSha256[String(version)];
  if (!pinned) return { ok: false, error: `there is no CLA version ${version}` };
  if (o.cla_sha256 !== pinned) return { ok: false, error: `cla_sha256 must be the fingerprint of CLA version ${version}: ${pinned}` };
  return { ok: true, value: { github_id: expectedId, github_login: o.github_login, name, cla_version: version, cla_sha256: pinned, statement: o.statement } };
}

const slug = /^[a-z0-9][a-z0-9-]{0,60}$/;

export function parseCorporate(text: string, config: ClaConfig): Parsed<CorporateRecord> {
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    return { ok: false, error: 'not valid JSON' };
  }
  const o = raw as Record<string, unknown>;
  if (!o || typeof o !== 'object' || Array.isArray(o)) return { ok: false, error: 'not an object' };
  if (Object.keys(o).sort().join(',') !== ['agreement_ref', 'cla_version', 'company', 'domains', 'logins'].join(',')) {
    return { ok: false, error: 'a corporate file must have exactly the keys agreement_ref, cla_version, company, domains, logins' };
  }
  const company = o.company;
  const ref = o.agreement_ref;
  if (typeof company !== 'string' || company.trim().length < 2 || company.length > 100 || CONTROL.test(company)) return { ok: false, error: 'company is invalid' };
  if (typeof ref !== 'string' || ref.trim().length < 3 || ref.length > 200 || CONTROL.test(ref)) return { ok: false, error: 'agreement_ref is invalid' };
  if (!Number.isInteger(o.cla_version)) return { ok: false, error: 'cla_version is invalid' };
  if (!Array.isArray(o.logins) || !o.logins.every((l) => typeof l === 'string' && LOGIN.test(l))) return { ok: false, error: 'logins must be GitHub logins' };
  if (!Array.isArray(o.domains) || !o.domains.every((d) => typeof d === 'string' && DOMAIN.test(d))) return { ok: false, error: 'domains must be lower-case domains' };
  if (o.logins.length === 0 && o.domains.length === 0) return { ok: false, error: 'give at least one login or domain' };
  void config;
  return { ok: true, value: { company, agreement_ref: ref, cla_version: o.cla_version as number, logins: o.logins as string[], domains: o.domains as string[] } };
}

export function isCorporateSlug(name: string): boolean {
  return slug.test(name);
}

function emailDomain(email: string): string {
  const parts = email.trim().split('@');
  return parts.length === 2 && parts[0].length > 0 ? parts[1].toLowerCase() : '';
}

export function parseCoAuthors(message: string): string[] {
  return [...message.matchAll(/^Co-authored-by:\s*(.+?)\s*$/gim)].map((m) => m[1]);
}

function underPrefix(path: string, prefix: string): boolean {
  return prefix.endsWith('/') ? path.startsWith(prefix) : path === prefix;
}

export function evaluatePr(input: EvalInput, config: ClaConfig): EvaluationResult {
  const out: EvaluationResult = { state: 'success', people: [], uncovered: [], unlinkedCommits: [], violations: [], coAuthors: [], messages: [] };
  const { commits } = input;
  if (!Array.isArray(commits) || commits.length === 0) {
    return { ...out, state: 'error', messages: ['The pull request has no commits to check.'] };
  }
  const relevant = commits.filter((c) => c.parentCount <= 1 && c.author.login !== 'web-flow');
  if (relevant.length === 0) return { ...out, messages: ['Only merge or GitHub-generated commits; nothing to check.'] };

  // 1. Who wrote what.
  const byId = new Map<number, { person: Person; commits: string[]; emails: string[] }>();
  for (const c of relevant) {
    for (const co of parseCoAuthors(c.message)) if (!out.coAuthors.includes(co)) out.coAuthors.push(co);
    const involved: Person[] = [c.author];
    if (c.committer.login !== 'web-flow' && c.committer.id !== c.author.id) involved.push(c.committer);
    for (const p of involved) {
      if (p.id === null || p.login === null) {
        if (!out.unlinkedCommits.includes(c.sha)) out.unlinkedCommits.push(c.sha);
        continue;
      }
      const e = byId.get(p.id) ?? { person: p, commits: [], emails: [] };
      e.commits.push(c.sha);
      if (p === c.author && c.authorEmail) e.emails.push(c.authorEmail);
      byId.set(p.id, e);
    }
  }
  const everyoneAllowlisted = byId.size > 0 && out.unlinkedCommits.length === 0 && [...byId.values()].every((e) => isAllowlisted(e.person, config));

  // 2. Rules for the files this pull request touches. The owner may change anything.
  const newlyValid = new Map<number, SignatureRecord>();
  if (!everyoneAllowlisted) {
    const sigPrefix = `${config.signaturesDir}/`;
    for (const f of input.files) {
      const paths = [f.path, ...(f.previousPath ? [f.previousPath] : [])];
      for (const p of paths) {
        if (config.protectedPaths.some((pp) => underPrefix(p, pp))) {
          out.violations.push(`Only the owner may change ${p}.`);
        }
      }
      if (!f.path.startsWith(sigPrefix) && !(f.previousPath ?? '').startsWith(sigPrefix)) continue;
      const m = new RegExp(`^${sigPrefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\d{1,15})\\.json$`).exec(f.path);
      if (!m || f.previousPath || f.status === 'removed' || f.status === 'renamed' || f.status === 'copied') {
        out.violations.push(`${f.path}: a pull request may only add or update its own file named ${sigPrefix}<your-github-id>.json.`);
        continue;
      }
      const id = Number.parseInt(m[1], 10);
      const person = byId.get(id)?.person;
      if (!person) {
        out.violations.push(`${f.path}: account ${id} did not write any commit in this pull request.`);
        continue;
      }
      const parsed = parseSignature(input.headTexts[f.path] ?? null, id, config);
      if (!parsed.ok) {
        out.violations.push(`${f.path}: ${parsed.error}.`);
        continue;
      }
      if (parsed.value.github_login.toLowerCase() !== (person.login ?? '').toLowerCase()) {
        out.violations.push(`${f.path}: github_login "${parsed.value.github_login}" is not the commit author "${person.login}".`);
        continue;
      }
      if (parsed.value.cla_version !== config.version) {
        out.violations.push(`${f.path}: sign the current CLA version (${config.version}), not ${parsed.value.cla_version}.`);
        continue;
      }
      const touchers = input.touchedBy[f.path] ?? [];
      const prShas = new Set(commits.map((c) => c.sha));
      const inPr = touchers.filter((s) => prShas.has(s));
      const authorOf = new Map(commits.map((c) => [c.sha, c.author.id] as const));
      if (inPr.length === 0 || inPr.some((s) => authorOf.get(s) !== id)) {
        out.violations.push(`${f.path}: it must be added in commits written by account ${id} (${person.login}) only.`);
        continue;
      }
      const prev = parseSignature(input.baseTexts[String(id)] ?? null, id, config);
      if (f.status === 'modified' && prev.ok && parsed.value.cla_version <= prev.value.cla_version) {
        out.violations.push(`${f.path}: an existing signature may only be updated to a newer CLA version.`);
        continue;
      }
      newlyValid.set(id, parsed.value);
    }
  }

  // 3. Coverage.
  for (const [id, e] of byId) {
    let reason: CoverReason | undefined;
    let company: string | undefined;
    if (isAllowlisted(e.person, config)) reason = 'allowlist';
    else if (newlyValid.has(id)) reason = 'signed';
    else {
      const base = parseSignature(input.baseTexts[String(id)] ?? null, id, config);
      if (base.ok && config.acceptedVersions.includes(base.value.cla_version)) reason = 'signed';
    }
    if (!reason) {
      const login = (e.person.login ?? '').toLowerCase();
      for (const c of input.corporate) {
        if (!config.acceptedVersions.includes(c.cla_version)) continue;
        // A linked commit author means GitHub matched the commit email to a verified email of that account.
        const byLogin = c.logins.some((l) => l.toLowerCase() === login);
        const byDomain = e.emails.some((m) => c.domains.some((d) => d.toLowerCase() === emailDomain(m)));
        if (byLogin || byDomain) {
          reason = 'corporate';
          company = c.company;
          break;
        }
      }
    }
    const r: PersonResult = { login: e.person.login, id, covered: reason !== undefined, commits: e.commits };
    if (reason) r.reason = reason;
    if (company) r.company = company;
    out.people.push(r);
    if (!r.covered) out.uncovered.push(r);
  }
  if (out.unlinkedCommits.length > 0) {
    out.messages.push(`Commit ${out.unlinkedCommits.map((s) => s.slice(0, 7)).join(', ')} is not linked to a GitHub account. Fix the author email (git config user.email) and push again.`);
  }
  if (out.uncovered.length > 0 || out.unlinkedCommits.length > 0 || out.violations.length > 0) out.state = 'failure';
  return out;
}

/** Makes text from the pull request safe to show in a comment: no HTML, no mentions, no links or images. */
function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/@/g, '&#64;')
    .replace(/[\\`[\]()!]/g, '\\$&')
    .replace(/\s+/g, ' ');
}

/** The exact file a contributor should add, with their id and login already filled in. */
export function signatureTemplate(person: { login: string | null; id: number | null }, config: ClaConfig): string {
  return JSON.stringify(
    {
      github_id: person.id,
      github_login: person.login,
      name: NAME_PLACEHOLDER,
      cla_version: config.version,
      cla_sha256: config.claSha256[String(config.version)],
      statement: statementFor(config),
    },
    null,
    2
  );
}

export function renderComment(result: EvaluationResult, config: ClaConfig, claUrl: string, opts: { error?: string } = {}): string {
  const lines: string[] = [config.botMarker];
  if (opts.error) {
    lines.push('**CLA check: could not decide**', '', `The check failed closed. Reason: ${esc(opts.error)}`, '', `If this keeps happening, write to ${config.contact}.`);
    return lines.join('\n');
  }
  if (result.state === 'success') {
    lines.push('**CLA check: passed**', '', 'Everyone who wrote commits in this pull request is covered.', '');
    for (const p of result.people) {
      const how = p.reason === 'allowlist' ? 'on the allowlist' : p.reason === 'corporate' ? `covered by the corporate agreement of ${esc(p.company ?? '')}` : 'has signed';
      lines.push(`- @${p.login}: ${how}`);
    }
    return lines.join('\n');
  }
  const n = result.uncovered.length + (result.unlinkedCommits.length > 0 ? 1 : 0);
  lines.push(n > 0 ? `**CLA check: ${n} contributor${n === 1 ? ' has' : 's have'} not signed**` : '**CLA check: this pull request breaks the CLA rules**', '');
  if (result.uncovered.length > 0) {
    lines.push(
      'Thank you for the contribution. ASTra is source-available under the Business Source License. To accept outside code, the owner needs a signed Contributor License Agreement. You sign once, by adding one small file to this pull request.',
      '',
      `1. Read the [CLA](${claUrl}). You keep your copyright; you give the owner the right to use your contribution, including under a commercial license.`,
      `2. Add the file named below to this pull request and put your full name in \`name\`. Change nothing else in it. You can create it on github.com with "Add file, Create new file". In a clone of the repository, \`npm run cla:sign -- <your-login> "Your Full Name"\` writes the file for you.`,
      '3. Push. This check runs again by itself.',
      `4. Contributing for your employer? Ask them to sign the corporate CLA instead; write to ${config.contact}.`,
      ''
    );
    for (const p of result.uncovered) {
      lines.push(`**@${p.login}** adds \`${signaturePath(config, p.id as number)}\`:`, '', '```json', signatureTemplate(p, config), '```', '');
    }
    lines.push('| Contributor | Status |', '| --- | --- |');
    for (const p of result.people) lines.push(`| @${p.login} | ${p.covered ? 'covered' : '**needs to sign**'} |`);
  }
  if (result.violations.length > 0) {
    lines.push('', '**Problems found**', '');
    for (const v of result.violations) lines.push(`- ${esc(v)}`);
  }
  for (const m of result.messages) lines.push('', esc(m));
  if (result.coAuthors.length > 0) {
    lines.push('', `Note: co-authors named in commit trailers are not checked (${result.coAuthors.length}). Each person who wrote code needs to sign in their own pull request.`);
  }
  return lines.join('\n');
}

export function toCsv(signatures: SignatureRecord[], corporate: CorporateRecord[]): string {
  const q = (v: string | number): string => {
    const s = String(v);
    const safe = /^[=+\-@\t\r]/.test(s) ? `'${s}` : s;
    return /[",\n\r]/.test(safe) ? `"${safe.replace(/"/g, '""')}"` : safe;
  };
  const rows = ['type,login,id,name_or_company,version,reference'];
  for (const s of signatures) rows.push(['individual', s.github_login, s.github_id, s.name, s.cla_version, `${'cla/signatures'}/${s.github_id}.json`].map(q).join(','));
  for (const c of corporate) rows.push(['corporate', c.logins.join(' '), '', c.company, c.cla_version, c.agreement_ref].map(q).join(','));
  return rows.join('\n') + '\n';
}
