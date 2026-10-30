/**
 * CLA check tests: the signature-file format, who may add which file, coverage,
 * fail-closed behaviour, hostile text, and the orchestration (in-memory API).
 * No test touches the network.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import {
  type ClaConfig,
  type CommitInfo,
  type CorporateRecord,
  type EvalInput,
  type Person,
  type PrFile,
  evaluatePr,
  isAllowlisted,
  parseCorporate,
  parseSignature,
  renderComment,
  signaturePath,
  claTextSha256,
  signatureTemplate,
  statementFor,
  toCsv,
  validateConfig,
  verifyClaText,
} from '../../scripts/cla/evaluate';
import { MemoryApi, reportError, runCheck } from '../../scripts/cla/run';

const ROOT = path.resolve(__dirname, '..', '..');
const config: ClaConfig = validateConfig(JSON.parse(fs.readFileSync(path.join(ROOT, '.github/cla/config.json'), 'utf8')));

const jane: Person = { login: 'janedoe', id: 1234567, type: 'User' };
const bob: Person = { login: 'bobsmith', id: 7654321, type: 'User' };
const owner: Person = { login: 'Userfrom1995', id: 99, type: 'User' };
const robot: Person = { login: 'dependabot[bot]', id: 49699333, type: 'Bot' };
const nobody: Person = { login: null, id: null };

function commit(author: Person, extra: Partial<CommitInfo> = {}): CommitInfo {
  return { sha: `${(author.id ?? 0).toString(16).padStart(7, '0')}abcdef0`, author, committer: author, authorEmail: 'x@example.org', parentCount: 1, message: 'fix: x', ...extra };
}

function sigText(p: Person, over: Record<string, unknown> = {}): string {
  const base: Record<string, unknown> = {
    github_id: p.id,
    github_login: p.login,
    name: 'Jane Doe',
    cla_version: 1,
    cla_sha256: config.claSha256['1'],
    statement: statementFor(config, 1),
    ...over,
  };
  for (const k of Object.keys(base)) if (base[k] === undefined) delete base[k];
  return JSON.stringify(base, null, 2) + '\n';
}

/** A config whose current version is 2, with a made-up fingerprint for it. */
const V2_SHA = 'b'.repeat(64);
const withV2 = (accepted = [1, 2]): ClaConfig => ({ ...config, version: 2, acceptedVersions: accepted, claSha256: { ...config.claSha256, '2': V2_SHA } });

const spath = (p: Person) => signaturePath(config, p.id as number);
const added = (p: Person): PrFile => ({ path: spath(p), status: 'added' });

function input(over: Partial<EvalInput> = {}): EvalInput {
  return { commits: [commit(jane)], files: [], headTexts: {}, touchedBy: {}, baseTexts: {}, corporate: [], ...over };
}

/** A pull request by `p` that adds their own valid signature file. */
function selfSign(p: Person, over: Partial<EvalInput> = {}): EvalInput {
  const c = commit(p);
  return input({
    commits: [c],
    files: [added(p)],
    headTexts: { [spath(p)]: sigText(p) },
    touchedBy: { [spath(p)]: [c.sha] },
    ...over,
  });
}

describe('config', () => {
  it('is valid and consistent with the repository', () => {
    expect(config.acceptedVersions).toContain(config.version);
    expect(fs.existsSync(path.join(ROOT, config.claPath))).toBe(true);
    expect(fs.existsSync(path.join(ROOT, config.signaturesDir))).toBe(true);
    expect(fs.existsSync(path.join(ROOT, config.corporateDir))).toBe(true);
  });
  it('pins the exact text of both agreements', () => {
    const v = String(config.version);
    expect(claTextSha256(fs.readFileSync(path.join(ROOT, config.claPath), 'utf8'))).toBe(config.claSha256[v]);
    expect(claTextSha256(fs.readFileSync(path.join(ROOT, 'CLA-CORPORATE.md'), 'utf8'))).toBe(config.corporateSha256[v]);
  });
  it('ignores line-ending differences when fingerprinting', () => expect(claTextSha256('a\r\nb\n')).toBe(claTextSha256('a\nb\n')));
  it('protects the CLA machinery and the corporate agreements', () => {
    for (const p of ['cla/corporate/', '.github/cla/', 'scripts/cla/', '.github/workflows/cla.yml', 'CLA.md']) expect(config.protectedPaths).toContain(p);
  });
  it('rejects broken configs', () => {
    const raw = JSON.parse(fs.readFileSync(path.join(ROOT, '.github/cla/config.json'), 'utf8'));
    expect(() => validateConfig(null)).toThrow();
    expect(() => validateConfig({ ...raw, schema: 1 })).toThrow();
    expect(() => validateConfig({ ...raw, acceptedVersions: [7] })).toThrow();
    expect(() => validateConfig({ ...raw, statement: 'no placeholder' })).toThrow();
    expect(() => validateConfig({ ...raw, protectedPaths: 'x' })).toThrow();
    expect(() => validateConfig({ ...raw, allowlist: null })).toThrow();
    expect(() => validateConfig({ ...raw, claSha256: {} })).toThrow();
    expect(() => validateConfig({ ...raw, claSha256: { 1: 'abc' } })).toThrow();
    expect(() => validateConfig({ ...raw, corporateSha256: undefined })).toThrow();
  });
  it('uses the exact statement', () => expect(statementFor(config, 1)).toBe('I have read the ASTra CLA v1 and I agree to it.'));
});

describe('signature file format', () => {
  const parse = (text: string | null, id = jane.id as number) => parseSignature(text, id, config);
  it('accepts a valid file', () => expect(parse(sigText(jane))).toEqual({ ok: true, value: { github_id: 1234567, github_login: 'janedoe', name: 'Jane Doe', cla_version: 1, cla_sha256: config.claSha256['1'], statement: statementFor(config, 1) } }));
  it('rejects a file that cannot be read', () => expect(parse(null)).toMatchObject({ ok: false }));
  it('rejects invalid JSON', () => expect(parse('{nope')).toMatchObject({ ok: false, error: expect.stringContaining('JSON') }));
  it.each(['[]', '"x"', 'null', '7'])('rejects non-object JSON %s', (t) => expect(parse(t)).toMatchObject({ ok: false }));
  it('rejects an extra key', () => expect(parse(sigText(jane, { extra: 1 }))).toMatchObject({ ok: false, error: expect.stringContaining('exactly the keys') }));
  it.each(['github_id', 'github_login', 'name', 'cla_version', 'cla_sha256', 'statement'])('rejects a missing key %s', (k) => expect(parse(sigText(jane, { [k]: undefined }))).toMatchObject({ ok: false }));
  it('rejects an id that does not match the file name', () => expect(parse(sigText(jane), 5)).toMatchObject({ ok: false, error: expect.stringContaining('does not match') }));
  it.each(['1234567', 1.5, -1, 0, null])('rejects github_id %s', (v) => expect(parse(sigText(jane, { github_id: v }))).toMatchObject({ ok: false }));
  it.each(['', '-bad', 'has space', 'a'.repeat(40), 'ünï'])('rejects github_login %j', (v) => expect(parse(sigText(jane, { github_login: v }))).toMatchObject({ ok: false }));
  it('rejects a statement that is not exact', () => {
    for (const s of ['I have read the ASTra CLA v1 and I agree to it', 'i have read the astra cla v1 and i agree to it.', ' I have read the ASTra CLA v1 and I agree to it.', 'I agree.', '']) {
      expect(parse(sigText(jane, { statement: s }))).toMatchObject({ ok: false });
    }
  });
  it('rejects a statement that names a different version than cla_version', () => expect(parse(sigText(jane, { statement: statementFor(config, 2) }))).toMatchObject({ ok: false }));
  it('accepts a newer version when the statement and fingerprint match it', () => expect(parseSignature(sigText(jane, { cla_version: 2, statement: statementFor(config, 2), cla_sha256: V2_SHA }), jane.id as number, withV2())).toMatchObject({ ok: true }));
  it('rejects a wrong, missing-case or malformed fingerprint', () => {
    for (const h of ['0'.repeat(64), config.claSha256['1'].toUpperCase(), config.claSha256['1'].slice(1), '', null, 5]) expect(parse(sigText(jane, { cla_sha256: h }))).toMatchObject({ ok: false, error: expect.stringContaining('cla_sha256') });
  });
  it('rejects the fingerprint of another version', () => expect(parseSignature(sigText(jane, { cla_version: 2, statement: statementFor(config, 2), cla_sha256: config.claSha256['1'] }), jane.id as number, withV2())).toMatchObject({ ok: false }));
  it('rejects a version that has no recorded text', () => expect(parse(sigText(jane, { cla_version: 2, statement: statementFor(config, 2) }))).toMatchObject({ ok: false, error: expect.stringContaining('no CLA version') }));
  it.each(['J', '', ' Jane', 'Jane ', 'Ja\u0000ne', 'Jane\nDoe', 'A'.repeat(121)])('rejects name %j', (n) => expect(parse(sigText(jane, { name: n }))).toMatchObject({ ok: false }));
  it('rejects a name that is not a string', () => expect(parse(sigText(jane, { name: 5 }))).toMatchObject({ ok: false }));
  it('accepts non-ASCII names', () => expect(parse(sigText(jane, { name: 'José Müller 山田太郎' }))).toMatchObject({ ok: true }));
  it('rejects a file larger than the limit', () => expect(parse(sigText(jane) + ' '.repeat(3000))).toMatchObject({ ok: false, error: expect.stringContaining('larger') }));
  it('rejects bad cla_version values', () => {
    for (const v of [0, -1, 1.5, '1', null]) expect(parse(sigText(jane, { cla_version: v }))).toMatchObject({ ok: false });
  });
  it('rejects the template as it is handed out, until the contributor types their own name', () => {
    const t = JSON.parse(signatureTemplate(jane, config));
    expect(t.name).toBe('Your Full Name');
    expect(parse(JSON.stringify(t))).toMatchObject({ ok: false, error: expect.stringContaining('placeholder') });
    expect(parse(JSON.stringify({ ...t, name: 'Jane Doe' }))).toMatchObject({ ok: true });
  });
});

describe('allowlist', () => {
  it('matches logins case-insensitively and bot patterns', () => {
    expect(isAllowlisted(owner, config)).toBe(true);
    expect(isAllowlisted({ login: 'userfrom1995', id: 1 }, config)).toBe(true);
    expect(isAllowlisted(robot, config)).toBe(true);
    expect(isAllowlisted(jane, config)).toBe(false);
    expect(isAllowlisted(nobody, config)).toBe(false);
  });
  it('does not let a look-alike login through', () => {
    expect(isAllowlisted({ login: 'Userfrom1995x', id: 1 }, config)).toBe(false);
    expect(isAllowlisted({ login: 'evil[bot]x', id: 1 }, config)).toBe(false);
    expect(isAllowlisted({ login: 'fake-bot', id: 1 }, config)).toBe(false);
  });
});

describe('coverage', () => {
  it('fails an author with no signature and gives them their file', () => {
    const r = evaluatePr(input(), config);
    expect(r.state).toBe('failure');
    expect(r.uncovered.map((u) => u.login)).toEqual(['janedoe']);
    const body = renderComment(r, config, 'https://x/CLA.md');
    expect(body).toContain(config.botMarker);
    expect(body).toContain('cla/signatures/1234567.json');
    expect(body).toContain('npm run cla:sign -- <your-login>');
    expect(body).toContain('"github_id": 1234567');
    expect(body).toContain(statementFor(config, 1));
  });
  it('passes an author who signed earlier (file on the base branch)', () => {
    const r = evaluatePr(input({ baseTexts: { '1234567': sigText(jane) } }), config);
    expect(r.state).toBe('success');
    expect(r.people[0]).toMatchObject({ login: 'janedoe', reason: 'signed' });
  });
  it('ignores a base file that is invalid or for an old, unaccepted version', () => {
    expect(evaluatePr(input({ baseTexts: { '1234567': '{bad' } }), config).state).toBe('failure');
    expect(evaluatePr(input({ baseTexts: { '1234567': sigText(jane, { cla_version: 9, statement: statementFor(config, 9) }) } }), config).state).toBe('failure');
  });
  it('ignores a base file that names a different account', () => {
    expect(evaluatePr(input({ baseTexts: { '1234567': sigText(bob, { github_id: 7654321 }) } }), config).state).toBe('failure');
  });
  it('passes when the pull request adds the author\'s own valid file', () => {
    const r = evaluatePr(selfSign(jane), config);
    expect(r.state).toBe('success');
    expect(r.violations).toEqual([]);
    expect(r.people[0].reason).toBe('signed');
  });
  it('needs every author: one signed, one not', () => {
    const r = evaluatePr(input({ commits: [commit(jane), commit(bob)], baseTexts: { '1234567': sigText(jane) } }), config);
    expect(r.state).toBe('failure');
    expect(r.uncovered.map((u) => u.login)).toEqual(['bobsmith']);
  });
  it('checks the committer as well as the author', () => {
    const r = evaluatePr(input({ commits: [commit(jane, { committer: bob })], baseTexts: { '1234567': sigText(jane) } }), config);
    expect(r.uncovered.map((u) => u.login)).toEqual(['bobsmith']);
  });
  it('ignores web-flow committers, merge commits and GitHub-generated commits', () => {
    const web: Person = { login: 'web-flow', id: 19864447 };
    expect(evaluatePr(input({ commits: [commit(jane, { committer: web })], baseTexts: { '1234567': sigText(jane) } }), config).state).toBe('success');
    expect(evaluatePr(input({ commits: [commit(jane), commit(bob, { parentCount: 2 })], baseTexts: { '1234567': sigText(jane) } }), config).state).toBe('success');
    expect(evaluatePr(input({ commits: [commit(web)] }), config).state).toBe('success');
  });
  it('fails a commit that is not linked to any account', () => {
    const r = evaluatePr(input({ commits: [commit(jane), commit(nobody, { sha: 'dead000' })], baseTexts: { '1234567': sigText(jane) } }), config);
    expect(r.state).toBe('failure');
    expect(r.unlinkedCommits).toEqual(['dead000']);
    expect(renderComment(r, config, 'u')).toContain('not linked to a GitHub account');
  });
  it('reports co-authors but does not require them', () => {
    const r = evaluatePr(input({ commits: [commit(jane, { message: 'x\n\nCo-authored-by: Eve <eve@example.org>' })], baseTexts: { '1234567': sigText(jane) } }), config);
    expect(r.state).toBe('success');
    expect(r.coAuthors).toEqual(['Eve <eve@example.org>']);
  });
  it('passes the owner and bots with no files and no signatures', () => {
    expect(evaluatePr(input({ commits: [commit(owner)] }), config).people[0].reason).toBe('allowlist');
    expect(evaluatePr(input({ commits: [commit(robot)] }), config).state).toBe('success');
  });
  it('errors when there are no commits', () => {
    expect(evaluatePr(input({ commits: [] }), config).state).toBe('error');
  });
});

describe('corporate agreements', () => {
  const corp = (over: Partial<CorporateRecord> = {}): CorporateRecord => ({ company: 'Acme Ltd', agreement_ref: 'ACME-2026-01', cla_version: 1, logins: [], domains: [], ...over });
  it('covers a listed login (case-insensitive)', () => {
    const r = evaluatePr(input({ corporate: [corp({ logins: ['JaneDoe'] })] }), config);
    expect(r.state).toBe('success');
    expect(r.people[0]).toMatchObject({ reason: 'corporate', company: 'Acme Ltd' });
  });
  it('covers a linked author whose commit email is on a listed domain', () => {
    expect(evaluatePr(input({ commits: [commit(jane, { authorEmail: 'jane@acme.com' })], corporate: [corp({ domains: ['acme.com'] })] }), config).state).toBe('success');
  });
  it('does not match a look-alike domain or a sub-string', () => {
    for (const e of ['jane@acme.com.evil.io', 'jane@notacme.com', 'jane@evil.io?@acme.com', 'acme.com']) {
      expect(evaluatePr(input({ commits: [commit(jane, { authorEmail: e })], corporate: [corp({ domains: ['acme.com'] })] }), config).state).toBe('failure');
    }
  });
  it('ignores an agreement for a version that is not accepted', () => {
    expect(evaluatePr(input({ corporate: [corp({ logins: ['janedoe'], cla_version: 9 })] }), config).state).toBe('failure');
  });
  it('parses corporate files strictly', () => {
    const ok = JSON.stringify(corp({ logins: ['janedoe'], domains: ['acme.com'] }));
    expect(parseCorporate(ok, config)).toMatchObject({ ok: true });
    expect(parseCorporate('{x', config)).toMatchObject({ ok: false });
    expect(parseCorporate(JSON.stringify({ ...corp({ logins: ['a'] }), extra: 1 }), config)).toMatchObject({ ok: false });
    expect(parseCorporate(JSON.stringify(corp()), config)).toMatchObject({ ok: false }); // no login and no domain
    expect(parseCorporate(JSON.stringify(corp({ domains: ['ACME.com'] })), config)).toMatchObject({ ok: false });
    expect(parseCorporate(JSON.stringify(corp({ logins: ['-x'] })), config)).toMatchObject({ ok: false });
    expect(parseCorporate(JSON.stringify(corp({ company: 'A\u0000' , logins: ['a'] })), config)).toMatchObject({ ok: false });
  });
});

describe('who may add which file', () => {
  const rule = (r: ReturnType<typeof evaluatePr>) => r.violations.join('\n');
  it('rejects a file that names another account', () => {
    const c = commit(jane);
    const r = evaluatePr(input({ commits: [c], files: [added(bob)], headTexts: { [spath(bob)]: sigText(bob) }, touchedBy: { [spath(bob)]: [c.sha] }, baseTexts: { '1234567': sigText(jane) } }), config);
    expect(r.state).toBe('failure');
    expect(rule(r)).toContain('did not write any commit');
  });
  it('rejects bob signing for jane when both wrote commits, if the file is not in the signer\'s own commit', () => {
    const cj = commit(jane);
    const cb = commit(bob);
    const r = evaluatePr(
      { commits: [cj, cb], files: [added(jane)], headTexts: { [spath(jane)]: sigText(jane) }, touchedBy: { [spath(jane)]: [cb.sha] }, baseTexts: { '7654321': sigText(bob) }, corporate: [] },
      config
    );
    expect(r.state).toBe('failure');
    expect(rule(r)).toContain('written by account 1234567');
  });
  it('rejects a file that was touched by both authors\' commits', () => {
    const cj = commit(jane);
    const cb = commit(bob);
    const r = evaluatePr({ commits: [cj, cb], files: [added(jane)], headTexts: { [spath(jane)]: sigText(jane) }, touchedBy: { [spath(jane)]: [cj.sha, cb.sha] }, baseTexts: { '7654321': sigText(bob) }, corporate: [] }, config);
    expect(r.state).toBe('failure');
  });
  it('rejects a file whose commit is not in the pull request', () => {
    const r = evaluatePr(selfSign(jane, { touchedBy: { [spath(jane)]: ['0000000ffffffff'] } }), config);
    expect(r.state).toBe('failure');
  });
  it('rejects a valid file that names the wrong login', () => {
    const r = evaluatePr(selfSign(jane, { headTexts: { [spath(jane)]: sigText(jane, { github_login: 'someoneelse' }) } }), config);
    expect(rule(r)).toContain('is not the commit author');
  });
  it('accepts the login with different capitalisation', () => {
    expect(evaluatePr(selfSign(jane, { headTexts: { [spath(jane)]: sigText(jane, { github_login: 'JaneDoe' }) } }), config).state).toBe('success');
  });
  it('rejects an invalid file and explains why', () => {
    const r = evaluatePr(selfSign(jane, { headTexts: { [spath(jane)]: sigText(jane, { statement: 'ok' }) } }), config);
    expect(r.state).toBe('failure');
    expect(rule(r)).toContain('statement must be exactly');
  });
  it('rejects an unreadable file', () => {
    expect(evaluatePr(selfSign(jane, { headTexts: { [spath(jane)]: null } }), config).state).toBe('failure');
  });
  it('rejects signing an old CLA version', () => {
    const r = evaluatePr(selfSign(jane, { headTexts: { [spath(jane)]: sigText(jane, { cla_version: 0 }) } }), config);
    expect(r.state).toBe('failure');
  });
  it.each([
    ['a file with a letter in the name', 'cla/signatures/janedoe.json'],
    ['a nested path', 'cla/signatures/x/1234567.json'],
    ['a different extension', 'cla/signatures/1234567.txt'],
    ['a leading zero path trick', 'cla/signatures/../signatures/1234567.json'],
    ['a stray file', 'cla/signatures/README.md'],
  ])('rejects %s', (_label, p) => {
    const r = evaluatePr(input({ baseTexts: { '1234567': sigText(jane) }, files: [{ path: p, status: 'added' }], headTexts: { [p]: sigText(jane) } }), config);
    expect(r.state).toBe('failure');
  });
  it('rejects deleting, renaming or copying a signature, even your own', () => {
    for (const status of ['removed', 'renamed', 'copied']) {
      const r = evaluatePr(input({ baseTexts: { '1234567': sigText(jane) }, files: [{ path: spath(jane), status, previousPath: status === 'removed' ? undefined : 'cla/signatures/1.json' }] }), config);
      expect(r.state).toBe('failure');
    }
    const moved = evaluatePr(input({ baseTexts: { '1234567': sigText(jane) }, files: [{ path: 'docs/x.json', status: 'renamed', previousPath: spath(bob) }] }), config);
    expect(moved.state).toBe('failure');
  });
  it('allows raising your own cla_version, and nothing else', () => {
    const newer = { cla_version: 2, statement: statementFor(config, 2) };
    // the config accepts only version 1 as current; the edit must name the current version
    const c = commit(jane);
    const same = evaluatePr({ commits: [c], files: [{ path: spath(jane), status: 'modified' }], headTexts: { [spath(jane)]: sigText(jane, { name: 'Jane Q Doe' }) }, touchedBy: { [spath(jane)]: [c.sha] }, baseTexts: { '1234567': sigText(jane) }, corporate: [] }, config);
    expect(same.state).toBe('failure');
    expect(same.violations.join('\n')).toContain('only be updated to a newer');
    void newer;
  });
  it('allows raising the version when the current version is higher', () => {
    const v2 = withV2();
    const c = commit(jane);
    const r = evaluatePr(
      { commits: [c], files: [{ path: spath(jane), status: 'modified' }], headTexts: { [spath(jane)]: sigText(jane, { cla_version: 2, statement: statementFor(v2, 2), cla_sha256: V2_SHA }) }, touchedBy: { [spath(jane)]: [c.sha] }, baseTexts: { '1234567': sigText(jane) }, corporate: [] },
      v2
    );
    expect(r.state).toBe('success');
  });
  it('treats the old version as still valid while it is accepted', () => {
    const v2 = withV2();
    expect(evaluatePr(input({ baseTexts: { '1234567': sigText(jane) } }), v2).state).toBe('success');
    const v2only = withV2([2]);
    expect(evaluatePr(input({ baseTexts: { '1234567': sigText(jane) } }), v2only).state).toBe('failure');
  });
});

describe('owner-only paths', () => {
  const touch = (p: string, who: Person = jane, status = 'modified') =>
    evaluatePr(input({ commits: [commit(who)], files: [{ path: p, status }], baseTexts: { [String(who.id)]: sigText(who) } }), config);
  it.each(['cla/corporate/acme.json', '.github/cla/config.json', 'scripts/cla/evaluate.ts', '.github/workflows/cla.yml', 'CLA.md', 'CLA-CORPORATE.md'])('blocks a contributor from changing %s', (p) => {
    const r = touch(p);
    expect(r.state).toBe('failure');
    expect(r.violations.join('\n')).toContain('Only the owner');
  });
  it('blocks a rename that moves a protected file away', () => {
    const r = evaluatePr(input({ baseTexts: { '1234567': sigText(jane) }, files: [{ path: 'docs/x.md', status: 'renamed', previousPath: 'CLA.md' }] }), config);
    expect(r.state).toBe('failure');
  });
  it('lets the owner change them', () => {
    for (const p of ['cla/corporate/acme.json', '.github/cla/config.json', 'CLA.md', spath(bob)]) expect(touch(p, owner).state).toBe('success');
  });
  it('does not let a protected-path change ride along with a bot or owner co-signature', () => {
    const r = evaluatePr(input({ commits: [commit(owner), commit(jane)], files: [{ path: 'cla/corporate/acme.json', status: 'added' }], baseTexts: { '1234567': sigText(jane) } }), config);
    expect(r.state).toBe('failure');
  });
  it('does not block ordinary files', () => {
    expect(touch('src/index.ts').state).toBe('success');
    expect(touch('docs/cla-notes.md').state).toBe('success');
  });
});

describe('hostile text', () => {
  it('escapes markdown, HTML and mentions that come from the pull request', () => {
    const evil = { login: 'x', id: 5 };
    const r = evaluatePr(input({ commits: [commit(evil)], files: [{ path: 'cla/signatures/<img src=x onerror=1>@everyone.json', status: 'added' }] }), config);
    const body = renderComment(r, config, 'u');
    expect(body).not.toContain('<img');
    expect(body).not.toContain('@everyone');
  });
  it('treats a name with HTML as data and never puts it in the comment', () => {
    const r = evaluatePr(selfSign(jane, { headTexts: { [spath(jane)]: sigText(jane, { name: '<script>alert(1)</script>' }) } }), config);
    expect(r.state).toBe('success');
    expect(renderComment(r, config, 'u')).not.toContain('<script>');
  });
  it('handles deeply odd JSON without throwing', () => {
    for (const t of ['{"__proto__":{"x":1}}', '{"github_id":1e999}', '\u0000', '{"a":'.repeat(500), '"' + 'x'.repeat(100000) + '"']) {
      expect(() => parseSignature(t, 1, config)).not.toThrow();
      expect(parseSignature(t, 1, config).ok).toBe(false);
    }
  });
});

describe('comment text', () => {
  it('says passed with the names of who is covered', () => {
    const body = renderComment(evaluatePr(input({ baseTexts: { '1234567': sigText(jane) } }), config), config, 'u');
    expect(body).toContain('CLA check: passed');
    expect(body).toContain('@janedoe: has signed');
  });
  it('lists problems separately from missing signatures', () => {
    const body = renderComment(evaluatePr(input({ commits: [commit(jane)], files: [{ path: 'CLA.md', status: 'modified' }], baseTexts: { '1234567': sigText(jane) } }), config), config, 'u');
    expect(body).toContain('breaks the CLA rules');
    expect(body).toContain('Only the owner may change CLA.md');
  });
  it('renders an error as failed closed', () => {
    const empty = evaluatePr(input({ commits: [] }), config);
    expect(renderComment(empty, config, 'u', { error: 'boom' })).toContain('could not decide');
  });
});

describe('csv export', () => {
  it('neutralises formula injection and quotes separators', () => {
    const csv = toCsv([{ github_id: 1, github_login: 'a', name: '=HYPERLINK("x")', cla_version: 1, cla_sha256: '', statement: '' }], [{ company: 'A, B', agreement_ref: '+1', cla_version: 1, logins: ['x', 'y'], domains: [] }]);
    expect(csv).toContain(`"'=HYPERLINK(""x"")"`);
    expect(csv).toContain('"A, B"');
    expect(csv).toContain("'+1");
  });
});

describe('orchestration (in-memory API)', () => {
  const pr = { number: 7, headSha: 'h'.repeat(40), state: 'open' };
  const base = (p: Person) => ({ [spath(p)]: sigText(p) });

  it('marks an unsigned pull request red and comments once', async () => {
    const api = new MemoryApi(pr, [commit(jane)], []);
    const r = await runCheck(api, config, 7);
    expect(r.state).toBe('failure');
    expect(api.statuses.at(-1)).toMatchObject({ sha: pr.headSha, state: 'failure' });
    expect(api.comments).toHaveLength(1);
    await runCheck(api, config, 7);
    expect(api.comments).toHaveLength(1);
  });
  it('turns green once the author\'s file is in the pull request, and updates the comment', async () => {
    const c = commit(jane);
    const api = new MemoryApi(pr, [c], []);
    await runCheck(api, config, 7);
    api.files = [added(jane)];
    api.head = { [spath(jane)]: sigText(jane) };
    api.touching = { [spath(jane)]: [c.sha] };
    const r = await runCheck(api, config, 7);
    expect(r.state).toBe('success');
    expect(api.statuses.at(-1)?.state).toBe('success');
    expect(api.comments).toHaveLength(1);
    expect(api.comments[0].body).toContain('passed');
  });
  it('is green and silent for a signed contributor', async () => {
    const api = new MemoryApi(pr, [commit(jane)], [], base(jane));
    expect((await runCheck(api, config, 7)).state).toBe('success');
    expect(api.comments).toHaveLength(0);
  });
  it('is green and silent for the owner', async () => {
    const api = new MemoryApi(pr, [commit(owner)], [{ path: 'CLA.md', status: 'modified' }]);
    expect((await runCheck(api, config, 7)).state).toBe('success');
    expect(api.comments).toHaveLength(0);
  });
  it('reads corporate agreements from the default branch', async () => {
    const api = new MemoryApi(pr, [commit(jane)], [], { 'cla/corporate/acme.json': JSON.stringify({ company: 'Acme', agreement_ref: 'A-1', cla_version: 1, logins: ['janedoe'], domains: [] }), 'cla/corporate/README.md': 'x' });
    expect((await runCheck(api, config, 7)).state).toBe('success');
  });
  it('fails closed on a broken corporate file on the default branch', async () => {
    const api = new MemoryApi(pr, [commit(jane)], [], { 'cla/corporate/acme.json': '{broken' });
    await expect(runCheck(api, config, 7)).rejects.toThrow(/invalid/);
    expect(api.statuses).toHaveLength(0);
  });
  it('fails closed when the API throws, and never sets success', async () => {
    const api = new MemoryApi(pr, [commit(jane)], []);
    api.listFiles = async () => {
      throw new Error('GitHub API GET /pulls/7/files failed with 502');
    };
    await expect(runCheck(api, config, 7)).rejects.toThrow('502');
    await reportError(api, config, 7, '502');
    expect(api.statuses.at(-1)?.state).toBe('error');
    expect(api.statuses.some((s) => s.state === 'success')).toBe(false);
    expect(api.comments[0].body).toContain('could not decide');
  });
  it('fails closed, and sets no status, when CLA.md on the default branch is not the pinned text', async () => {
    const api = new MemoryApi(pr, [commit(jane)], [], { 'CLA.md': 'a quietly edited agreement' });
    await expect(runCheck(api, config, 7)).rejects.toThrow(/does not match the fingerprint/);
    expect(api.statuses).toHaveLength(0);
  });
  it('verifies the CLA text', () => {
    const text = fs.readFileSync(path.join(ROOT, config.claPath), 'utf8');
    expect(() => verifyClaText(text, config)).not.toThrow();
    expect(() => verifyClaText(text.replace(/\n/g, '\r\n'), config)).not.toThrow();
    expect(() => verifyClaText(`${text} `, config)).toThrow();
    expect(() => verifyClaText(null, config)).toThrow(/could not be read/);
  });
  it('handles the sample fixture', async () => {
    const fx = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/cla/unsigned-pr.json'), 'utf8'));
    const api = new MemoryApi(fx.pr, fx.commits, fx.files, fx.base, fx.head, fx.touching);
    expect((await runCheck(api, config, fx.pr.number)).state).toBe('failure');
  });
});

describe('workflow safety', () => {
  const yml = fs.readFileSync(path.join(ROOT, '.github/workflows/cla.yml'), 'utf8');
  it('runs only on pull_request_target and never on comments', () => {
    expect(yml).toContain('pull_request_target:');
    expect(yml).not.toContain('issue_comment');
  });
  it('never checks out the pull request code and cannot write to the repository', () => {
    expect(yml).toContain('contents: read');
    expect(yml).not.toMatch(/contents:\s*write/);
    expect(yml).not.toMatch(/ref:\s*\$\{\{\s*github\.event\.pull_request/);
    expect(yml).not.toContain('github.head_ref');
    expect(yml).toContain('persist-credentials: false');
  });
  it('pins actions to full commit SHAs', () => {
    for (const m of yml.matchAll(/^\s*-?\s*uses:\s*(\S+)/gm)) expect(m[1]).toMatch(/@[0-9a-f]{40}$/);
  });
  it('does not interpolate pull request text into a shell', () => {
    const runs = [...yml.matchAll(/run:\s*(.+)/g)].map((m) => m[1]);
    for (const r of runs) expect(r).not.toContain('${{');
  });
  it('has no other workflow that writes signatures', () => {
    expect(fs.existsSync(path.join(ROOT, '.github/workflows/cla-admin.yml'))).toBe(false);
  });
});

describe('the way to sign is reachable', () => {
  it('has an npm script for the signing command, and the guides name it', () => {
    const scripts = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).scripts as Record<string, string>;
    expect(scripts['cla:sign']).toContain('scripts/cla/sign.ts');
    expect(fs.existsSync(path.join(ROOT, 'scripts/cla/sign.ts'))).toBe(true);
    for (const guide of ['CONTRIBUTING.md', 'docs/contributing.md', '.github/PULL_REQUEST_TEMPLATE.md']) {
      expect(fs.readFileSync(path.join(ROOT, guide), 'utf8'), guide).toContain('npm run cla:sign');
    }
  });
});
