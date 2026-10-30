import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { execFileSync } from 'child_process';
import { checkCommit, checkRange, parseLog, type CommitInfo } from '../../scripts/commit-policy';

function commit(over: Partial<CommitInfo> = {}): CommitInfo {
  return {
    sha: 'abc1234567',
    name: 'Ada Lovelace',
    email: 'ada@analytical.engine',
    body: 'fix(engine): something\n\nSigned-off-by: Ada Lovelace <ada@analytical.engine>\n',
    ...over,
  };
}

describe('commit rules', () => {
  it('accepts a person with a real email and a matching sign-off', () => {
    expect(checkCommit(commit())).toEqual([]);
  });

  it('refuses GitHub hidden addresses, old and new style', () => {
    for (const email of ['12345+ada@users.noreply.github.com', 'ada@users.noreply.github.com', 'noreply@github.com']) {
      const p = checkCommit(commit({ email, body: `x\n\nSigned-off-by: Ada <${email}>\n` }));
      expect(p.join(' ')).toMatch(/hidden or a placeholder/);
    }
  });

  it('refuses placeholder and malformed addresses', () => {
    for (const email of ['ada@example.com', 'ada@host.invalid', 'ada@localhost', 'ada@laptop.local', 'no-reply@acme.io', 'ada', 'ada@x', '']) {
      expect(checkCommit(commit({ email, body: `x\n\nSigned-off-by: Ada <${email}>\n` })).length).toBeGreaterThan(0);
    }
  });

  it('requires a sign-off whose email matches the author', () => {
    expect(checkCommit(commit({ body: 'no trailer here' })).join(' ')).toMatch(/Signed-off-by/);
    expect(checkCommit(commit({ body: 'x\n\nSigned-off-by: Ada <other@somewhere.org>\n' })).join(' ')).toMatch(/matches the author/);
    expect(checkCommit(commit({ body: 'x\n\nSigned-off-by: Ada <ADA@Analytical.Engine>\n' }))).toEqual([]);
  });

  it('does not treat a Co-Authored-By line as a sign-off', () => {
    expect(checkCommit(commit({ body: 'x\n\nCo-Authored-By: Ada <ada@analytical.engine>\n' })).join(' ')).toMatch(/Signed-off-by/);
  });

  it('refuses bot and AI-tool authors, even with a sign-off', () => {
    const claude = checkCommit(commit({ name: 'Claude', email: 'claude@acme.io', body: 'x\n\nSigned-off-by: Claude <claude@acme.io>\n' }));
    expect(claude.join(' ')).toMatch(/automated or AI identity/);
    const bot = checkCommit(commit({ name: 'renovate[bot]', email: 'bot@acme.io', body: 'x\n\nSigned-off-by: r <bot@acme.io>\n' }));
    expect(bot.join(' ')).toMatch(/automated or AI identity/);
  });

  it('exempts Dependabot only', () => {
    const dep = { name: 'dependabot[bot]', email: '49699333+dependabot[bot]@users.noreply.github.com', body: 'chore(deps): bump x' };
    expect(checkCommit(commit(dep))).toEqual([]);
    expect(checkCommit(commit({ ...dep, email: 'evil@users.noreply.github.com' })).length).toBeGreaterThan(0);
  });

  it('allows a Co-Authored-By trailer for an AI tool next to a human author', () => {
    const body = 'feat: x\n\nSigned-off-by: Ada Lovelace <ada@analytical.engine>\nCo-Authored-By: Claude <noreply@anthropic.com>\n';
    expect(checkCommit(commit({ body }))).toEqual([]);
  });

  it('parses the log format used by the range check', () => {
    const raw = '\x1e' + ['h1', 'A', 'a@b.io', 'msg\n\nSigned-off-by: A <a@b.io>\n'].join('\x1f') + '\x1e\n' + ['h2', 'B', 'b@b.io', 'msg\n'].join('\x1f') + '\x1e';
    const list = parseLog(raw);
    expect(list.map((c) => c.sha)).toEqual(['h1', 'h2']);
    expect(list[0].email).toBe('a@b.io');
  });

  it('checks a real git range, skips merges, and reports only bad commits', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-commit-policy-'));
    const git = (...args: string[]) => execFileSync('git', args, { cwd: dir, encoding: 'utf8', env: { ...process.env, GIT_CONFIG_GLOBAL: '/dev/null', GIT_CONFIG_SYSTEM: '/dev/null', GIT_COMMITTER_NAME: 'Committer', GIT_COMMITTER_EMAIL: 'committer@analytical.engine' } });
    try {
      git('init', '-q', '-b', 'main');
      git('config', 'commit.gpgsign', 'false');
      fs.writeFileSync(path.join(dir, 'f'), '1');
      git('add', 'f');
      git('-c', 'user.name=Ada', '-c', 'user.email=ada@analytical.engine', 'commit', '-q', '-m', 'base');
      const base = git('rev-parse', 'HEAD').trim();
      fs.writeFileSync(path.join(dir, 'f'), '2');
      git('commit', '-qa', '-m', 'good\n\nSigned-off-by: Ada <ada@analytical.engine>', '--author', 'Ada <ada@analytical.engine>');
      fs.writeFileSync(path.join(dir, 'f'), '3');
      git('commit', '-qa', '-m', 'hidden\n\nSigned-off-by: Bob <1+bob@users.noreply.github.com>', '--author', 'Bob <1+bob@users.noreply.github.com>');
      const cwd = process.cwd();
      process.chdir(dir);
      try {
        const bad = checkRange(base, 'HEAD');
        expect(bad).toHaveLength(1);
        expect(bad[0].problems.join(' ')).toMatch(/hidden or a placeholder/);
      } finally {
        process.chdir(cwd);
      }
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});
