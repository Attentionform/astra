import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { initAstraRepo, getHistory, checkoutCommit, getHistoryCommitCount, getRepoDir } from '../../packages/engine/src/git';

describe('history reads never follow a redirected .git entry', () => {
  let dir: string;
  let victim: string;
  let attacker: string;

  beforeEach(async () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-gitdir-'));
    victim = path.join(dir, 'victim.docx');
    attacker = path.join(dir, 'attacker.docx');
    fs.writeFileSync(victim, 'v');
    fs.writeFileSync(attacker, 'a');
    await initAstraRepo(victim, '<xml>private contents</xml>', {});
    const sidecar = getRepoDir(attacker);
    fs.mkdirSync(sidecar, { recursive: true });
    fs.writeFileSync(path.join(sidecar, '.git'), `gitdir: ${path.join(getRepoDir(victim), '.git')}\n`);
  });

  afterEach(() => {
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('getHistory refuses the redirect instead of returning the other repository', async () => {
    await expect(getHistory(attacker)).rejects.toMatchObject({ code: 'E_IO' });
    await expect(getHistory(attacker)).rejects.toThrow(/not a repository directory/);
  });

  it('checkoutCommit refuses the redirect and returns none of the other repository', async () => {
    const [head] = await getHistory(victim);
    await expect(checkoutCommit(attacker, head.oid)).rejects.toMatchObject({ code: 'E_IO' });
  });

  it('the commit count is 0 for a redirected sidecar', async () => {
    expect(await getHistoryCommitCount(getRepoDir(attacker))).toBe(0);
    expect(await getHistoryCommitCount(getRepoDir(victim))).toBe(1);
  });

  it('a normal sidecar still reads and checks out', async () => {
    const history = await getHistory(victim);
    expect(history).toHaveLength(1);
    const out = await checkoutCommit(victim, history[0].oid);
    expect(out.xml).toContain('private contents');
  });

  it('a symlinked repository directory is still usable', async () => {
    const linked = path.join(dir, 'linked.docx');
    fs.writeFileSync(linked, 'l');
    const sidecar = getRepoDir(linked);
    fs.mkdirSync(sidecar, { recursive: true });
    fs.symlinkSync(path.join(getRepoDir(victim), '.git'), path.join(sidecar, '.git'));
    expect(await getHistory(linked)).toHaveLength(1);
  });
});
