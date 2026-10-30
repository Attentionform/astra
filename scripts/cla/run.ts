/**
 * CLA check entry point, run by .github/workflows/cla.yml on the DEFAULT
 * branch (never on pull request code).
 *
 *   npx tsx scripts/cla/run.ts
 *   npx tsx scripts/cla/run.ts --dry-run tests/fixtures/cla/<fixture>.json
 *
 * Only pull_request_target events. The check reads the pull request through the
 * API, decides with evaluate.ts, and writes a commit status plus one comment.
 * It never writes to the repository. Fails closed: any error sets the status to
 * `error` and exits 1; success is only ever set when everything checks out.
 */
import * as fs from 'node:fs';
import { pathToFileURL } from 'node:url';
import {
  type ClaConfig,
  type CommitInfo,
  type CorporateRecord,
  type EvaluationResult,
  type PrFile,
  evaluatePr,
  parseCorporate,
  renderComment,
  signaturePath,
  validateConfig,
  verifyClaText,
} from './evaluate';
import { type ClaApi, GitHubApi, type PrInfo } from './github';

export async function runCheck(api: ClaApi, config: ClaConfig, prNumber: number): Promise<EvaluationResult> {
  // Signatures name a fingerprint of the CLA text; refuse to decide if the text on the default branch is not that text.
  verifyClaText(await api.readBase(config.claPath), config);
  const pr = await api.getPr(prNumber);
  const commits = await api.listCommits(prNumber);
  const files = await api.listFiles(prNumber);

  const sigPrefix = `${config.signaturesDir}/`;
  const sigFiles = files.filter((f) => f.status !== 'removed' && f.path.startsWith(sigPrefix));
  const headTexts: Record<string, string | null> = {};
  const touchedBy: Record<string, string[]> = {};
  for (const f of sigFiles) {
    headTexts[f.path] = await api.readHead(f.path, pr.headSha);
    touchedBy[f.path] = await api.commitsTouching(f.path, pr.headSha);
  }

  // Signatures already on the default branch, for everyone who wrote a commit.
  const ids = new Set<number>();
  for (const c of commits) for (const p of [c.author, c.committer]) if (p.id !== null) ids.add(p.id);
  const baseTexts: Record<string, string | null> = {};
  for (const id of ids) baseTexts[String(id)] = await api.readBase(signaturePath(config, id));

  const corporate: CorporateRecord[] = [];
  for (const name of (await api.listBaseDir(config.corporateDir)).filter((n) => n.endsWith('.json'))) {
    const text = await api.readBase(`${config.corporateDir}/${name}`);
    const parsed = parseCorporate(text ?? '', config);
    if (!parsed.ok) throw new Error(`${config.corporateDir}/${name} on the default branch is invalid: ${parsed.error}`);
    corporate.push(parsed.value);
  }

  const result = evaluatePr({ commits, files, headTexts, touchedBy, baseTexts, corporate }, config);
  await publish(api, config, pr, result);
  return result;
}

async function publish(api: ClaApi, config: ClaConfig, pr: PrInfo, result: EvaluationResult): Promise<void> {
  const state = result.state;
  const description =
    state === 'success'
      ? 'All contributors are covered by the CLA'
      : state === 'failure'
        ? result.uncovered.length + result.unlinkedCommits.length > 0
          ? `${result.uncovered.length + (result.unlinkedCommits.length > 0 ? 1 : 0)} contributor(s) must sign the CLA`
          : 'This pull request breaks the CLA rules'
        : result.messages[0] ?? 'The CLA check could not decide';
  await api.setStatus(pr.headSha, state, description, config.statusContext);
  const existing = await api.findBotComment(pr.number, config.botMarker);
  if (state === 'success' && !existing) return;
  const body = renderComment(result, config, api.claUrl());
  if (existing && existing.body === body) return;
  await api.upsertBotComment(pr.number, body, existing?.id ?? null);
}

/** Best effort: record the failure on the pull request, never turn it green. */
export async function reportError(api: ClaApi, config: ClaConfig, prNumber: number, message: string): Promise<void> {
  try {
    const pr = await api.getPr(prNumber);
    await api.setStatus(pr.headSha, 'error', `CLA check failed closed: ${message}`, config.statusContext);
    const existing = await api.findBotComment(pr.number, config.botMarker);
    const empty: EvaluationResult = { state: 'error', people: [], uncovered: [], unlinkedCommits: [], violations: [], coAuthors: [], messages: [message] };
    await api.upsertBotComment(pr.number, renderComment(empty, config, api.claUrl(), { error: message }), existing?.id ?? null);
  } catch {
    // The original error is the one that matters.
  }
}

/** In-memory API used by --dry-run and the tests. Files are keyed by path; head files by `<headSha>:<path>`. */
export class MemoryApi implements ClaApi {
  statuses: { sha: string; state: string; description: string }[] = [];
  comments: { id: number; body: string }[] = [];
  constructor(
    public pr: PrInfo,
    public commits: CommitInfo[],
    public files: PrFile[],
    public base: Record<string, string> = {},
    public head: Record<string, string> = {},
    public touching: Record<string, string[]> = {}
  ) {}
  async getPr(): Promise<PrInfo> {
    return this.pr;
  }
  async listCommits(): Promise<CommitInfo[]> {
    return this.commits;
  }
  async listFiles(): Promise<PrFile[]> {
    return this.files;
  }
  async readBase(path: string): Promise<string | null> {
    // The CLA text itself is read from this working tree unless a test supplies another.
    if (path === 'CLA.md' && !(path in this.base)) return fs.readFileSync(new URL('../../CLA.md', import.meta.url), 'utf8');
    return this.base[path] ?? null;
  }
  async listBaseDir(path: string): Promise<string[]> {
    const prefix = `${path}/`;
    return Object.keys(this.base)
      .filter((p) => p.startsWith(prefix) && !p.slice(prefix.length).includes('/'))
      .map((p) => p.slice(prefix.length));
  }
  async readHead(path: string): Promise<string | null> {
    return this.head[path] ?? null;
  }
  async commitsTouching(path: string): Promise<string[]> {
    return this.touching[path] ?? [];
  }
  async findBotComment(_pr: number, marker: string): Promise<{ id: number; body: string } | null> {
    return this.comments.find((c) => c.body.startsWith(marker)) ?? null;
  }
  async upsertBotComment(_pr: number, body: string, existingId: number | null): Promise<void> {
    const hit = this.comments.find((c) => c.id === existingId);
    if (hit) hit.body = body;
    else this.comments.push({ id: this.comments.length + 1, body });
  }
  async setStatus(sha: string, state: string, description: string): Promise<void> {
    this.statuses.push({ sha, state, description });
  }
  claUrl(): string {
    return 'https://example.invalid/CLA.md';
  }
}

function readPrNumber(): number {
  const name = process.env.GITHUB_EVENT_NAME;
  const path = process.env.GITHUB_EVENT_PATH;
  if (name !== 'pull_request_target') throw new Error(`unsupported event ${name ?? '(none)'}`);
  if (!path) throw new Error('GITHUB_EVENT_PATH is required');
  const n = JSON.parse(fs.readFileSync(path, 'utf8'))?.pull_request?.number;
  if (!Number.isInteger(n)) throw new Error('pull request number missing');
  return n;
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const config = validateConfig(JSON.parse(fs.readFileSync('.github/cla/config.json', 'utf8')));
  if (args[0] === '--dry-run') {
    const fx = JSON.parse(fs.readFileSync(args[1], 'utf8'));
    const api = new MemoryApi(fx.pr, fx.commits, fx.files ?? [], fx.base ?? {}, fx.head ?? {}, fx.touching ?? {});
    const result = await runCheck(api, config, fx.pr.number);
    console.log(JSON.stringify({ state: result.state, statuses: api.statuses, comments: api.comments }, null, 2));
    process.exit(result.state === 'success' ? 0 : 1);
  }
  const token = process.env.GITHUB_TOKEN;
  const repo = process.env.GITHUB_REPOSITORY;
  const defaultBranch = process.env.DEFAULT_BRANCH || 'main';
  if (!token || !repo) throw new Error('GITHUB_TOKEN and GITHUB_REPOSITORY are required');
  const api = new GitHubApi(repo, token, config.claPath, defaultBranch);
  const prNumber = readPrNumber();
  try {
    const result = await runCheck(api, config, prNumber);
    console.log(`CLA check: ${result.state}`);
    // A red check is a normal outcome for an unsigned pull request: the workflow itself succeeds.
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`CLA check failed closed: ${message}`);
    await reportError(api, config, prNumber, message);
    process.exit(1);
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((e) => {
    console.error(e instanceof Error ? e.message : e);
    process.exit(1);
  });
}
