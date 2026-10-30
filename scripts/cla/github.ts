/**
 * Thin GitHub API layer for the CLA check. No decisions are made here.
 * Uses Node's built-in fetch and the workflow token; no dependencies.
 * Everything is read-only except the commit status and the one comment.
 */
/* eslint-disable @typescript-eslint/no-explicit-any -- GitHub API responses are untyped JSON that is validated field by field below. */
import type { CommitInfo, Person, PrFile } from './evaluate';

export interface PrInfo {
  number: number;
  headSha: string;
  state: string;
}

export interface ClaApi {
  getPr(pr: number): Promise<PrInfo>;
  listCommits(pr: number): Promise<CommitInfo[]>;
  listFiles(pr: number): Promise<PrFile[]>;
  /** Text of a file on the default branch; null when it does not exist. */
  readBase(path: string): Promise<string | null>;
  /** Names of the files directly inside a directory on the default branch; empty when it does not exist. */
  listBaseDir(path: string): Promise<string[]>;
  /** Text of a file at the pull request's head commit; null when it does not exist. */
  readHead(path: string, headSha: string): Promise<string | null>;
  /** Shas of the commits reachable from `headSha` that changed `path`. */
  commitsTouching(path: string, headSha: string): Promise<string[]>;
  findBotComment(pr: number, marker: string): Promise<{ id: number; body: string } | null>;
  upsertBotComment(pr: number, body: string, existingId: number | null): Promise<void>;
  setStatus(sha: string, state: 'success' | 'failure' | 'error' | 'pending', description: string, context: string): Promise<void>;
  claUrl(): string;
}

interface ApiUser {
  login: string;
  id: number;
  type?: string;
}

function person(u: ApiUser | null | undefined): Person {
  return u && typeof u.login === 'string' && Number.isInteger(u.id) ? { login: u.login, id: u.id, type: u.type } : { login: null, id: null };
}

const encodePath = (p: string): string => p.split('/').map(encodeURIComponent).join('/');

export class GitHubApi implements ClaApi {
  constructor(
    private readonly repo: string,
    private readonly token: string,
    private readonly claPath: string,
    private readonly defaultBranch: string
  ) {
    if (!/^[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/.test(repo)) throw new Error('GITHUB_REPOSITORY is not owner/name');
  }

  private async call(method: string, url: string, body?: unknown, okStatuses: number[] = []): Promise<{ status: number; data: any }> {
    const res = await fetch(`https://api.github.com${url}`, {
      method,
      headers: {
        Authorization: `Bearer ${this.token}`,
        Accept: 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'astra-cla-check',
        ...(body === undefined ? {} : { 'Content-Type': 'application/json' }),
      },
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    const text = await res.text();
    let data: any = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch {
      data = null;
    }
    if (!res.ok && !okStatuses.includes(res.status)) {
      throw new Error(`GitHub API ${method} ${url.split('?')[0]} failed with ${res.status}`);
    }
    return { status: res.status, data };
  }

  private async readAt(path: string, ref: string): Promise<string | null> {
    const { status, data } = await this.call('GET', `/repos/${this.repo}/contents/${encodePath(path)}?ref=${encodeURIComponent(ref)}`, undefined, [404]);
    if (status === 404) return null;
    if (Array.isArray(data) || data?.type !== 'file' || data?.encoding !== 'base64' || typeof data.content !== 'string') {
      throw new Error(`${path} could not be read as a file`);
    }
    return Buffer.from(data.content, 'base64').toString('utf8');
  }

  async getPr(pr: number): Promise<PrInfo> {
    const { data } = await this.call('GET', `/repos/${this.repo}/pulls/${pr}`);
    if (!data?.head?.sha) throw new Error('pull request has no head sha');
    return { number: pr, headSha: data.head.sha, state: data.state };
  }

  async listCommits(pr: number): Promise<CommitInfo[]> {
    const out: CommitInfo[] = [];
    for (let page = 1; page <= 4; page++) {
      const { data } = await this.call('GET', `/repos/${this.repo}/pulls/${pr}/commits?per_page=100&page=${page}`);
      if (!Array.isArray(data)) throw new Error('unexpected commits response');
      for (const c of data) {
        out.push({
          sha: String(c.sha),
          author: person(c.author),
          committer: person(c.committer),
          authorEmail: String(c.commit?.author?.email ?? ''),
          parentCount: Array.isArray(c.parents) ? c.parents.length : 1,
          message: String(c.commit?.message ?? ''),
        });
      }
      if (data.length < 100) return out;
    }
    throw new Error('pull request has more commits than the check can read');
  }

  async listFiles(pr: number): Promise<PrFile[]> {
    const out: PrFile[] = [];
    for (let page = 1; page <= 30; page++) {
      const { data } = await this.call('GET', `/repos/${this.repo}/pulls/${pr}/files?per_page=100&page=${page}`);
      if (!Array.isArray(data)) throw new Error('unexpected files response');
      for (const f of data) {
        if (typeof f.filename !== 'string' || typeof f.status !== 'string') throw new Error('unexpected file entry');
        out.push({ path: f.filename, status: f.status, ...(typeof f.previous_filename === 'string' ? { previousPath: f.previous_filename } : {}) });
      }
      if (data.length < 100) return out;
    }
    // GitHub lists at most 3000 files; a list that long cannot be checked completely.
    throw new Error('pull request changes too many files for the check to read');
  }

  readBase(path: string): Promise<string | null> {
    return this.readAt(path, this.defaultBranch);
  }

  async listBaseDir(path: string): Promise<string[]> {
    const { status, data } = await this.call('GET', `/repos/${this.repo}/contents/${encodePath(path)}?ref=${encodeURIComponent(this.defaultBranch)}`, undefined, [404]);
    if (status === 404) return [];
    if (!Array.isArray(data)) throw new Error(`${path} is not a directory`);
    return data.filter((e: any) => e?.type === 'file' && typeof e.name === 'string').map((e: any) => e.name as string);
  }

  readHead(path: string, headSha: string): Promise<string | null> {
    return this.readAt(path, headSha);
  }

  async commitsTouching(path: string, headSha: string): Promise<string[]> {
    const out: string[] = [];
    for (let page = 1; page <= 4; page++) {
      const { data } = await this.call('GET', `/repos/${this.repo}/commits?sha=${encodeURIComponent(headSha)}&path=${encodeURIComponent(path)}&per_page=100&page=${page}`);
      if (!Array.isArray(data)) throw new Error('unexpected commits response');
      for (const c of data) out.push(String(c.sha));
      if (data.length < 100) return out;
    }
    throw new Error(`too many commits touch ${path}`);
  }

  async findBotComment(pr: number, marker: string): Promise<{ id: number; body: string } | null> {
    for (let page = 1; page <= 10; page++) {
      const { data } = await this.call('GET', `/repos/${this.repo}/issues/${pr}/comments?per_page=100&page=${page}`);
      if (!Array.isArray(data)) throw new Error('unexpected comments response');
      const hit = data.find((c: any) => typeof c.body === 'string' && c.body.startsWith(marker) && c.user?.type === 'Bot');
      if (hit) return { id: hit.id, body: hit.body };
      if (data.length < 100) return null;
    }
    return null;
  }

  async upsertBotComment(pr: number, body: string, existingId: number | null): Promise<void> {
    if (existingId) await this.call('PATCH', `/repos/${this.repo}/issues/comments/${existingId}`, { body });
    else await this.call('POST', `/repos/${this.repo}/issues/${pr}/comments`, { body });
  }

  async setStatus(sha: string, state: 'success' | 'failure' | 'error' | 'pending', description: string, context: string): Promise<void> {
    await this.call('POST', `/repos/${this.repo}/statuses/${sha}`, { state, description: description.slice(0, 140), context });
  }

  claUrl(): string {
    return `https://github.com/${this.repo}/blob/${this.defaultBranch}/${this.claPath}`;
  }
}
