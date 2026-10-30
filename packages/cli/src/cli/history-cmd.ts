import * as fs from 'fs';
import * as path from 'path';
import {
  DocxParser,
  checkoutCommit,
  commitChange,
  getEngineErrorCode,
  getHistory,
  getHistoryCommitCount,
  getRepoDir,
  initAstraRepo,
  isTrackedPackageFile,
  validateDocxBuffer,
} from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_OK, type GlobalOpts, print, printJson, printStderr } from './output';
import { failHuman, readErrRemedyProp } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { readDocx, writeDocx } from './io';

// ── Passive Git sidecar commands (history / preview-revert / revert / redo) ──
// The CLI never creates the `.astra/` sidecar on its own: reads refuse when it
// is absent (naming `history --init` as the explicit opt-in), and mutations
// additionally refuse when the working file was modified since the last
// recorded commit (a revert would silently overwrite unrecorded changes).
// All sidecar/git work reuses `@attentionform/astra-engine`'s git module — no second git
// layer here — and the commands stay stateless: open the file, operate,
// exit. Sidecars are never deleted.

type SidecarCommits = Awaited<ReturnType<typeof getHistory>>;
type CommitPackage = Awaited<ReturnType<typeof checkoutCommit>>;

/** Commit-message cap for CLI-authored history commits (same cap the MCP history writer uses). */
const HISTORY_COMMIT_MESSAGE_MAX = 500;
/** Human-output cap for file lists; --json keeps the full list. */
const HISTORY_FILES_SHOWN = 5;
/** A revert commit's subject carries the short oid of the state it restored. */
const REVERT_SUBJECT_RE = /^Revert to [0-9a-f]{7,40}\b/;

function shortOid(oid: string | null | undefined): string {
  return typeof oid === 'string' ? oid.slice(0, 7) : 'none';
}

function commitSubject(message: unknown): string {
  return String(message ?? '').split(/\r?\n/)[0].trim();
}

function capCommitMessage(message: string): string {
  return message.slice(0, HISTORY_COMMIT_MESSAGE_MAX);
}

function plural(count: number, one: string, many: string): string {
  return count === 1 ? one : many;
}

function shownList(names: readonly string[], cap = HISTORY_FILES_SHOWN): string {
  const shown = names.slice(0, cap).join(', ');
  const more = names.length > cap ? `; +${names.length - cap} more` : '';
  return `${shown}${more}`;
}

/** Sidecar `.git` state: a real repo directory, a stray non-directory entry, or absent. */
type SidecarGitState = 'repo' | 'stray' | 'absent';

/**
 * Classify the sidecar's `.git` entry. `stray` is a refusal state: a `.git`
 * FILE (e.g. a leftover `gitdir:` link) redirects the repository elsewhere;
 * isomorphic-git follows it, so writes would land outside the sidecar and
 * reads would loop back to the "no sidecar" gate — never treat it as absent.
 * Symlinks are followed like isomorphic-git does: a link to a real directory
 * is a usable repo; a link to a file or a broken link is stray.
 */
function sidecarGitState(file: string): SidecarGitState {
  const gitPath = path.join(getRepoDir(file), '.git');
  let entry: fs.Stats;
  try {
    entry = fs.lstatSync(gitPath);
  } catch {
    return 'absent';
  }
  if (entry.isSymbolicLink()) {
    try {
      return fs.statSync(gitPath).isDirectory() ? 'repo' : 'stray';
    } catch {
      return 'stray';
    }
  }
  return entry.isDirectory() ? 'repo' : 'stray';
}

/** Is a valid sidecar chain present for this document? */
function hasHistorySidecar(file: string): boolean {
  return sidecarGitState(file) === 'repo';
}

/** Refusal for a `.git`-not-a-directory sidecar (stray file, not a repo). */
function strayGitEntryError(cmd: string, file: string): Error {
  return Object.assign(
    new Error(
      `${cmd}: the history sidecar for "${file}" has a ".git" entry that is a file, not a directory — ` +
        'refusing to read or write it (a stray .git file can redirect the repository outside the sidecar)',
    ),
    {
      code: 'E_IO' as const,
      remedy:
        `remove the stray ".git" file inside the ".astra/<basename>/" sidecar directory, then recreate the ` +
        `sidecar with \`${CLI_NAME} history <docx> --init\`, then retry`,
    },
  );
}

/**
 * Refuse when no sidecar exists: a usage error (exit 2) naming the explicit
 * `history --init` opt-in — never a silent init, never a fake empty success.
 * A stray `.git` file (see sidecarGitState) is an operational refusal instead:
 * the sidecar path exists but is not a usable repository.
 */
function sidecarGate(out: GlobalOpts, cmd: string, file: string): number | null {
  const state = sidecarGitState(file);
  if (state === 'repo') return null;
  if (state === 'stray') {
    return failHuman(out, strayGitEntryError(cmd, file), { file, fallbackCode: 'E_IO' });
  }
  return usage(
    out,
    `${cmd}: no history sidecar exists for "${file}" — nothing has been recorded yet. ` +
      `Remedy: create it explicitly with \`${CLI_NAME} history <docx> --init\` (alias --track; ` +
      `opt-in checkpointing — ordinary commands never create it), then retry`,
  );
}

/** Tracked package files (everything but document.xml) from a live snapshot. */
function trackedExtras(snapshot: Map<string, Buffer>): Map<string, Buffer> {
  const out = new Map<string, Buffer>();
  for (const [name, buf] of snapshot) {
    if (name !== 'word/document.xml' && isTrackedPackageFile(name)) out.set(name, buf);
  }
  return out;
}

function extrasRecord(extras: Map<string, Buffer>): Record<string, Buffer> {
  return Object.fromEntries(extras);
}

/** Byte-diff two package views; word/document.xml first, then sorted extras. */
function packageDiff(
  liveXml: string,
  liveExtras: Map<string, Buffer>,
  otherXml: string,
  otherExtras: Map<string, Buffer>,
): string[] {
  const out: string[] = [];
  if (!Buffer.from(liveXml, 'utf8').equals(Buffer.from(otherXml, 'utf8'))) out.push('word/document.xml');
  const names = new Set<string>([...liveExtras.keys(), ...otherExtras.keys()]);
  for (const name of [...names].sort()) {
    const a = liveExtras.get(name);
    const b = otherExtras.get(name);
    if (!a || !b || !a.equals(b)) out.push(name);
  }
  return out;
}

/** Package files that changed between the working file and a recorded commit. */
async function divergedFiles(file: string, parser: DocxParser, commitOid: string): Promise<string[]> {
  const recorded = await checkoutCommit(file, commitOid);
  const snapshot = parser.capturePackageSnapshot();
  const liveXml = snapshot.get('word/document.xml')?.toString('utf8') ?? '';
  return packageDiff(liveXml, trackedExtras(snapshot), recorded.xml, recorded.extra);
}

function sidecarWarning(commits: SidecarCommits): string | null {
  const carried = (commits as unknown as { warning?: unknown }).warning;
  return typeof carried === 'string' && carried.length > 0 ? carried : null;
}

/** Read the sidecar commit list; engine read failures surface as E_IO. */
async function readCommits(cmd: string, file: string): Promise<SidecarCommits> {
  try {
    return await getHistory(file);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    // A 0-commit (unborn) chain cannot be logged. An interrupted init, a
    // deleted ref, or corruption all look the same from here (no commits
    // reachable from HEAD), so the generic "directory is readable" hint
    // misleads. Point at the explicit repair: `history --init` records the
    // baseline over the empty chain.
    const empty = (await getHistoryCommitCount(getRepoDir(file))) === 0;
    throw Object.assign(new Error(`${cmd}: cannot read the recorded history for "${file}": ${msg}`), {
      code: 'E_IO' as const,
      remedy: empty
        ? `the sidecar has no commits reachable from HEAD (an interrupted init, a deleted ref, or corruption looks the same) — repair it with \`${CLI_NAME} history <docx> --init\`, then retry`
        : 'check the .astra/ sidecar directory is readable, then retry',
    });
  }
}

function conflictError(cmd: string, file: string, headOid: string, changed: string[]): Error {
  return Object.assign(
    new Error(
      `${cmd}: "${file}" was modified since the last recorded commit ${shortOid(headOid)} — ` +
        `refusing to overwrite changes that were never recorded ` +
        `(${changed.length} changed package ${plural(changed.length, 'file', 'files')}: ${shownList(changed)})`,
    ),
    {
      code: 'E_CONFLICT' as const,
      remedy:
        'record the current state first via the MCP save_document tool, or manually commit it into ' +
        'the .astra/ sidecar (a standard git repo); then retry, or copy the unrecorded changes aside first',
    },
  );
}

function stepsDepthError(cmd: string, steps: number, depth: number): Error {
  return Object.assign(
    new Error(`${cmd}: --steps ${steps} exceeds the recorded history depth (${depth} recorded ${plural(depth, 'commit', 'commits')})`),
    {
      code: 'E_NOT_FOUND' as const,
      remedy:
        depth <= 1
          ? `only ${depth} recorded commit exists — a revert needs at least two (a baseline plus one recorded change)`
          : `retry with --steps < ${depth}, or browse the recorded commits with \`${CLI_NAME} history <docx>\``,
    },
  );
}

function nothingToRedoError(): Error {
  return Object.assign(
    new Error('redo: nothing to redo — the latest recorded commit is not a revert'),
    {
      code: 'E_NOT_FOUND' as const,
      remedy: `run \`${CLI_NAME} revert <docx> --steps 1\` first to create a redoable state, then retry`,
    },
  );
}

interface ParsedDocSteps {
  file: string;
  steps: number;
}

/**
 * A token that looks like a git revision: HEAD / HEAD~n / HEAD^n, a reflog
 * expression (`@{1}`), an ancestry shortcut (`~`/`~n`), or a full 40-hex oid.
 * Users pass one as a second positional expecting commit-oid targeting.
 */
function isGitRefLike(token: string): boolean {
  return /^(?:head(?:[~^]\d*|@\{[^}]*\})*|@\{[^}]*\}|~-?\d*|[0-9a-f]{40})$/i.test(token);
}

/**
 * Remedy for a git-ref-style second positional: name the real contract.
 * History is depth-addressed here (`--steps <n>`), commit-oid targeting is
 * MCP-only (`revert_document` / `preview_revert`; redo only moves forward by
 * depth), and `history` browses the recorded commits.
 */
function gitRefPositionalRemedy(cmd: string, token: string): string {
  const oidTool = cmd === 'preview-revert' ? 'preview_revert' : cmd === 'revert' ? 'revert_document' : undefined;
  return (
    `${cmd}: "${token}" looks like a git revision, but the recorded history is addressed by depth, not by revision. ` +
    `Remedy: use \`--steps <n>\` (n recorded commits ${cmd === 'redo' ? 'forward' : 'back'}; 1 = the ${cmd === 'redo' ? 'next' : 'previous'} recorded state) and browse commits with ` +
    `\`${CLI_NAME} history <docx>\`` +
    (oidTool ? `; targeting a specific commit oid is available only over MCP (${oidTool})` : '')
  );
}

/** Parse `<docx> [--steps N]`; returns a usage exit code instead when invalid. */
function parseDocStepsArgs(out: GlobalOpts, cmd: string, args: string[]): ParsedDocSteps | number {
  let file: string | undefined;
  let stepsRaw: string | undefined;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--steps') {
      const value = args[++i];
      if (value === undefined) return usage(out, `${cmd} --steps requires a value`);
      stepsRaw = value;
    } else if (a.startsWith('--steps=')) {
      stepsRaw = a.slice('--steps='.length);
    } else if (a.startsWith('--')) {
      return usage(out, `${cmd}: unknown flag "${a}"`);
    } else if (!file) {
      file = a;
    } else if (isGitRefLike(a)) {
      return usage(out, gitRefPositionalRemedy(cmd, a));
    } else {
      return usage(out, `${cmd} takes exactly one <docx> argument`);
    }
  }
  if (!file) return usage(out, `${cmd} requires <docx>`);
  let steps = 1;
  if (stepsRaw !== undefined) {
    const n = Number(stepsRaw);
    if (!Number.isInteger(n) || n < 1) return usage(out, `${cmd} --steps must be an integer >= 1 (got "${stepsRaw}")`);
    steps = n;
  }
  return { file, steps };
}

interface RestoreOptions {
  cmd: 'revert' | 'redo';
  file: string;
  parser: DocxParser;
  /** Exact working-file bytes read before the command: the rollback source. */
  prevBytes: Buffer;
  target: CommitPackage;
  targetOid: string;
  commitMessage: string;
}

/**
 * Shared mutation core for revert/redo: merge the target commit's tracked
 * package files into the live package, rezip, gate the candidate, write the
 * file, then record the restored state as a new head commit (append-only:
 * refs never move). A commit failure restores the pre-command bytes.
 */
async function restoreAndCommit(opts: RestoreOptions): Promise<{ commitOid: string; commitMessage: string }> {
  const { cmd, file, parser, prevBytes, target, targetOid, commitMessage } = opts;
  const snapshot = parser.capturePackageSnapshot();
  for (const key of [...snapshot.keys()]) {
    if (key !== 'word/document.xml' && isTrackedPackageFile(key)) snapshot.delete(key);
  }
  for (const [name, data] of target.extra) snapshot.set(name, data);
  snapshot.set('word/document.xml', Buffer.from(target.xml, 'utf8'));
  parser.restorePackageSnapshot(snapshot);
  const buf = parser.rezip();
  const candidate = await validateDocxBuffer(buf);
  if (!candidate.ok) {
    const first = candidate.errors[0];
    const detail = first?.description ? `: ${first.description}` : '';
    throw Object.assign(
      new Error(`${cmd}: the package recorded at commit ${shortOid(targetOid)} failed validation${detail}`),
      {
        code: 'E_VALIDATION' as const,
        remedy: `browse commits with \`${CLI_NAME} history <docx>\` and restore a different one, or re-export the file from Word and retry`,
      },
    );
  }
  try {
    await writeDocx(file, buf);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw Object.assign(new Error(`${cmd}: write failed: ${msg}`), {
      code: 'E_IO' as const,
      remedy: 'check the document directory is writable and retry',
    });
  }
  const restoredXml = parser.readDocumentXml();
  const extras = extrasRecord(trackedExtras(parser.capturePackageSnapshot()));
  try {
    // Race guard: the sidecar must still exist at commit time — a vanished
    // chain must refuse, never be silently re-created by the commit path.
    if (!hasHistorySidecar(file)) {
      throw Object.assign(new Error(`${cmd}: the history sidecar disappeared before the commit`), {
        code: 'E_NOT_FOUND' as const,
        remedy: `create it explicitly with \`${CLI_NAME} history <docx> --init\`, then retry`,
      });
    }
    const commitOid = await commitChange(file, restoredXml, commitMessage, extras);
    return { commitOid, commitMessage };
  } catch (err: unknown) {
    let rollbackOk = true;
    try {
      await writeDocx(file, prevBytes);
    } catch {
      rollbackOk = false;
    }
    const code = getEngineErrorCode(err) ?? ('E_IO' as const);
    const msg = err instanceof Error ? err.message : String(err);
    const baseRemedy = readErrRemedyProp(err) ?? `retry \`${CLI_NAME} ${cmd} <docx>\` to re-record history`;
    throw Object.assign(
      new Error(
        `${cmd}: history commit failed; working file ${rollbackOk ? 'restored to the pre-command bytes' : 'could NOT be rolled back'}: ${msg}`,
      ),
      {
        code,
        remedy: rollbackOk
          ? baseRemedy
          : `${baseRemedy}; also restore "${file}" from a backup — the rollback write failed`,
      },
    );
  }
}

// ── history ─────────────────────────────────────────────────────────────────

async function initHistorySidecar(out: GlobalOpts, file: string): Promise<number> {
  const buf = await readDocx(file);
  const parser = new DocxParser(buf);
  const xml = parser.readDocumentXml();
  const extras = extrasRecord(trackedExtras(parser.capturePackageSnapshot()));
  try {
    await initAstraRepo(file, xml, extras);
  } catch (err: unknown) {
    // Mirror revert's write wrap: a sidecar create/write refusal (e.g. EACCES
    // on a read-only document directory) is an IO failure, not a document
    // validation one — keep engine codes that ride along (E_SANDBOX/E_LIMIT),
    // otherwise E_IO with the writable-directory remedy.
    const code = getEngineErrorCode(err) ?? ('E_IO' as const);
    const msg = err instanceof Error ? err.message : String(err);
    throw Object.assign(new Error(`history --init: cannot record the baseline for "${file}": ${msg}`), {
      code,
      remedy: readErrRemedyProp(err) ?? 'check the document directory is writable and retry',
    });
  }
  const commits = await readCommits('history', file);
  const warning = sidecarWarning(commits);
  if (warning) printStderr(`${CLI_NAME}: warning: ${warning}`);
  const head = commits[0]?.oid ?? null;
  print(
    out,
    `HISTORY INIT ${file}: recorded baseline ${shortOid(head)} (${Object.keys(extras).length} tracked package ` +
      `${plural(Object.keys(extras).length, 'file', 'files')}); ${commits.length} ${plural(commits.length, 'commit', 'commits')}.`,
  );
  printJson(out, {
    ok: true,
    command: 'history',
    file,
    initialized: true,
    commitOid: head,
    totalCommits: commits.length,
    currentCommit: head,
    trackedFiles: Object.keys(extras).length,
    warning,
  });
  return EXIT_OK;
}

export async function cmdHistory(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'history');
  let file: string | undefined;
  let init = false;
  for (const a of args) {
    if (a === '--init' || a === '--track') init = true;
    else if (a.startsWith('--')) return usage(out, `history: unknown flag "${a}"`);
    else if (!file) file = a;
    else return usage(out, 'history takes exactly one <docx> argument');
  }
  if (!file) return usage(out, 'history requires <docx>');
  try {
    if (init) {
      const state = sidecarGitState(file);
      // A stray `.git` file is refused (never followed/replaced); only a
      // truly absent sidecar is created.
      if (state === 'stray') {
        return failHuman(out, strayGitEntryError('history', file), { file, fallbackCode: 'E_IO' });
      }
      if (state !== 'repo') return await initHistorySidecar(out, file);
      // Kill-mid-init repair: an existing sidecar with 0 commits (unborn HEAD)
      // is not a usable history — `--init` records the baseline over it instead
      // of dead-ending every command on the unreadable chain.
      if ((await getHistoryCommitCount(getRepoDir(file))) === 0) {
        return await initHistorySidecar(out, file);
      }
    }
    const gate = sidecarGate(out, 'history', file);
    if (gate !== null) return gate;
    const commits = await readCommits('history', file);
    const warning = sidecarWarning(commits);
    if (warning) printStderr(`${CLI_NAME}: warning: ${warning}`);
    const head = commits[0]?.oid ?? null;
    if (init) {
      print(
        out,
        `HISTORY INIT ${file}: history already initialized (${commits.length} ${plural(commits.length, 'commit', 'commits')}, ` +
          `head ${shortOid(head)}); nothing to do.`,
      );
    }
    print(out, `HISTORY ${file}: ${commits.length} ${plural(commits.length, 'commit', 'commits')}, head ${shortOid(head)}`);
    for (const c of commits) {
      print(out, `  ${shortOid(c.oid)} ${c.timestamp.toISOString()} ${c.author}: ${commitSubject(c.message)}`);
    }
    printJson(out, {
      ok: true,
      command: 'history',
      file,
      ...(init ? { initialized: false } : {}),
      totalCommits: commits.length,
      currentCommit: head,
      history: commits.map((c) => ({
        oid: c.oid,
        message: commitSubject(c.message),
        author: c.author,
        timestamp: c.timestamp.toISOString(),
        timestampMs: c.timestamp.getTime(),
      })),
      warning,
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}

// ── preview-revert (read-only) ──────────────────────────────────────────────

export async function cmdPreviewRevert(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'preview-revert');
  const parsed = parseDocStepsArgs(out, 'preview-revert', args);
  if (typeof parsed === 'number') return parsed;
  const { file, steps } = parsed;
  const gate = sidecarGate(out, 'preview-revert', file);
  if (gate !== null) return gate;
  try {
    const buf = await readDocx(file);
    const parser = new DocxParser(buf);
    const commits = await readCommits('preview-revert', file);
    if (steps >= commits.length) throw stepsDepthError('preview-revert', steps, commits.length);
    const targetOid = commits[steps].oid;
    const target = await checkoutCommit(file, targetOid);
    const snapshot = parser.capturePackageSnapshot();
    const liveXml = snapshot.get('word/document.xml')?.toString('utf8') ?? '';
    const affectedFiles = packageDiff(liveXml, trackedExtras(snapshot), target.xml, target.extra);
    const changed = await divergedFiles(file, parser, commits[0].oid);
    const warning = sidecarWarning(commits);
    const head = commits[0].oid;
    const subject = commitSubject(commits[steps].message);
    print(
      out,
      `PREVIEW-REVERT ${file}: steps=${steps} -> ${shortOid(targetOid)} "${subject}" ` +
        `(${commits[steps].author}; ${commits.length} recorded ${plural(commits.length, 'commit', 'commits')})`,
    );
    print(
      out,
      `  affected package files (${affectedFiles.length}): ${affectedFiles.length > 0 ? shownList(affectedFiles) : 'none — the target state matches the working file'}`,
    );
    print(
      out,
      changed.length > 0
        ? `  working file: DIVERGED from head ${shortOid(head)} — a real revert would refuse E_CONFLICT (changed: ${shownList(changed)})`
        : `  working file: clean at head ${shortOid(head)} — a real revert would proceed`,
    );
    if (warning) printStderr(`${CLI_NAME}: warning: ${warning}`);
    printJson(out, {
      ok: true,
      command: 'preview-revert',
      file,
      targetOid,
      targetMessage: subject,
      targetAuthor: commits[steps].author,
      steps,
      affectedFiles,
      currentHeadOid: head,
      totalCommits: commits.length,
      diverged: changed.length > 0,
      divergedFiles: changed,
      warning,
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}

// ── revert / redo (conditional mutations) ───────────────────────────────────

export async function cmdRevert(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'revert');
  const parsed = parseDocStepsArgs(out, 'revert', args);
  if (typeof parsed === 'number') return parsed;
  const { file, steps } = parsed;
  const gate = sidecarGate(out, 'revert', file);
  if (gate !== null) return gate;
  try {
    const buf = await readDocx(file);
    const parser = new DocxParser(buf);
    const commits = await readCommits('revert', file);
    if (steps >= commits.length) throw stepsDepthError('revert', steps, commits.length);
    const headOid = commits[0].oid;
    const changed = await divergedFiles(file, parser, headOid);
    if (changed.length > 0) throw conflictError('revert', file, headOid, changed);
    const targetOid = commits[steps].oid;
    const target = await checkoutCommit(file, targetOid);
    const subject = commitSubject(commits[steps].message);
    const commitMessage = capCommitMessage(`Revert to ${shortOid(targetOid)}: "${subject}"`);
    const result = await restoreAndCommit({ cmd: 'revert', file, parser, prevBytes: buf, target, targetOid, commitMessage });
    print(
      out,
      `REVERT OK ${file}: restored ${shortOid(targetOid)} "${subject}" (steps=${steps}); new commit ${shortOid(result.commitOid)}`,
    );
    printJson(out, {
      ok: true,
      command: 'revert',
      file,
      steps,
      revertedToCommit: targetOid,
      commitOid: result.commitOid,
      commitMessage: result.commitMessage,
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}

export async function cmdRedo(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'redo');
  let file: string | undefined;
  for (const a of args) {
    if (a.startsWith('--')) return usage(out, `redo: unknown flag "${a}"`);
    else if (!file) file = a;
    else return usage(out, 'redo takes exactly one <docx> argument');
  }
  if (!file) return usage(out, 'redo requires <docx>');
  const gate = sidecarGate(out, 'redo', file);
  if (gate !== null) return gate;
  try {
    const buf = await readDocx(file);
    const parser = new DocxParser(buf);
    const commits = await readCommits('redo', file);
    const head = commits[0];
    // Stateless redo: the redoable state is the parent (pre-revert state) of a
    // revert head. Each redo records a new commit, so it consumes the revert.
    if (!head || commits.length < 2 || !REVERT_SUBJECT_RE.test(commitSubject(head.message))) {
      throw nothingToRedoError();
    }
    const headOid = head.oid;
    const changed = await divergedFiles(file, parser, headOid);
    if (changed.length > 0) throw conflictError('redo', file, headOid, changed);
    const targetOid = commits[1].oid;
    const target = await checkoutCommit(file, targetOid);
    const subject = commitSubject(commits[1].message);
    const commitMessage = capCommitMessage(`Redo: restore ${shortOid(targetOid)} "${subject}"`);
    const result = await restoreAndCommit({ cmd: 'redo', file, parser, prevBytes: buf, target, targetOid, commitMessage });
    print(
      out,
      `REDO OK ${file}: re-applied the state replaced by revert ${shortOid(headOid)}; new commit ${shortOid(result.commitOid)}`,
    );
    printJson(out, {
      ok: true,
      command: 'redo',
      file,
      restoredCommit: targetOid,
      commitOid: result.commitOid,
      commitMessage: result.commitMessage,
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}
