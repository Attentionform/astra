import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, commitChange, getRepoDir, isTrackedPackageFile } from '@attentionform/astra-engine';
import { CLI_NAME, printStderr } from './output';

// Helpers shared by the history commands and by commands that record a revision when they write
// (batch). The CLI never creates the `.astra/` sidecar on its own: recording happens only when a
// sidecar repository already exists.

/** Commit-message cap for CLI-authored history commits (same cap the MCP history writer uses). */
const HISTORY_COMMIT_MESSAGE_MAX = 500;

export function capCommitMessage(message: string): string {
  return message.slice(0, HISTORY_COMMIT_MESSAGE_MAX);
}

/** Sidecar `.git` state: a real repo directory, a stray non-directory entry, or absent. */
export type SidecarGitState = 'repo' | 'stray' | 'absent';

/**
 * Classify the sidecar's `.git` entry. `stray` is a refusal state: a `.git`
 * FILE (e.g. a leftover `gitdir:` link) redirects the repository elsewhere;
 * isomorphic-git follows it, so writes would land outside the sidecar and
 * reads would loop back to the "no sidecar" gate — never treat it as absent.
 * Symlinks are followed like isomorphic-git does: a link to a real directory
 * is a usable repo; a link to a file or a broken link is stray.
 */
export function sidecarGitState(file: string): SidecarGitState {
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
export function hasHistorySidecar(file: string): boolean {
  return sidecarGitState(file) === 'repo';
}

/**
 * Tracked package files (everything but document.xml) from a live snapshot. `isTracked` selects the
 * rule: the default is the complete-snapshot rule; pass trackedPredicateFor(commit) to compare with a
 * commit recorded by an older ASTra version.
 */
export function trackedExtras(
  snapshot: Map<string, Buffer>,
  isTracked: (name: string) => boolean = isTrackedPackageFile,
): Map<string, Buffer> {
  const out = new Map<string, Buffer>();
  for (const [name, buf] of snapshot) {
    if (name !== 'word/document.xml' && isTracked(name)) out.set(name, buf);
  }
  return out;
}

export function extrasRecord(extras: Map<string, Buffer>): Record<string, Buffer> {
  return Object.fromEntries(extras);
}

export type BatchRevisionResult = { recorded: true; commit: string } | { recorded: false; reason: string } | null;

/**
 * Records the document just written by a command as a history revision, but only when the document
 * already has a history repository and the command wrote the document itself (not an `--out` copy).
 * Returns null when nothing was due. A failure is reported on stderr and in the result; the command
 * keeps its exit code, because the file was written.
 */
export async function recordBatchRevision(args: {
  file: string;
  target: string;
  parser: DocxParser;
  message: string;
  noHistory: boolean;
}): Promise<BatchRevisionResult> {
  if (args.noHistory) return null;
  if (path.resolve(args.target) !== path.resolve(args.file)) return null;
  if (sidecarGitState(args.file) !== 'repo') return null;
  try {
    const oid = await commitChange(
      args.file,
      args.parser.readDocumentXml(),
      capCommitMessage(args.message),
      extrasRecord(trackedExtras(args.parser.capturePackageSnapshot())),
    );
    return { recorded: true, commit: String(oid).slice(0, 7) };
  } catch (err: unknown) {
    const reason = err instanceof Error ? err.message : String(err);
    printStderr(`${CLI_NAME}: WARNING: the edit was saved but the history revision was not recorded: ${reason}`);
    return { recorded: false, reason };
  }
}
