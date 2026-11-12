import { BASELINE_MAX_ERRORS, DocxParser, commitStateIfDiffersFromHead, repairDocxBuffer } from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_OK, type GlobalOpts, applyNormalizerForDisplay, formatNormalizedInline, previewNormalizedForDisplay, print, printJson, printStderr, type NormalizedRewrite } from './output';
import { classifyFileReadError, failHuman, failWithCode, inferBatchCode, withGatedEngineOutput } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import * as path from 'path';
import { macroFlagsUsageError, macroOptsFromFlags, macroOutputRefusal, macroStripNameRefusal, readDocx, writeDocx } from './io';
import { judgeGate, refuses, introducedText, toleratedFields, toleratedNote, type ToleratedVerdict } from './baseline-gate';
import { extrasRecord, hasHistorySidecar, recordBatchRevision, refuseWhenHistoryBusy, trackedExtras } from './history-support';
import { type BackupState, ensureBackupCopy, resolveBackupPath, truncLine } from './batch-cmd';

/**
 * Repair-as-action entry for legacy files: runs the engine allowlist
 * normalizer (display-only against the core API — see the surface block at
 * the top of this file), then the strict validation gate, then writes.
 * Chosen over a `validate --fix` alias: `validate` is pinned read-only
 * fail-closed by scripts, so a mutating alias would break its contract.
 */
export async function cmdNormalize(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'normalize');
  let file: string | undefined;
  let outPath: string | undefined;
  let dryRun = false;
  let strict = false;
  let noHistory = false;
  let backupGiven = false;
  let backupRaw: string | undefined;
  let allowMacros = false;
  let stripMacros = false;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--out') {
      const next = args[++i];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'normalize --out requires a value (output .docx path)');
      }
      outPath = next;
    } else if (a.startsWith('--out=')) {
      const v = a.slice('--out='.length);
      if (v.length === 0) return usage(out, 'normalize --out requires a value (output .docx path)');
      outPath = v;
    } else if (a === '--dry-run') dryRun = true;
    else if (a === '--strict') strict = true;
    else if (a === '--no-history') noHistory = true;
    else if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a === '--backup') {
      backupGiven = true;
      const nextBackup = args[i + 1];
      if (nextBackup !== undefined && nextBackup.length > 0 && !nextBackup.startsWith('-')) backupRaw = args[++i];
    } else if (a.startsWith('--backup=')) {
      backupGiven = true;
      backupRaw = a.slice('--backup='.length);
      if (backupRaw.length === 0) {
        return usage(out, 'normalize --backup requires a path value or bare --backup for the default <input>.bak.docx');
      }
    } else if (a.startsWith('--')) return usage(out, `normalize: unknown flag "${a}"`);
    else if (!file) file = a;
    else return usage(out, 'normalize takes exactly one <docx> argument');
  }
  if (!file) return usage(out, 'normalize requires exactly one <docx> argument');
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  const macroOpts = macroOptsFromFlags({ allowMacros, stripMacros });
  const target = outPath ?? file;
  if (stripMacros && !dryRun) {
    const nameRefusal = macroStripNameRefusal(
      'normalize',
      target,
      `write the result to a .docx with --out <name>.docx, or use --allow-macros to keep the macro parts in the .docm`,
    );
    if (nameRefusal !== null) return usage(out, nameRefusal);
  }
  const backupState: BackupState = {
    path: backupGiven && !dryRun ? resolveBackupPath(file, backupRaw) : undefined,
    done: false,
  };
  if (backupGiven && dryRun) {
    printStderr(`${CLI_NAME}: note: --backup skipped: --dry-run writes nothing`);
  }
  let buf: Buffer;
  try {
    buf = await readDocx(file);
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_IO' });
  }
  if (dryRun) {
    let wouldBe: NormalizedRewrite[];
    try {
      wouldBe = previewNormalizedForDisplay(buf, macroOpts);
    } catch (err: unknown) {
      const read = classifyFileReadError(err, file);
      if (read !== undefined) return failWithCode(out, read.message, read.code, undefined, read.remedy, err);
      throw err;
    }
    print(out, `DRY-RUN OK ${file} -> ${target} (would-be ${formatNormalizedInline(wouldBe)}, writes nothing)`);
    for (const n of wouldBe) print(out, `would-be normalize: ${truncLine(n, 500)}`);
    printJson(out, { ok: true, dryRun: true, file, output: target, normalized: wouldBe });
    return EXIT_OK;
  }
  const base = applyNormalizerForDisplay(buf, macroOpts);
  let buffer = base.buffer;
  let normalized: string[] = base.normalized;
  let remaining = -1;
  let rolledBack: { rule: string; reason: string }[] = [];
  try {
    // Repair rules: each kept rule adds no validator error and keeps the visible text.
    const repair = await withGatedEngineOutput(out, () => repairDocxBuffer(buffer));
    buffer = repair.buffer;
    normalized = [...normalized, ...repair.normalized];
    remaining = repair.remaining;
    rolledBack = repair.rolledBack;
  } catch {
    /* a package the repair cannot read keeps the plain normalizer result; the gate below judges it */
  }
  let tolerated: ToleratedVerdict | undefined;
  try {
    const gate = new DocxParser(buffer, macroOpts);
    const nameRefusal = macroOutputRefusal('normalize', target, gate.getMacroParts(), stripMacros);
    if (nameRefusal !== null) return usage(out, nameRefusal);
    const validation = await withGatedEngineOutput(out, () => gate.validateOOXML({ maxErrors: BASELINE_MAX_ERRORS }));
    const verdict = await judgeGate(validation, strict, () => buf);
    if (verdict.kind === 'tolerated') tolerated = verdict;
    if (refuses(verdict)) {
      const first = validation.errors[0] as
        | { description?: unknown; message?: unknown; xpath?: unknown; xPath?: unknown; path?: unknown }
        | undefined;
      const firstDesc =
        (typeof first?.description === 'string' && first.description) ||
        (typeof first?.message === 'string' && first.message) ||
        (first ? JSON.stringify(first) : 'OOXML invalid');
      return failWithCode(
        out,
        verdict.kind === 'introduced'
          ? `normalize applied but ${introducedText(verdict, String(firstDesc).slice(0, 500))}`
          : `normalize applied but validation still fails, nothing written: ${String(firstDesc).slice(0, 500)}.`,
        'E_VALIDATION',
        undefined,
        'fix the flagged content in Word and retry',
        JSON.stringify(validation.errors.slice(0, 3))
      );
    }
  } catch (err: unknown) {
    const read = classifyFileReadError(err, file);
    if (read !== undefined) return failWithCode(out, read.message, read.code, undefined, read.remedy, err);
    // The macro refusal carries its own remedy (the two flags); keep it instead of the generic one below.
    if (err instanceof Error && err.message.includes('macro/active content parts')) {
      return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
    }
    return failWithCode(
      out,
      err instanceof Error ? err.message : String(err),
      inferBatchCode(err, 'E_VALIDATION'),
      undefined,
      'pick a valid .docx and retry. Nothing was written',
      err
    );
  }
  if (normalized.length > 0) {
    const busy = await refuseWhenHistoryBusy(out, { file, target, noHistory });
    if (busy !== undefined) return busy;
  }
  const backupCode = await ensureBackupCopy(out, file, buf, target, backupState);
  if (backupCode !== undefined) return backupCode;
  // History: with a history repository and an in-place write, record the state before the change (when it
  // differs from the head) and then the result, so one revert returns the original.
  const recordHistory = !noHistory && normalized.length > 0 && path.resolve(target) === path.resolve(file) && hasHistorySidecar(file);
  let before: DocxParser | undefined;
  if (recordHistory) {
    try {
      before = new DocxParser(buf, macroOpts);
    } catch {
      before = undefined;
    }
  }
  try {
    await writeDocx(target, buffer);
  } catch (err: unknown) {
    const rawMsg = err instanceof Error ? err.message : String(err);
    return failWithCode(out, rawMsg, 'E_IO', undefined, 'check the output directory is writable, then retry (nothing was written)', err);
  }
  let history: Awaited<ReturnType<typeof recordBatchRevision>> = null;
  if (recordHistory) {
    if (before) {
      await commitStateIfDiffersFromHead(
        file,
        before.readDocumentXml(),
        'normalize: state before repairs',
        extrasRecord(trackedExtras(before.capturePackageSnapshot())),
      );
    }
    history = await recordBatchRevision({
      file,
      target,
      parser: new DocxParser(buffer, macroOpts),
      message: `normalize: ${normalized.length} rewrite${normalized.length === 1 ? '' : 's'}`,
      noHistory,
    });
  }
  if (tolerated) printStderr(`${CLI_NAME}: NOTE: ${toleratedNote(tolerated)}`);
  print(out, `NORMALIZED ${file} -> ${target} (${formatNormalizedInline(normalized)})`);
  for (const n of normalized) print(out, `normalized: ${truncLine(n, 500)}`);
  for (const r of rolledBack) print(out, `repair dropped: ${r.rule}: ${r.reason}`);
  printJson(out, {
    ok: true,
    file,
    output: target,
    normalized,
    remaining,
    rolledBack,
    ...(history ? { history } : {}),
    ...toleratedFields(tolerated),
    ...(backupState.path !== undefined ? { backup: backupState.path } : {}),
  });
  return EXIT_OK;
}
