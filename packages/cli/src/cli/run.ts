import { CLI_NAME, CLI_VERSION, EXIT_OK, type GlobalOpts, print, printJson } from './output';
import { failHuman } from './errors';
import { printCommandHelp, printUsage, usage } from './usage';
import { cmdBatch } from './batch-cmd';
import { cmdInspect } from './inspect-cmd';
import { cmdValidate } from './validate-cmd';
import { cmdSanitize } from './sanitize-cmd';
import { cmdNormalize } from './normalize-cmd';
import { cmdExport } from './export-cmd';
import { cmdSearch } from './search-cmd';
import { cmdCreate } from './create-cmd';
import { cmdImport } from './import-cmd';
import { cmdStats } from './stats-cmd';
import { cmdVisual } from './visual-cmd';
import { cmdHistory, cmdPreviewRevert, cmdRedo, cmdRevert } from './history-cmd';

/**
 * Testable entrypoint: returns the process exit code instead of exiting.
 * Never throws for expected failures (fail-closed codes); unexpected throws
 * are caught and mapped to EXIT_FAIL.
 */
export async function run(argv: string[]): Promise<number> {
  const out: GlobalOpts = { json: false, verbose: false };
  const rest: string[] = [];
  for (const a of argv) {
    if (a === '--json') out.json = true;
    else if (a === '--verbose') out.verbose = true;
    else rest.push(a);
  }
  const [cmd, ...cmdArgs] = rest;
  if (cmd === undefined) return usage(out);
  try {
    switch (cmd) {
      case 'inspect':
        return await cmdInspect(out, cmdArgs);
      case 'validate':
        return await cmdValidate(out, cmdArgs);
      case 'sanitize':
        return await cmdSanitize(out, cmdArgs);
      case 'batch':
        return await cmdBatch(out, cmdArgs);
      case 'normalize':
        return await cmdNormalize(out, cmdArgs);
      case 'export':
        return await cmdExport(out, cmdArgs);
      case 'search':
        return await cmdSearch(out, cmdArgs);
      case 'create':
        return await cmdCreate(out, cmdArgs);
      case 'import':
        return await cmdImport(out, cmdArgs);
      case 'stats':
        return await cmdStats(out, cmdArgs);
      case 'visual':
        return await cmdVisual(out, cmdArgs);
      case 'history':
        return await cmdHistory(out, cmdArgs);
      case 'preview-revert':
        return await cmdPreviewRevert(out, cmdArgs);
      case 'revert':
        return await cmdRevert(out, cmdArgs);
      case 'redo':
        return await cmdRedo(out, cmdArgs);
      case '--version':
      case '-v':
      case 'version':
        print(out, `${CLI_NAME} ${CLI_VERSION}`);
        printJson(out, { ok: true, name: CLI_NAME, version: CLI_VERSION });
        return EXIT_OK;
      case '--help':
      case '-h':
      case 'help':
        if (cmdArgs.length === 0) return printUsage(out);
        if (['inspect', 'validate', 'sanitize', 'batch', 'normalize', 'export', 'search', 'create', 'import', 'stats', 'visual', 'history', 'preview-revert', 'revert', 'redo'].includes(cmdArgs[0])) {
          return printCommandHelp(out, cmdArgs[0]);
        }
        return usage(out, `unknown help topic "${cmdArgs[0]}"`);
      default:
        return usage(out, `unknown command "${cmd}". Expected inspect|validate|sanitize|batch|normalize|export|search|create|import|stats|visual|history|preview-revert|revert|redo.`);
    }
  } catch (err: unknown) {
    return failHuman(out, err, { fallbackCode: 'E_VALIDATION' });
  }
}
