import { describe, it, expect, afterEach } from 'vitest';
import * as path from 'path';
import packageJson from '../package.json';
import {
  CLI_USAGE_ERROR,
  SERVER_VERSION,
  buildUsageText,
  getServedToolNames,
  initAllowedDirs,
  main,
} from '../src/index';

const ALL_TOOLS = [
  'open_document',
  'inspect_document',
  'search_content',
  'execute_batch',
  'save_document',
  'close_document',
  'new_document',
  'compose_document',
  'export_document',
  'visual_inspect',
  'get_document_history',
  'preview_revert',
  'revert_document',
  'redo_document',
  'validate_document',
  'get_stats',
  'sanitize_document',
  'normalize_document',
  'import_document',
];

let stdoutChunks: string[] = [];
let stderrChunks: string[] = [];
let consoleErrorCalls: string[] = [];
const origStdoutWrite = process.stdout.write.bind(process.stdout);
const origStderrWrite = process.stderr.write.bind(process.stderr);
const origConsoleError = console.error;

function captureIO(): void {
  stdoutChunks = [];
  stderrChunks = [];
  consoleErrorCalls = [];
  (process.stdout as unknown as { write: (...a: unknown[]) => boolean }).write = ((chunk: unknown) => {
    stdoutChunks.push(String(chunk));
    return true;
  }) as never;
  (process.stderr as unknown as { write: (...a: unknown[]) => boolean }).write = ((chunk: unknown) => {
    stderrChunks.push(String(chunk));
    return true;
  }) as never;
  console.error = (...args: unknown[]) => {
    consoleErrorCalls.push(args.map(String).join(' '));
  };
}

afterEach(() => {
  (process.stdout as unknown as { write: unknown }).write = origStdoutWrite;
  (process.stderr as unknown as { write: unknown }).write = origStderrWrite;
  console.error = origConsoleError;
  delete process.env.ASTRA_EXTRA_DOC_ROOTS;
  delete process.env.ASTRA_ALLOWED_DIR;
});

function allOutput(): string {
  return [...stdoutChunks, ...stderrChunks, ...consoleErrorCalls].join('\n');
}

describe('mcp-server cli flags', () => {
  it('served tool list holds all nineteen tools', () => {
    expect(getServedToolNames().sort()).toEqual([...ALL_TOOLS].sort());
    expect(getServedToolNames()).toHaveLength(19);
  });

  it('--help prints usage on stdout, exit 0, server never starts', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--help']);
    expect(code).toBe(0);
    const out = stdoutChunks.join('');
    expect(out).toContain('Usage:');
    expect(out).toContain('--allowed-dir');
    expect(out).toContain('stdio');
    expect(out).toContain('Examples:');
    expect(out).toMatch(/19 tools/);
    for (const tool of ALL_TOOLS) expect(out).toContain(tool);
    expect(stderrChunks.join('')).toBe('');
    expect(allOutput()).not.toContain('serving');
  });

  it('-h behaves like --help', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '-h']);
    expect(code).toBe(0);
    expect(stdoutChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('--version prints the package.json version, exit 0, server never starts', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--version']);
    expect(code).toBe(0);
    expect(stdoutChunks.join('').trim()).toBe((packageJson as { version: string }).version);
    expect(stdoutChunks.join('').trim()).toBe(SERVER_VERSION);
    expect(allOutput()).not.toContain('serving');
  });

  it('-V behaves like --version', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '-V']);
    expect(code).toBe(0);
    expect(stdoutChunks.join('').trim()).toBe((packageJson as { version: string }).version);
    expect(allOutput()).not.toContain('serving');
  });

  it('unknown flag prints usage to stderr, exit 2, server never starts', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--nope']);
    expect(code).toBe(CLI_USAGE_ERROR);
    expect(CLI_USAGE_ERROR).toBe(2);
    expect(stdoutChunks.join('')).toBe('');
    expect(stderrChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('single-dash unknown flag exits 2 with usage on stderr', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '-x']);
    expect(code).toBe(2);
    expect(stderrChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('--allowed-dir without a value exits 2 with usage on stderr', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--allowed-dir']);
    expect(code).toBe(2);
    expect(stderrChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('--allowed-dir with empty value exits 2', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--allowed-dir=']);
    expect(code).toBe(2);
    expect(stderrChunks.join('')).toContain('Usage:');
  });

  it('initAllowedDirs accepts space and equals forms', () => {
    initAllowedDirs(['node', 'x', '--allowed-dir', '/tmp/astra-a']);
    expect(process.env.ASTRA_EXTRA_DOC_ROOTS ?? '').toContain(path.resolve('/tmp/astra-a'));
    delete process.env.ASTRA_EXTRA_DOC_ROOTS;
    initAllowedDirs(['node', 'x', '--allowed-dir=/tmp/astra-b']);
    expect(process.env.ASTRA_EXTRA_DOC_ROOTS ?? '').toContain(path.resolve('/tmp/astra-b'));
  });

  it('usage text stays in sync with the served tool list', () => {
    const usage = buildUsageText();
    expect(usage).toContain(`${getServedToolNames().length} tools`);
    for (const tool of getServedToolNames()) expect(usage).toContain(tool);
  });

  it('-- alone exits 2 with usage on stderr', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--']);
    expect(code).toBe(2);
    expect(stdoutChunks.join('')).toBe('');
    expect(stderrChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('flags after -- are positional: -- --help exits 2, never prints help', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--', '--help']);
    expect(code).toBe(2);
    expect(stdoutChunks.join('')).toBe('');
    expect(stderrChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('flags after -- are positional: -- --version exits 2', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--', '--version']);
    expect(code).toBe(2);
    expect(stdoutChunks.join('')).toBe('');
    expect(stderrChunks.join('')).toContain('Usage:');
  });

  it('initAllowedDirs stops at --: flags after it stay positional', () => {
    initAllowedDirs(['node', 'x', '--allowed-dir', '/tmp/astra-a', '--', '--allowed-dir', '/tmp/astra-b']);
    const roots = process.env.ASTRA_EXTRA_DOC_ROOTS ?? '';
    expect(roots).toContain(path.resolve('/tmp/astra-a'));
    expect(roots).not.toContain(path.resolve('/tmp/astra-b'));
  });

  it('--allowed-dir with a dash-leading space value exits 2', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--allowed-dir', '-x']);
    expect(code).toBe(2);
    expect(stdoutChunks.join('')).toBe('');
    expect(stderrChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('--allowed-dir=-foo exits 2 like the space form with a dash value', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--allowed-dir=-foo']);
    expect(code).toBe(2);
    expect(stdoutChunks.join('')).toBe('');
    expect(stderrChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('initAllowedDirs rejects dash-leading values in both forms', () => {
    initAllowedDirs(['node', 'x', '--allowed-dir', '-x']);
    expect(process.env.ASTRA_EXTRA_DOC_ROOTS ?? '').toBe('');
    delete process.env.ASTRA_EXTRA_DOC_ROOTS;
    initAllowedDirs(['node', 'x', '--allowed-dir=-foo']);
    expect(process.env.ASTRA_EXTRA_DOC_ROOTS ?? '').toBe('');
  });

  it('--help wins over --version when both are present', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--help', '--version']);
    expect(code).toBe(0);
    expect(stdoutChunks.join('')).toContain('Usage:');
    expect(allOutput()).not.toContain('serving');
  });

  it('--allow-client-session-ids is a known flag (rides with --help, shown in usage)', async () => {
    captureIO();
    const code = await main(['node', 'astra-mcp-server', '--allow-client-session-ids', '--help']);
    expect(code).toBe(0);
    expect(stdoutChunks.join('')).toContain('--allow-client-session-ids');
    expect(allOutput()).not.toContain('serving');
  });
});
