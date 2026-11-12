#!/usr/bin/env node
/**
 * Installs the packed packages the way a user does, into a clean global prefix, and runs them
 * through the command shims that npm writes for this operating system (astra.cmd on Windows):
 * the CLI, the MCP server over stdio, and `npx`. It also covers the paths that have failed on
 * Windows: a folder named with a trailing separator, a folder whose name has a space, an absolute
 * image path with a drive letter, and the 8.3 short form of a folder.
 *
 *   node scripts/platform-smoke.mjs [--platform-out <dir>]
 *
 * <dir>/<platform> (for example platform-out/win32-x64) holds this platform's validator package when
 * it was built in this run; otherwise the validator binary comes from npm at the wrapper's range.
 * Run it from the repository root after `npm run build:packages`.
 */
import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const isWin = process.platform === 'win32';
const host = `${process.platform}-${process.arch}`;
const argv = process.argv.slice(2);
const platformOut = path.resolve(argv.includes('--platform-out') ? argv[argv.indexOf('--platform-out') + 1] : path.join(root, 'platform-out'));

const work = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-smoke-'));
const pkgs = path.join(work, 'pkgs');
const prefix = path.join(work, 'global');
const binDir = isWin ? prefix : path.join(prefix, 'bin');
fs.mkdirSync(pkgs, { recursive: true });

// A one-pixel PNG, for the image steps.
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');

let step = 0;
function log(message) {
  console.log(`[smoke ${++step}] ${message}`);
}
function fail(message) {
  console.error(`[smoke] FAILED: ${message}`);
  process.exit(1);
}

/**
 * Quotes an argument for cmd.exe, which is what runs a .cmd shim. The C runtime reads a backslash
 * that comes before a double quote as an escape, so backslashes before a quote, and any at the very
 * end (they sit before the closing quote), are doubled, so a folder given with a trailing backslash reaches the program unchanged.
 */
const q = (a) => {
  if (!isWin) return String(a);
  const escaped = String(a).replace(/(\\*)"/g, '$1$1\\"').replace(/(\\+)$/, '$1$1');
  return `"${escaped}"`;
};

/** The environment with the global bin folder first on PATH (the variable is `Path` on Windows). */
function envWithBin(extra = {}) {
  const env = { ...process.env, ...extra };
  const key = Object.keys(env).find((k) => k.toLowerCase() === 'path') ?? 'PATH';
  env[key] = binDir + path.delimiter + (env[key] ?? '');
  return env;
}

/** Runs a command and returns { status, stdout, stderr }. On Windows it goes through cmd.exe. */
function run(cmd, args, { cwd, env } = {}) {
  const res = isWin
    ? spawnSync([cmd, ...args.map(q)].join(' '), { shell: true, encoding: 'utf8', cwd, env: env ?? process.env })
    : spawnSync(cmd, args, { encoding: 'utf8', cwd, env: env ?? process.env });
  return { status: res.status ?? -1, stdout: res.stdout ?? '', stderr: res.stderr ?? '' };
}

function mustRun(label, cmd, args, opts) {
  const res = run(cmd, args, opts);
  if (res.status !== 0) fail(`${label} exited ${res.status}\n${res.stdout}\n${res.stderr}`);
  return res;
}

// 1. Pack every package.
log(`platform ${host}; work folder ${work}`);
for (const pkg of ['ooxml-validator', 'engine', 'cli', 'mcp-server']) {
  mustRun(`npm pack ${pkg}`, 'npm', ['pack', path.join(root, 'packages', pkg), '--pack-destination', pkgs]);
}
const extra = [];
if (fs.existsSync(path.join(platformOut, host))) {
  mustRun('npm pack platform package', 'npm', ['pack', path.join(platformOut, host), '--pack-destination', pkgs]);
  log(`using the ${host} validator binary built in this run`);
} else {
  const wrapper = JSON.parse(fs.readFileSync(path.join(root, 'packages', 'ooxml-validator', 'package.json'), 'utf8'));
  extra.push(`@attentionform/ooxml-validator-${host}@${wrapper.validatorBinary.range}`);
  log(`using the ${host} validator binary from npm (${wrapper.validatorBinary.range})`);
}
const tarballs = fs.readdirSync(pkgs).filter((f) => f.endsWith('.tgz')).map((f) => path.join(pkgs, f));

// 2. A global install, which is what writes the astra and astra-mcp-server shims.
mustRun('npm install -g', 'npm', ['install', '-g', '--prefix', prefix, '--no-audit', '--no-fund', ...tarballs, ...extra]);
const env = envWithBin();

// 3. The shims print something, and the version is a version.
for (const bin of ['astra', 'astra-mcp-server']) {
  const res = mustRun(`${bin} --version`, bin, ['--version'], { env });
  if (!/\d+\.\d+\.\d+/.test(res.stdout)) fail(`${bin} --version printed ${JSON.stringify(res.stdout)}`);
  log(`${bin} --version: ${res.stdout.trim()}`);
}
{
  const res = mustRun('astra --help', 'astra', ['--help'], { env });
  if (!res.stdout.includes('Usage: astra')) fail('astra --help printed no usage');
  log('astra --help prints its usage');
}

// 4. Make, inspect and validate a document, in a folder with a space in its name.
const spaced = path.join(work, 'Smoke Docs');
fs.mkdirSync(spaced, { recursive: true });
fs.writeFileSync(path.join(spaced, 'logo.png'), PNG);
const memo = path.join(spaced, 'memo.docx');
mustRun('astra create', 'astra', ['create', memo, '--template', 'memo'], { env });
mustRun('astra inspect', 'astra', ['inspect', memo], { env });
const validated = mustRun('astra validate', 'astra', ['validate', memo], { env });
if (!/VALID/.test(validated.stdout)) fail(`astra validate said ${JSON.stringify(validated.stdout)}`);
log('create, inspect and validate work in a folder whose name has a space');

// 5. A folder given with a trailing separator, and a relative image.
fs.writeFileSync(path.join(spaced, 'rel.md'), '# Relative\n\n![logo](logo.png)\n');
mustRun(
  'astra create --from-md with --allowed-dir ending in a separator',
  'astra',
  ['create', path.join(spaced, 'rel.docx'), '--from-md', path.join(spaced, 'rel.md'), '--allowed-dir', spaced + path.sep],
  { env }
);
log('--allowed-dir with a trailing separator works');

// 6. An absolute image path. On Windows this starts with a drive letter.
const plain = path.join(work, 'plain');
fs.mkdirSync(plain, { recursive: true });
fs.writeFileSync(path.join(plain, 'logo.png'), PNG);
const absolute = path.join(plain, 'logo.png').replace(/\\/g, '/');
fs.writeFileSync(path.join(plain, 'abs.md'), `# Absolute\n\n![logo](${absolute})\n`);
mustRun(
  'astra create --from-md with an absolute image path',
  'astra',
  ['create', path.join(plain, 'abs.docx'), '--from-md', path.join(plain, 'abs.md'), '--allowed-dir', plain],
  { env }
);
log(`an absolute image path works (${absolute})`);

// 7. npx from a project that has the packages installed.
const project = path.join(work, 'project');
fs.mkdirSync(project, { recursive: true });
mustRun('npm init', 'npm', ['init', '-y'], { cwd: project });
mustRun('npm install into a project', 'npm', ['install', '--no-audit', '--no-fund', ...tarballs, ...extra], { cwd: project });
{
  const res = mustRun('npx astra --version', 'npx', ['--no-install', 'astra', '--version'], { cwd: project });
  if (!/\d+\.\d+\.\d+/.test(res.stdout)) fail(`npx astra --version printed ${JSON.stringify(res.stdout)}`);
  log(`npx astra --version: ${res.stdout.trim()}`);
}

// 8. The MCP server over stdio, through its shim: handshake, list the tools, make a document.
function mcpSession(allowedDir, outPath) {
  return new Promise((resolve, reject) => {
    const args = ['--allowed-dir', allowedDir];
    const child = isWin
      ? spawn(['astra-mcp-server', ...args.map(q)].join(' '), { shell: true, env, stdio: ['pipe', 'pipe', 'pipe'] })
      : spawn('astra-mcp-server', args, { env, stdio: ['pipe', 'pipe', 'pipe'] });
    // On Windows the shell owns the Node server and its validator descendants.
    // Killing the shell alone leaves their pipes open and keeps this smoke process alive.
    const stopChild = () => {
      if (isWin && child.pid) spawnSync('taskkill', ['/pid', String(child.pid), '/t', '/f'], { stdio: 'ignore' });
      else child.kill();
    };
    let buffer = '';
    let stderr = '';
    const replies = new Map();
    const timer = setTimeout(() => {
      stopChild();
      reject(new Error(`the MCP server did not answer in time\nstderr: ${stderr}\nstdout: ${buffer}`));
    }, 90_000);
    const send = (message) => child.stdin.write(JSON.stringify(message) + '\n');
    child.stderr.on('data', (d) => (stderr += d));
    child.stdout.on('data', (d) => {
      buffer += d;
      let nl;
      while ((nl = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, nl).trim();
        buffer = buffer.slice(nl + 1);
        if (!line) continue;
        let msg;
        try {
          msg = JSON.parse(line);
        } catch {
          continue;
        }
        if (msg.id !== undefined) replies.set(msg.id, msg);
        if (msg.id === 1) {
          send({ jsonrpc: '2.0', method: 'notifications/initialized' });
          send({ jsonrpc: '2.0', id: 2, method: 'tools/list' });
          send({ jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'new_document', arguments: { outPath, template: 'blank' } } });
        }
        if (replies.has(2) && replies.has(3)) {
          clearTimeout(timer);
          child.stdin.end();
          if (isWin) stopChild();
          else setTimeout(stopChild, 3000).unref();
          resolve({ tools: replies.get(2), created: replies.get(3) });
        }
      }
    });
    child.on('error', reject);
    send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-03-26', capabilities: {}, clientInfo: { name: 'smoke', version: '1' } } });
  });
}

async function checkMcp(label, allowedDir, outPath) {
  const { tools, created } = await mcpSession(allowedDir, outPath);
  const count = tools.result?.tools?.length ?? 0;
  if (count < 10) fail(`${label}: tools/list returned ${count} tools`);
  if (created.result?.isError || created.error) fail(`${label}: new_document failed: ${JSON.stringify(created)}`);
  if (!fs.existsSync(outPath)) fail(`${label}: new_document reported success but ${outPath} does not exist`);
  log(`${label}: ${count} tools, new_document wrote ${path.basename(outPath)}`);
}

await checkMcp('MCP server, folder with a space', spaced, path.join(spaced, 'mcp.docx'));
await checkMcp('MCP server, folder ending in a separator', plain + path.sep, path.join(plain, 'mcp.docx'));

// 9. Windows only: the 8.3 short form of a folder, the way os.tmpdir() often spells it.
if (isWin) {
  const long = path.join(work, 'A Long Folder Name');
  fs.mkdirSync(long, { recursive: true });
  const res = run('cmd', ['/c', 'for', '%I', 'in', `("${long}")`, 'do', '@echo', '%~sI']);
  const short = res.stdout.trim().split(/\r?\n/).pop() ?? '';
  if (short && short.toLowerCase() !== long.toLowerCase()) {
    await checkMcp('MCP server, allowed folder in short form and file in long form', short, path.join(long, 'short-allowed.docx'));
    await checkMcp('MCP server, allowed folder in long form and file in short form', long, path.join(short, 'short-file.docx'));
  } else {
    log('this machine has no 8.3 short names; the short-name step is skipped');
  }
}

log('all steps passed');
// The validator program can still be shutting down when the MCP servers stop, and Windows will not
// delete a running .exe. The checks are done, so a folder that will not go is only reported.
try {
  fs.rmSync(work, { recursive: true, force: true, maxRetries: 5, retryDelay: 500 });
} catch (err) {
  console.warn(`[smoke] note: could not remove ${work} (${err instanceof Error ? err.message : err}); the runner cleans it up`);
}
