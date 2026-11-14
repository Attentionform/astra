#!/usr/bin/env node
/**
 * Edge cases of the Word renderer, against the real Word on this machine. No mocks. Windows only.
 * Run after `npm run build --workspace=@attentionform/astra-engine`. Skips (exit 0) when Word is not installed,
 * or when a Word process is already running (the engine must not share it, and neither may this script).
 *
 *   1. A password-protected document      -> E_VALIDATION, a document failure; no Word process left.
 *   2. A damaged document                  -> an error, never a hang; no Word process left.
 *   3. A Word that is not ours: another client's hidden Word is left untouched while ASTra uses its own process;
 *      with --interactive-word, the user's own Word (one document open) keeps its process and its document:
 *      ASTra never quits or changes a Word it did not start.
 *   4. A timeout                           -> E_RENDER that names the leftover process and says ASTra did not stop
 *      it; the script then ends that process itself, so the machine is clean.
 *   5. Two renders at once                 -> both succeed (they run one after the other).
 *   6. A temp folder with spaces and non-ASCII characters in its name.
 *   7. After all of it, a normal render still works (nothing was left in a bad state).
 */
import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const engine = createRequire(import.meta.url)(path.join(root, 'packages', 'engine', 'dist', 'index.js'));

if (process.platform !== 'win32') {
  console.log('word-edge-smoke: SKIP (not Windows)');
  process.exit(0);
}
const word = engine.findWord();
if (!word) {
  console.log('word-edge-smoke: SKIP (Word is not installed)');
  process.exit(process.argv.includes('--expect-word') ? 1 : 0);
}

const ps = (script, timeout = 60000) => spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { encoding: 'utf8', windowsHide: true, timeout });
const wordPids = () =>
  ps('(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id) -join ","')
    .stdout.trim()
    .split(',')
    .filter(Boolean)
    .map(Number);
/** Ends Word processes by id, gently first. Only ever called with ids of processes this script started. */
const endOwn = (pids) => {
  for (const pid of pids) {
    ps(`$p = Get-Process -Id ${Number(pid)} -ErrorAction SilentlyContinue; if ($p) { $null = $p.CloseMainWindow(); Start-Sleep -Seconds 2; $p = Get-Process -Id ${Number(pid)} -ErrorAction SilentlyContinue; if ($p) { Stop-Process -Id ${Number(pid)} -Force } }`);
  }
};
const fail = (m) => {
  console.error(`word-edge-smoke: FAIL: ${m}`);
  process.exit(1);
};
const ok = (m) => console.log(`ok  ${m}`);

if (wordPids().length > 0) {
  console.log('word-edge-smoke: SKIP (a Word process is already running)');
  process.exit(0);
}

const base = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-word-edge-')));
const plain = () => {
  const parser = engine.DocxParser.createNewDocument('blank');
  parser.insertRootNode('<w:p><w:r><w:t>Edge case document</w:t></w:r></w:p>', undefined, 'append');
  return parser.rezip();
};
const exportWord = (buf, name, extra = {}) => engine.convertDocxToPdfDetailed(buf, path.join(base, name), { renderer: 'word', ...extra });
const settle = async (ms = 4000) => {
  const until = Date.now() + ms;
  while (Date.now() < until && wordPids().length > 0) await new Promise((r) => setTimeout(r, 300));
  return wordPids();
};
let holder;
try {
  // ---- 1. password-protected ---------------------------------------------------------------------------------
  {
    const src = path.join(base, 'plain.docx');
    const enc = path.join(base, 'locked.docx');
    fs.writeFileSync(src, plain());
    const made = ps(
      `$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0; try { $d = $w.Documents.Open('${src}'); $d.Password = 'astra-test'; $d.SaveAs2('${enc}'); $d.Close(0) } finally { $w.Quit() }`
    );
    if (made.status !== 0 || !fs.existsSync(enc)) fail(`could not make the password-protected test document: ${made.stderr || made.stdout}`);
    await settle();
    const err = await exportWord(fs.readFileSync(enc), 'locked.pdf').then(() => null, (e) => e);
    if (!err) fail('a password-protected document exported without an error');
    if (err.code !== 'E_VALIDATION' || err.failureClass !== 'document') fail(`password-protected: expected E_VALIDATION/document, got ${err.code}/${err.failureClass}: ${err.message}`);
    if (!/password/i.test(err.message)) fail(`password-protected: the message does not mention a password: ${err.message}`);
    const left = await settle();
    if (left.length) fail(`password-protected: Word process(es) left behind: ${left}`);
    ok(`password-protected document -> ${err.code} (${err.failureClass}): "${err.message.slice(0, 110)}"`);
  }

  // ---- 2. damaged document -----------------------------------------------------------------------------------
  {
    const good = plain();
    const damaged = Buffer.concat([good.subarray(0, Math.floor(good.length / 2)), Buffer.alloc(200, 7)]);
    const t0 = Date.now();
    // The reference check cannot read a cut-off package and refuses it first (E_SECURITY); allowing network references
    // lets the damaged file reach Word, which is the case under test.
    const refused = await exportWord(damaged, 'damaged0.pdf').then(() => null, (e) => e);
    if (!refused || refused.code !== 'E_SECURITY') fail(`damaged document without the allowance: expected E_SECURITY, got ${refused && refused.code}`);
    const err = await exportWord(damaged, 'damaged.pdf', { timeoutMs: 60000, allowNetworkReferences: true }).then(() => null, (e) => e);
    // Word may repair a lightly damaged package and render it; a refusal must be a clear error, never a hang.
    if (Date.now() - t0 > 55000) fail('the damaged document took nearly the whole timeout');
    const left = await settle();
    if (left.length) fail(`damaged: Word process(es) left behind: ${left}`);
    ok(`damaged document -> ${err ? `${err.code} (${err.failureClass})` : 'Word repaired it and rendered'} in ${((Date.now() - t0) / 1000).toFixed(1)} s`);
  }

  // ---- 3a. another automation client's Word ------------------------------------------------------------------
  // A Word that a script started through COM is single-use: a second client gets its own process. ASTra's render
  // must work, use its own process, and leave the other one exactly as it was.
  {
    holder = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-Command', "$w = New-Object -ComObject Word.Application; $w.Visible = $false; [Console]::Out.WriteLine('HELD ' + $w.Version); [Console]::Out.Flush(); $null = [Console]::In.ReadLine(); try { $w.Quit() } catch {}"],
      { windowsHide: true, stdio: ['pipe', 'pipe', 'ignore'] }
    );
    let out = '';
    holder.stdout.on('data', (d) => (out += d));
    const until = Date.now() + 30000;
    while (!/HELD/.test(out) && Date.now() < until) await new Promise((r) => setTimeout(r, 200));
    if (!/HELD/.test(out)) fail('could not start the other client\'s Word instance');
    await new Promise((r) => setTimeout(r, 1000));
    const before = wordPids();
    if (before.length !== 1) fail(`expected one holder Word process, found ${before}`);
    const r = await exportWord(plain(), 'other-client.pdf').then((x) => x, (e) => e);
    const after = wordPids();
    if (r instanceof Error) fail(`another client's Word is hidden and single-use, so the render should have worked: ${r.code}: ${r.message}`);
    if (JSON.stringify(after) !== JSON.stringify(before)) fail(`the other client's Word changed: before ${before}, after ${after}`);
    ok(`another client's hidden Word -> ASTra used its own process and left pid ${before} untouched`);
    holder.stdin.end('\n'); // the other client quits its own Word
    holder = undefined;
    const left = await settle(8000);
    if (left.length) {
      endOwn(left); // the holder did not quit: end exactly the process it started
      fail(`the other client's Word did not quit when asked: ${left}`);
    }
  }

  // ---- 3b. the user's own Word (opt-in: it puts a minimized Word on the taskbar for a few seconds) -------------
  // Word started the usual way, with a document open. A render must not touch it: same process, same open
  // document, still running. (Word gives an automation client its own process, so the render itself works.)
  if (process.argv.includes('--interactive-word')) {
    const userDoc = path.join(base, 'users-own.docx');
    fs.writeFileSync(userDoc, plain());
    spawnSync('powershell.exe', ['-NoProfile', '-Command', `Start-Process -FilePath '${word.exe}' -ArgumentList '/q "${userDoc}"' -WindowStyle Minimized`], { windowsHide: true });
    const until = Date.now() + 40000;
    while (wordPids().length === 0 && Date.now() < until) await new Promise((r) => setTimeout(r, 300));
    const before = wordPids();
    if (before.length !== 1) {
      endOwn(before);
      fail(`expected one interactive Word process, found ${before}`);
    }
    // The window title carries the open document's name (and "Protected View" for a file from the temp folder).
    const state = () => ps(`(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { "$($_.Id): $($_.MainWindowTitle)" }) -join '; '`).stdout.trim();
    let stateBefore = '';
    for (const stop = Date.now() + 40000; Date.now() < stop; ) {
      stateBefore = state();
      if (/users-own/.test(stateBefore)) break;
      await new Promise((r) => setTimeout(r, 1000));
    }
    if (!/users-own/.test(stateBefore)) {
      endOwn(before);
      fail(`the user's Word does not have the document open as expected: "${stateBefore}"`);
    }
    const r = await exportWord(plain(), 'beside-user.pdf').then((x) => x, (e) => e);
    const after = wordPids();
    const stateAfter = state();
    if (r instanceof Error) {
      endOwn(before);
      fail(`a render beside the user's own Word should work, got ${r.code}: ${r.message}`);
    }
    if (r.renderer.used !== 'word') fail('the render did not use Word');
    if (JSON.stringify(after) !== JSON.stringify(before)) {
      endOwn(after);
      fail(`the user's Word process changed: before ${before}, after ${after}`);
    }
    if (stateAfter !== stateBefore) {
      endOwn(after);
      fail(`the user's Word state changed: "${stateBefore}" -> "${stateAfter}"`);
    }
    ok(`the user's own Word (pid ${before}, one document open) -> render worked in its own process; the user's Word is unchanged ("${stateAfter}")`);
    endOwn(after); // the user's Word in this test is ours: close its window (the file in it is unchanged, so nothing is asked)
    const left = await settle(8000);
    if (left.length) fail(`the interactive Word did not close: ${left}`);
  } else {
    console.log("skip the user's own Word (add --interactive-word to run it; it shows a minimized Word on the taskbar for a few seconds)");
  }

  // ---- 4. a timeout ------------------------------------------------------------------------------------------
  {
    const err = await exportWord(plain(), 'timeout.pdf', { timeoutMs: 1000 }).then(() => null, (e) => e);
    if (!err) {
      ok('timeout: Word answered within 1 s on this machine (nothing to check)');
    } else {
      if (err.code !== 'E_RENDER') fail(`timeout: expected E_RENDER, got ${err.code}: ${err.message}`);
      if (!/did not stop it|did not respond/.test(err.message)) fail(`timeout: the message does not say what happened to Word: ${err.message}`);
      ok(`timeout -> ${err.code} (${err.failureClass}): "${err.message.slice(0, 150)}"`);
    }
    // ASTra never force-ends Word. The script does, so the machine is clean for what follows.
    const stray = wordPids();
    if (stray.length) {
      endOwn(stray);
      console.log(`    (a Word process was left running after the timeout, as documented: ${stray}; ended by this script)`);
    }
    const left = await settle(8000);
    if (left.length) fail(`could not clean up after the timeout test: ${left}`);
  }

  // ---- 5. two at once ----------------------------------------------------------------------------------------
  {
    const t0 = Date.now();
    const [a, b] = await Promise.all([exportWord(plain(), 'par-a.pdf'), exportWord(plain(), 'par-b.pdf')]);
    if (a.renderer.used !== 'word' || b.renderer.used !== 'word') fail('two simultaneous renders did not both use Word');
    const left = await settle();
    if (left.length) fail(`parallel: Word process(es) left behind: ${left}`);
    ok(`two renders at once -> both succeeded in ${((Date.now() - t0) / 1000).toFixed(1)} s; no Word left`);
  }

  // ---- 6. an awkward temp folder -----------------------------------------------------------------------------
  {
    const awkward = path.join(base, 'dossier tëst ñ 文書');
    fs.mkdirSync(awkward);
    const saved = process.env.ASTRA_TMPDIR;
    process.env.ASTRA_TMPDIR = awkward;
    try {
      const r = await exportWord(plain(), 'awkward.pdf');
      if (r.renderer.used !== 'word') fail('awkward folder: Word was not used');
      const images = await engine.renderDocxToImages(plain(), { renderer: 'word', pdfToImage: 'windows', dpi: 72, pages: [1] });
      if (images.images.length !== 1) fail('awkward folder: no page image');
      ok('temp folder with spaces and non-ASCII characters -> Word export and page image both work');
    } finally {
      if (saved === undefined) delete process.env.ASTRA_TMPDIR;
      else process.env.ASTRA_TMPDIR = saved;
    }
  }

  // ---- 7. still healthy --------------------------------------------------------------------------------------
  {
    const r = await exportWord(plain(), 'final.pdf');
    if (!fs.readFileSync(r.path).subarray(0, 5).equals(Buffer.from('%PDF-'))) fail('final export is not a PDF');
    const left = await settle();
    if (left.length) fail(`final: Word process(es) left behind: ${left}`);
    ok('a normal render after all of it works, and leaves no Word process');
  }
} catch (error) {
  console.error(`word-edge-smoke: FAIL: ${error && error.stack ? error.stack : String(error)}`);
  process.exitCode = 1;
} finally {
  if (holder) {
    holder.stdin.end('\n');
    holder = undefined;
  }
  fs.rmSync(base, { recursive: true, force: true });
}
if (!process.exitCode) console.log('word-edge-smoke: PASS');
