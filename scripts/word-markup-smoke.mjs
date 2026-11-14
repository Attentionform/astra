#!/usr/bin/env node
/**
 * Release checks for "show markup" (tracked changes and comments in renders). Real programs, no mocks.
 * Run after `npm run build --workspace=@attentionform/astra-engine`.
 *
 *   node scripts/word-markup-smoke.mjs [--soffice <path>]... [--out <dir>] [--skip-word] [--skip-libreoffice]
 *                                       [--expect-word] [--check-minimum]
 *
 * LibreOffice acceptance checks, any platform, once per `--soffice` binary (default: the one ASTra finds):
 *   L1  one comment and one tracked deletion: with the margin-comments option the PDF text has the comment text, its
 *       author and the deleted text; without the option the comment text is absent and the deleted text is present
 *   L2a three authors: every comment text and author is in the PDF
 *   L2b one long (about 700-character) comment with two comments close below it: recorded, never a failure, and not
 *       part of the minimum. Observed on every tested version: the long comment is cut short (LIBREOFFICE_LONG_COMMENT_WARNING)
 *   L3  a comment spanning a tracked change, and a comment on the last line of a page: comment text present (the page
 *       counts with and without the option are printed, not asserted)
 *   L4  a plain document with the option on: renders, and its page text equals the render without the option
 *   Each binary prints `LOTABLE <version> L1=.. L2a=.. L2b=.. L3=.. L4=..`. `--check-minimum` asserts that
 *   MIN_LIBREOFFICE_FOR_MARKUP equals the lowest version that passed L1, L2a, L3 and L4 (give it every tested binary).
 *
 * Word checks, Windows only (they also need Word; `--expect-word` makes a missing Word a failure):
 *   R1  Word starts in "No Markup" (set on the SAME automation instance and document, right before the pin):
 *       the start state was really in effect, every comment and deleted text is in the PDF, the view is restored
 *   R2  comments, insertions and deletions by three reviewers: every comment and every deleted text in the PDF
 *   R3  Word starts with one reviewer hidden: all reviewers' comments present, restored
 *   R4  Word starts with comments off: comments present, restored
 *   R14 the default is unchanged: no option and `showMarkup:false` give the same page count and page text, no comments
 *   R15 isolation: with the user's own interactive Word open (only started here, only when no Word runs), it keeps its
 *       process and window through a markup render
 *   R16 the tested Word build
 * Page images for the owner to look at are written under the output folder. Exit code 1 if a machine check fails.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const engine = require(path.join(root, 'packages', 'engine', 'dist', 'index.js'));
const { AdmZip } = engine;

const argv = process.argv.slice(2);
const flag = (n) => argv.includes(n);
const values = (n) => argv.flatMap((a, i) => (a === n && argv[i + 1] ? [argv[i + 1]] : []));
const outRoot = values('--out')[0] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'astra-markup-smoke-'));
fs.mkdirSync(outRoot, { recursive: true });

const results = [];
function record(id, ok, detail, lookAt = '') {
  results.push({ id, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'} ${id}: ${detail}${lookAt ? `  [look at: ${lookAt}]` : ''}`);
}
const info = (m) => console.log(`     ${m}`);

// ---- fixtures ---------------------------------------------------------------------------------------------------
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const upper = (a) => a.toUpperCase();
const LONG = Array.from({ length: 60 }, (_, i) => `LONGTOKEN${String(i + 1).padStart(3, '0')}`);

/** A DOCX with tracked changes and comments built from raw parts (a real review document, not the engine's own markup). */
function reviewDocx({ authors = ['Alice', 'Bob', 'Carol'], longComment = false, spanning = false, pageFill = 0, plain = false } = {}) {
  const body = [];
  const comments = [];
  let cid = 0;
  for (let i = 0; i < pageFill; i++) body.push(`<w:p><w:r><w:t>Filler line ${i + 1} of the page.</w:t></w:r></w:p>`);
  authors.forEach((a, i) => {
    const date = `2026-01-0${i + 1}T10:00:00Z`;
    if (plain) {
      body.push(`<w:p><w:r><w:t>Plain paragraph ${i + 1}. Nothing is tracked here.</w:t></w:r></w:p>`);
      return;
    }
    const del = `<w:del w:id="${10 + i}" w:author="${a}" w:date="${date}"><w:r><w:delText>DELETED-BY-${upper(a)}</w:delText></w:r></w:del>`;
    const ins = `<w:ins w:id="${20 + i}" w:author="${a}" w:date="${date}"><w:r><w:t>INSERTED-BY-${upper(a)}</w:t></w:r></w:ins>`;
    const id = cid++;
    const text = longComment && i === 0 ? LONG.join(' ') : `COMMENT-BY-${upper(a)}`;
    comments.push(`<w:comment w:id="${id}" w:author="${a}" w:date="${date}" w:initials="${a[0]}"><w:p><w:r><w:t>${esc(text)}</w:t></w:r></w:p></w:comment>`);
    const start = `<w:commentRangeStart w:id="${id}"/>`;
    const end = `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
    body.push(
      spanning
        ? `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i + 1} keeps this text. </w:t></w:r>${start}${del}${ins}${end}</w:p>`
        : `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i + 1} keeps this text. </w:t></w:r>${del}${ins}${start}<w:r><w:t xml:space="preserve"> commented words ${i + 1}</w:t></w:r>${end}</w:p>`
    );
  });
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${plain ? '' : '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'}</Types>`;
  const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const drels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${plain ? '' : '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'}</Relationships>`;
  const zip = new AdmZip();
  zip.addFile('[Content_Types].xml', Buffer.from(ct));
  zip.addFile('_rels/.rels', Buffer.from(rels));
  zip.addFile('word/document.xml', Buffer.from(doc));
  zip.addFile('word/_rels/document.xml.rels', Buffer.from(drels));
  if (!plain) zip.addFile('word/comments.xml', Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W}>${comments.join('')}</w:comments>`));
  return zip.toBuffer();
}

// ---- PDF text ---------------------------------------------------------------------------------------------------
function findTool(name) {
  const exts = process.platform === 'win32' ? ['.exe', '.com', ''] : [''];
  for (const dir of (process.env.PATH ?? '').split(path.delimiter)) {
    for (const e of exts) {
      const f = path.join(dir.replace(/^"|"$/g, ''), name + e);
      if (dir && fs.existsSync(f)) return f;
    }
  }
  return null;
}
const pdftotext = findTool('pdftotext');
if (!pdftotext) {
  console.error('word-markup-smoke: FAIL: pdftotext (Poppler) is needed to read the PDFs and is not on PATH');
  process.exit(1);
}
const pdfText = (pdf) => spawnSync(pdftotext, [pdf, '-'], { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }).stdout ?? '';
const pdfPages = (pdf) => {
  const t = spawnSync(pdftotext, [pdf, '-'], { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }).stdout ?? '';
  return t.split('\f').filter((p, i, a) => i < a.length - 1 || p.trim()).length;
};
const norm = (t) => t.replace(/\s+/g, ' ').trim();
const has = (text, needle) => text.includes(needle);

// ---- LibreOffice ------------------------------------------------------------------------------------------------
function soffice(bin, docx, filter, label) {
  const dir = fs.mkdtempSync(path.join(outRoot, `lo-${label}-`));
  const input = path.join(dir, 'in.docx');
  fs.writeFileSync(input, docx);
  const r = spawnSync(bin, [`-env:UserInstallation=${pathToFileURL(path.join(dir, 'profile')).href}`, '--headless', '--convert-to', filter, '--outdir', dir, input], { encoding: 'utf8', timeout: 180000 });
  const pdf = path.join(dir, 'in.pdf');
  if (r.status !== 0 || !fs.existsSync(pdf)) throw new Error(`soffice failed (${r.status}): ${(r.stderr || r.stdout || '').slice(0, 300)}`);
  return pdf;
}

function sofficeVersion(bin) {
  const file = process.platform === 'win32' && /soffice\.exe$/i.test(bin) ? bin.replace(/soffice\.exe$/i, 'soffice.com') : bin;
  const r = spawnSync(file, ['--version'], { encoding: 'utf8', timeout: 30000 });
  const m = /^LibreOffice\s+(\d+\.\d+\.\d+(?:\.\d+)?)/m.exec(r.stdout ?? '');
  return m ? m[1] : null;
}

const WITH = engine.pdfFilterSpec(true, { markup: true });
const WITHOUT = engine.pdfFilterSpec(true);

function libreOfficeChecks(bin) {
  const version = sofficeVersion(bin) ?? 'unknown';
  console.log(`\nLibreOffice ${version} (${bin})`);
  const verdict = {};
  const run = (id, fn, lookAt) => {
    try {
      const msg = fn();
      verdict[id] = true;
      record(`${id} [${version}]`, true, msg, lookAt);
    } catch (e) {
      verdict[id] = false;
      record(`${id} [${version}]`, false, e instanceof Error ? e.message : String(e));
    }
  };
  const need = (cond, msg) => {
    if (!cond) throw new Error(msg);
  };

  run('L1', () => {
    const d = reviewDocx({ authors: ['Alice'] });
    const a = pdfText(soffice(bin, d, WITH, 'l1w'));
    const b = pdfText(soffice(bin, d, WITHOUT, 'l1o'));
    need(has(a, 'COMMENT-BY-ALICE'), 'with the option: comment text missing');
    need(has(a, 'Alice'), 'with the option: comment author missing');
    need(has(a, 'DELETED-BY-ALICE'), 'with the option: deleted text missing');
    need(!has(b, 'COMMENT-BY-ALICE'), 'without the option the comment text is present (the option made no difference)');
    need(has(b, 'DELETED-BY-ALICE'), 'without the option the deleted text is missing');
    return 'comment text, author and deleted text present with the option; comment absent without it';
  }, 'the margin box');

  run('L2a', () => {
    const t = pdfText(soffice(bin, reviewDocx(), WITH, 'l2a'));
    for (const a of ['Alice', 'Bob', 'Carol']) need(has(t, a), `author ${a} missing`);
    for (const a of ['ALICE', 'BOB', 'CAROL']) need(has(t, `COMMENT-BY-${a}`), `comment by ${a} missing`);
    return 'three authors, every comment text and author present';
  }, 'colours per author');

  // L2b: recorded, never a failure. The comment box can run past the page edge and cut each line off.
  {
    const t = pdfText(soffice(bin, reviewDocx({ longComment: true }), WITH, 'l2b'));
    const missing = LONG.filter((tok) => !has(t, tok));
    verdict.L2b = missing.length === 0;
    info(`L2b [${version}]: long comment ${missing.length === 0 ? 'complete' : `CUT OFF (${missing.length} of ${LONG.length} tokens missing, first ${missing[0]})`} (recorded)`);
  }

  run('L3', () => {
    const spanning = reviewDocx({ authors: ['Alice', 'Bob'], spanning: true });
    const t = pdfText(soffice(bin, spanning, WITH, 'l3a'));
    need(has(t, 'COMMENT-BY-ALICE') && has(t, 'COMMENT-BY-BOB'), 'comment spanning a tracked change is missing');
    const fill = reviewDocx({ authors: ['Alice'], pageFill: 44 });
    const w = soffice(bin, fill, WITH, 'l3b');
    const o = soffice(bin, fill, WITHOUT, 'l3c');
    need(has(pdfText(w), 'COMMENT-BY-ALICE'), 'comment near the end of a page is missing');
    info(`page counts near a page end: with the option ${pdfPages(w)}, without ${pdfPages(o)} (recorded, not asserted)`);
    return 'comment spanning a change, and a comment near a page end';
  }, 'where the note icon falls');

  run('L4', () => {
    const d = reviewDocx({ plain: true });
    const a = pdfText(soffice(bin, d, WITH, 'l4w'));
    const b = pdfText(soffice(bin, d, WITHOUT, 'l4o'));
    need(norm(a) === norm(b), 'a plain document renders different text with the option');
    return 'a plain document renders the same text with and without the option';
  });

  const all = ['L1', 'L2a', 'L3', 'L4'].every((k) => verdict[k]);
  console.log(`LOTABLE ${version} ${['L1', 'L2a', 'L2b', 'L3', 'L4'].map((k) => `${k}=${verdict[k] ? 'pass' : 'fail'}`).join(' ')}`);
  return { version, all };
}

// ---- Word -------------------------------------------------------------------------------------------------------
async function wordChecks() {
  const word = engine.findWord();
  if (!word) {
    if (flag('--expect-word')) {
      record('R0', false, 'Word was expected and is not installed or not registered');
      return;
    }
    console.log('\nWord: not found, Word checks skipped');
    return;
  }
  const ps = (cmd) => spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', cmd], { encoding: 'utf8', timeout: 60000 }).stdout?.trim() ?? '';
  const build = ps(`(Get-Item '${word.exe.replace(/'/g, "''")}').VersionInfo.FileVersion`);
  console.log(`\nWord ${word.version} (${word.exe})`);
  record('R16', Boolean(build), `Word build ${build || 'unreadable'}`);
  console.log(`WORDBUILD ${build}`);

  const dir = fs.mkdtempSync(path.join(outRoot, 'word-'));
  const docx = path.join(dir, 'review.docx');
  fs.writeFileSync(docx, reviewDocx());
  const comments = ['COMMENT-BY-ALICE', 'COMMENT-BY-BOB', 'COMMENT-BY-CAROL'];
  const deleted = ['DELETED-BY-ALICE', 'DELETED-BY-BOB', 'DELETED-BY-CAROL'];

  async function render(id, startState) {
    const out = path.join(dir, `${id}.pdf`);
    const o = await engine.wordDocxToPdf(docx, out, { markup: true, tagged: false, ...(startState ? { startState } : {}) });
    return { out, outcome: o, text: pdfText(out) };
  }
  const allPresent = (t) => comments.every((c) => has(t, c)) && deleted.every((d) => has(t, d));

  async function startStateRow(id, state, expectStart, why) {
    try {
      const { outcome, text } = await render(id, state);
      const m = outcome.markup;
      const started = m?.startState ?? '';
      const okStart = expectStart.every((s) => started.includes(s));
      const okText = allPresent(text);
      const okRestore = m?.restore?.ok === true && m?.endState === m?.startState;
      record(id, okStart && okText && okRestore, `${why}: start "${started}", end "${m?.endState ?? ''}", comments and deleted text ${okText ? 'all present' : 'MISSING'}, restore ${m?.restore?.ok ? 'ok' : 'FAILED'}`, 'the page');
    } catch (e) {
      record(id, false, e instanceof Error ? e.message : String(e));
    }
  }
  await startStateRow('R1', { noMarkup: true }, ['revs=0', 'filter=0'], 'start in No Markup');
  await startStateRow('R3', { hideFirstReviewer: true }, ['hidden=1'], 'start with one reviewer hidden');
  await startStateRow('R4', { showComments: false }, ['comments=0'], 'start with comments off');
  await startStateRow('R1b', { noMarkup: true, markupMode: 0, showComments: false, hideFirstReviewer: true }, ['revs=0', 'mode=0', 'comments=0', 'hidden=1'], 'every setting bad at once');

  try {
    const { outcome, text } = await render('R2', undefined);
    record('R2', allPresent(text) && outcome.markup?.mode === 2, `three reviewers: comments and deleted text ${allPresent(text) ? 'all present' : 'MISSING'}, mode ${outcome.markup?.mode}`, 'colours per reviewer');
    // Page images for the owner, through the real engine path.
    const r = await engine.renderDocxToImages(fs.readFileSync(docx), { renderer: 'word', showMarkup: true });
    info(`page images (markup, Word): ${r.outDir}`);
  } catch (e) {
    record('R2', false, e instanceof Error ? e.message : String(e));
  }

  try {
    const a = path.join(dir, 'default-a.pdf');
    const b = path.join(dir, 'default-b.pdf');
    const ra = await engine.convertDocxToPdfDetailed(fs.readFileSync(docx), a, { renderer: 'word' });
    const rb = await engine.convertDocxToPdfDetailed(fs.readFileSync(docx), b, { renderer: 'word', showMarkup: false });
    const ta = pdfText(a);
    const tb = pdfText(b);
    const same = norm(ta) === norm(tb) && pdfPages(a) === pdfPages(b);
    const clean = !comments.some((c) => has(ta, c)) && !deleted.some((d) => has(ta, d));
    record('R14', same && clean && ra.markup === undefined && rb.markup === undefined, `default render: same text and page count with no option and showMarkup:false, ${clean ? 'no comments or deleted text' : 'COMMENTS PRESENT'}, no markup block`);
  } catch (e) {
    record('R14', false, e instanceof Error ? e.message : String(e));
  }

  // R15: only when no Word is running at all, so nobody's own Word is ever touched. The Word started here is ours.
  const running = Number(ps('(Get-Process WINWORD -ErrorAction SilentlyContinue | Measure-Object).Count')) || 0;
  if (running > 0) {
    console.log(`SKIP R15: ${running} Word process(es) already running; not touching them`);
  } else {
    const startCmd = `$p = Start-Process -FilePath '${word.exe.replace(/'/g, "''")}' -ArgumentList '/q','/n','${docx.replace(/'/g, "''")}' -PassThru; $p.Id`;
    const pid = Number(ps(startCmd));
    try {
      await new Promise((r) => setTimeout(r, 8000));
      const before = ps(`$p = Get-Process -Id ${pid} -ErrorAction SilentlyContinue; if ($p) { "$($p.Id)|$($p.MainWindowHandle -ne 0)|$($p.MainWindowTitle)" }`);
      const { text } = await render('R15', undefined);
      const after = ps(`$p = Get-Process -Id ${pid} -ErrorAction SilentlyContinue; if ($p) { "$($p.Id)|$($p.MainWindowHandle -ne 0)|$($p.MainWindowTitle)" }`);
      record('R15', Boolean(before) && before === after && allPresent(text), `interactive Word pid ${pid}: before "${before}", after "${after}"; the render succeeded`, 'nothing else to look at');
    } catch (e) {
      record('R15', false, e instanceof Error ? e.message : String(e));
    } finally {
      // Ours only, by exact PID, gracefully; a leftover is reported, never forced.
      ps(`$p = Get-Process -Id ${pid} -ErrorAction SilentlyContinue; if ($p) { [void]$p.CloseMainWindow(); [void]$p.WaitForExit(15000) }`);
      const left = ps(`if (Get-Process -Id ${pid} -ErrorAction SilentlyContinue) { 'running' }`);
      if (left) console.log(`NOTE: the interactive Word this check started (pid ${pid}) is still running; close it from its window.`);
    }
  }
}

// ---- main -------------------------------------------------------------------------------------------------------
console.log(`word-markup-smoke: output folder ${outRoot}`);
const tested = [];
if (!flag('--skip-libreoffice')) {
  const bins = values('--soffice');
  if (bins.length === 0) {
    const found = engine.getVisualBackendStatus().sofficePath;
    if (found) bins.push(found);
  }
  if (bins.length === 0) record('L0', false, 'no LibreOffice found (pass --soffice <path>)');
  for (const b of bins) tested.push(libreOfficeChecks(b));
  if (flag('--check-minimum')) {
    const passing = tested.filter((t) => t.all).map((t) => t.version);
    const cmp = (a, b) => {
      const x = a.split('.').map(Number);
      const y = b.split('.').map(Number);
      for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] - y[i];
      return 0;
    };
    passing.sort(cmp);
    const lowest = passing[0]?.split('.').slice(0, 3).join('.');
    record('R16-min', lowest === engine.MIN_LIBREOFFICE_FOR_MARKUP, `MIN_LIBREOFFICE_FOR_MARKUP is ${engine.MIN_LIBREOFFICE_FOR_MARKUP}; the lowest tested version that passed L1, L2a, L3 and L4 is ${lowest ?? 'none'}`);
  }
}
if (!flag('--skip-word') && process.platform === 'win32') await wordChecks();
else if (process.platform !== 'win32') console.log('\nWord checks: SKIP (not Windows)');

const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length} passed, ${failed.length} failed. Output: ${outRoot}`);
process.exit(failed.length > 0 ? 1 : 0);
