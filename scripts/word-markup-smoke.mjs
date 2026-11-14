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
 *   R17 a long comment (60 words) with two comments close below it: every word is in the PDF text
 *   R17b a very long comment (300 words): recorded, never a failure. Word's balloon has a maximum height and ends with an
 *       ellipsis (about 150 words on 16.0.14334.20974); WORD_LONG_COMMENT_WARNING says so
 *   R16 the tested Word build
 * Recorded cases (`--cases`; they produce documentation and never fail the run), once for each engine that is available:
 *   R5 comment replies and a resolved comment   R6 comments in a header, a footer, a footnote and a table cell
 *   R7 15 comments on one page                  R8 formatting, move and table-row revisions
 *   R9 right-to-left text with a comment        R10 a 30-page document with comments on three pages
 *   R11 w:revisionView hiding markup            R12 a comment anchored in the middle of a sentence
 *   R13 a plain control document
 *   Each prints `CASE <id> <engine>: <facts>`; the page images are written under the output folder for the owner to look at.
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
const longTokens = (n) => Array.from({ length: n }, (_, i) => `LONGTOKEN${String(i + 1).padStart(3, '0')}`);
const LONG = longTokens(60);

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
    const text = longComment && i === 0 ? longTokens(typeof longComment === 'number' ? longComment : 60).join(' ') : `COMMENT-BY-${upper(a)}`;
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

  // R17: what Word does with a long comment is measured, not assumed (the warning text depends on it).
  for (const [id, n] of [['R17', 60], ['R17b', 300]]) {
    try {
      const f = path.join(dir, `${id}.docx`);
      fs.writeFileSync(f, reviewDocx({ longComment: n }));
      const out = path.join(dir, `${id}.pdf`);
      const o = await engine.wordDocxToPdf(f, out, { markup: true, tagged: false });
      const t = pdfText(out);
      const missing = longTokens(n).filter((tok) => !has(t, tok));
      const detail = `${n}-word comment with two comments below it: ${missing.length === 0 ? 'every word is in the PDF' : `${missing.length} of ${n} words missing, first ${missing[0]}`}`;
      if (id === 'R17') record(id, missing.length === 0 && o.markup?.mode === 2 && o.markup.warnings.includes(engine.WORD_LONG_COMMENT_WARNING), `${detail}; the result carries the long-comment warning`, 'the long balloon');
      else {
        info(`${id}: ${detail} (recorded)`);
        // The ellipsis at the foot of a cut balloon is a button Word draws; whether it reaches the PDF text is only noted here. The page image shows it.
        info(`${id}: an ellipsis character ${/…|\.\.\./.test(t.slice(t.indexOf('LONGTOKEN001'))) ? 'is' : 'is not'} in the PDF text (recorded; look at the page image)`);
      }
      const r = await engine.renderDocxToImages(fs.readFileSync(f), { renderer: 'word', showMarkup: true });
      info(`page images (${id}, Word): ${r.outDir}`);
    } catch (e) {
      const why = e instanceof Error ? e.message : String(e);
      if (id === 'R17') record(id, false, why);
      else info(`${id}: could not be rendered (${why.slice(0, 160)}); it is a recorded case, so this is not a failure`);
    }
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
  // The Word of the render just before can take a few seconds to exit, so wait up to 30 seconds for it before deciding.
  const wordCount = () => Number(ps('(Get-Process WINWORD -ErrorAction SilentlyContinue | Measure-Object).Count')) || 0;
  let running = wordCount();
  for (let waited = 0; running > 0 && waited < 30000; waited += 3000) {
    await new Promise((r) => setTimeout(r, 3000));
    running = wordCount();
  }
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

// ---- recorded cases R5 to R13 -------------------------------------------------------------------------------------
const W15 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const CT_BASE = {
  main: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  commentsEx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
};
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** A package from raw parts. `parts`: [{ name, xml, ct, rel }]; `sect`: extra sectPr children (header/footer references). */
function packageDocx({ body, parts = [], sect = '' }) {
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W} ${W15}><w:body>${body}<w:sectPr>${sect}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const overrides = parts.map((x) => `<Override PartName="/word/${x.name}" ContentType="${x.ct}"/>`).join('');
  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_BASE.main}"/>${overrides}</Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RT}/officeDocument" Target="word/document.xml"/></Relationships>`;
  const drels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${parts.map((x, i) => `<Relationship Id="${x.id ?? `rId${i + 1}`}" Type="${x.rel}" Target="${x.name}"/>`).join('')}</Relationships>`;
  const zip = new AdmZip();
  zip.addFile('[Content_Types].xml', Buffer.from(ct));
  zip.addFile('_rels/.rels', Buffer.from(rels));
  zip.addFile('word/document.xml', Buffer.from(doc));
  zip.addFile('word/_rels/document.xml.rels', Buffer.from(drels));
  for (const x of parts) zip.addFile(`word/${x.name}`, Buffer.from(x.xml));
  return zip.toBuffer();
}
const cmt = (id, author, text, extra = '') => `<w:comment w:id="${id}" w:author="${author}" w:date="2026-01-0${(id % 8) + 1}T10:00:00Z" w:initials="${author[0]}"><w:p ${extra}><w:r><w:t>${esc(text)}</w:t></w:r></w:p></w:comment>`;
const commentsPart = (items) => ({ name: 'comments.xml', ct: CT_BASE.comments, rel: `${RT}/comments`, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W} ${W15}>${items.join('')}</w:comments>` });
const anchored = (id, inner) => `<w:commentRangeStart w:id="${id}"/>${inner}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
const para = (inner) => `<w:p>${inner}</w:p>`;
const run = (t) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;

function caseFixtures() {
  const f = {};
  // R5: a thread with a reply, and a resolved comment
  f.R5 = packageDocx({
    body: para(anchored(0, run('Threaded text. ')) + anchored(1, run('')) ) + para(anchored(2, run('Resolved text.'))),
    parts: [
      commentsPart([cmt(0, 'Alice', 'THREAD-ROOT-TEXT', 'w14:paraId="0000000A"'), cmt(1, 'Bob', 'THREAD-REPLY-TEXT', 'w14:paraId="0000000B"'), cmt(2, 'Carol', 'RESOLVED-COMMENT-TEXT', 'w14:paraId="0000000C"')]),
      { name: 'commentsExtended.xml', ct: CT_BASE.commentsEx, rel: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx ${W} ${W15}><w15:commentEx w15:paraId="0000000A" w15:done="0"/><w15:commentEx w15:paraId="0000000B" w15:paraIdParent="0000000A" w15:done="0"/><w15:commentEx w15:paraId="0000000C" w15:done="1"/></w15:commentsEx>` },
    ],
  });
  // R6a to R6d: a comment in a header, a footer, a footnote and a table cell (one document each, so a load failure shows which one)
  const story = (name, tag, text, id) => {
    const isHeader = name.startsWith('header');
    return { name, id: `rId${tag}`, ct: isHeader ? CT_BASE.header : CT_BASE.footer, rel: `${RT}/${isHeader ? 'header' : 'footer'}`, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${isHeader ? 'hdr' : 'ftr'} ${W}>${para(anchored(id, run(text)))}</w:${isHeader ? 'hdr' : 'ftr'}>` };
  };
  f.R6a = packageDocx({
    body: para(run('Body text under a header.')),
    sect: '<w:headerReference w:type="default" r:id="rIdH"/>',
    parts: [commentsPart([cmt(0, 'Alice', 'HEADER-COMMENT')]), story('header1.xml', 'H', 'Header text.', 0)],
  });
  f.R6b = packageDocx({
    body: para(run('Body text above a footer.')),
    sect: '<w:footerReference w:type="default" r:id="rIdF"/>',
    parts: [commentsPart([cmt(0, 'Bob', 'FOOTER-COMMENT')]), story('footer1.xml', 'F', 'Footer text.', 0)],
  });
  f.R6c = packageDocx({
    body: para(run('Body text with a footnote.') + `<w:r><w:footnoteReference w:id="1"/></w:r>`),
    parts: [
      commentsPart([cmt(0, 'Carol', 'FOOTNOTE-COMMENT')]),
      { name: 'footnotes.xml', ct: CT_BASE.footnotes, rel: `${RT}/footnotes`, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${para('<w:r><w:footnoteRef/></w:r>' + anchored(0, run(' Footnote text.')))}</w:footnote></w:footnotes>` },
    ],
  });
  f.R6d = packageDocx({
    body: `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${para(anchored(0, run('Table cell text.')))}</w:tc></w:tr></w:tbl>` + para(run('After the table.')),
    parts: [commentsPart([cmt(0, 'Dave', 'TABLE-COMMENT')])],
  });
  // R7: 15 comments on one page
  const many = Array.from({ length: 15 }, (_, i) => i);
  f.R7 = packageDocx({
    body: many.map((i) => para(anchored(i, run(`Line ${i + 1} of fifteen commented lines.`)))).join(''),
    parts: [commentsPart(many.map((i) => cmt(i, ['Alice', 'Bob', 'Carol'][i % 3], `MANY-COMMENT-${String(i + 1).padStart(2, '0')}`)))],
  });
  // R8: formatting change, a move, an inserted table row
  const rev = (tag, id, inner) => `<w:${tag} w:id="${id}" w:author="Alice" w:date="2026-01-01T10:00:00Z">${inner}</w:${tag}>`;
  f.R8 = packageDocx({
    body:
      para(`<w:r><w:rPr><w:b/>${rev('rPrChange', 30, '<w:rPr/>')}</w:rPr><w:t>FORMATCHANGED-TEXT</w:t></w:r>`) +
      para(`<w:moveFromRangeStart w:id="31" w:name="mv1" w:author="Alice" w:date="2026-01-01T10:00:00Z"/>${rev('moveFrom', 32, '<w:r><w:t>MOVED-TEXT</w:t></w:r>')}<w:moveFromRangeEnd w:id="31"/>`) +
      para(run('A paragraph between.')) +
      para(`<w:moveToRangeStart w:id="33" w:name="mv1" w:author="Alice" w:date="2026-01-01T10:00:00Z"/>${rev('moveTo', 34, '<w:r><w:t>MOVED-TEXT</w:t></w:r>')}<w:moveToRangeEnd w:id="33"/>`) +
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:trPr>${rev('ins', 35, '')}</w:trPr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${para(run('ROWINSERTED-TEXT'))}</w:tc></w:tr></w:tbl>`,
  });
  // R9: right to left
  f.R9 = packageDocx({
    body: `<w:p><w:pPr><w:bidi/></w:pPr>${`<w:commentRangeStart w:id="0"/><w:r><w:rPr><w:rtl/></w:rPr><w:t>שלום עולם זהו טקסט בעברית</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`}</w:p>` + para(run('Left to right text after it.')),
    parts: [commentsPart([cmt(0, 'Alice', 'RTL-COMMENT-TEXT')])],
  });
  // R10: 30 pages, comments on pages 3, 12 and 25
  const pages = Array.from({ length: 30 }, (_, i) => i + 1);
  const commentOn = { 3: 0, 12: 1, 25: 2 };
  f.R10 = packageDocx({
    body: pages.map((n) => para((n in commentOn ? anchored(commentOn[n], run(`Page ${n} text.`)) : run(`Page ${n} text.`)) + (n < 30 ? '<w:r><w:br w:type="page"/></w:r>' : ''))).join(''),
    parts: [commentsPart([cmt(0, 'Alice', 'PAGE-COMMENT-3'), cmt(1, 'Bob', 'PAGE-COMMENT-12'), cmt(2, 'Carol', 'PAGE-COMMENT-25')])],
  });
  // R11: the document asks for a view that hides markup
  f.R11 = packageDocx({
    body: para(anchored(0, run('Text with a comment. ')) + `<w:del w:id="40" w:author="Alice" w:date="2026-01-01T10:00:00Z"><w:r><w:delText>HIDDEN-VIEW-DELETED</w:delText></w:r></w:del>`),
    parts: [
      commentsPart([cmt(0, 'Alice', 'HIDDEN-VIEW-COMMENT')]),
      { name: 'settings.xml', ct: CT_BASE.settings, rel: `${RT}/settings`, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W}><w:revisionView w:markup="0" w:comments="0" w:insDel="0" w:formatting="0"/></w:settings>` },
    ],
  });
  // R12: a comment in the middle of a sentence
  f.R12 = packageDocx({
    body: para(run('The quick brown fox ') + anchored(0, run('MIDSENTENCE')) + run(' jumps over the lazy dog and keeps running to the far end of the line.')),
    parts: [commentsPart([cmt(0, 'Alice', 'MIDSENTENCE-COMMENT')])],
  });
  // R13: plain control
  f.R13 = packageDocx({ body: para(run('A plain paragraph.')) + para(run('Another plain paragraph.')) });
  return f;
}

const CASE_FACTS = {
  R5: (t) => `root ${has(t, 'THREAD-ROOT-TEXT')}, reply ${has(t, 'THREAD-REPLY-TEXT')}, resolved ${has(t, 'RESOLVED-COMMENT-TEXT')}`,
  R6a: (t) => `header comment ${has(t, 'HEADER-COMMENT')}, header text ${has(t, 'Header text.')}`,
  R6b: (t) => `footer comment ${has(t, 'FOOTER-COMMENT')}, footer text ${has(t, 'Footer text.')}`,
  R6c: (t) => `footnote comment ${has(t, 'FOOTNOTE-COMMENT')}, footnote text ${has(t, 'Footnote text.')}`,
  R6d: (t) => `table comment ${has(t, 'TABLE-COMMENT')}, cell text ${has(t, 'Table cell text.')}`,
  R7: (t) => `${Array.from({ length: 15 }, (_, i) => `MANY-COMMENT-${String(i + 1).padStart(2, '0')}`).filter((c) => has(t, c)).length} of 15 comments in the text`,
  R8: (t) => `formatted-text ${has(t, 'FORMATCHANGED-TEXT')}, moved-text x${(t.match(/MOVED-TEXT/g) ?? []).length}, inserted-row ${has(t, 'ROWINSERTED-TEXT')}, "Formatted" note ${/Formatted/.test(t)}`,
  R9: (t) => `comment ${has(t, 'RTL-COMMENT-TEXT')}, hebrew text ${has(t, 'שלום') || has(t, 'םולש')}`,
  R10: (t) => `comments 3/12/25: ${['PAGE-COMMENT-3', 'PAGE-COMMENT-12', 'PAGE-COMMENT-25'].map((c) => has(t, c)).join('/')}`,
  R11: (t) => `comment ${has(t, 'HIDDEN-VIEW-COMMENT')}, deleted text ${has(t, 'HIDDEN-VIEW-DELETED')}`,
  R12: (t) => `comment ${has(t, 'MIDSENTENCE-COMMENT')}, sentence words ${['quick', 'brown', 'fox', 'MIDSENTENCE', 'jumps', 'lazy', 'dog'].every((w) => has(t, w))}`,
  R13: () => '',
};

/** `render(buffer, markup, id)` returns a PDF path. Recorded, never a failure: a case that throws is recorded as the error. */
async function recordedCases(engineName, render) {
  console.log(`\nRecorded cases (${engineName})`);
  const fixtures = caseFixtures();
  for (const id of Object.keys(fixtures)) {
    const one = async (markup) => {
      try {
        const pdf = await render(fixtures[id], markup, `${id}-${markup ? 'on' : 'off'}`);
        return { text: pdfText(pdf), pages: pdfPages(pdf) };
      } catch (e) {
        return { error: (e instanceof Error ? e.message : String(e)).replace(/\s+/g, ' ').slice(0, 120) };
      }
    };
    const on = await one(true);
    const off = await one(false);
    const facts = (r) => (r.error ? `could not render (${r.error})` : CASE_FACTS[id](r.text));
    const detail = on.error || off.error ? '' : `; pages ${on.pages} with, ${off.pages} without; text ${norm(on.text) === norm(off.text) ? 'identical' : 'differs'}`;
    console.log(`CASE ${id} ${engineName}: with markup: ${facts(on)}; without: ${facts(off)}${detail}`);
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
  for (const b of bins) {
    tested.push(libreOfficeChecks(b));
    if (flag('--cases')) await recordedCases(`libreoffice ${sofficeVersion(b) ?? ''}`.trim(), async (buf, markup, id) => soffice(b, buf, markup ? WITH : WITHOUT, id));
  }
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
if (!flag('--skip-word') && process.platform === 'win32') {
  await wordChecks();
  if (flag('--cases') && engine.findWord()) {
    const wdir = fs.mkdtempSync(path.join(outRoot, 'word-cases-'));
    await recordedCases('word', async (buf, markup, id) => {
      const f = path.join(wdir, `${id}.docx`);
      fs.writeFileSync(f, buf);
      const out = path.join(wdir, `${id}.pdf`);
      await engine.wordDocxToPdf(f, out, { markup, tagged: false });
      return out;
    });
  }
}
else if (process.platform !== 'win32') console.log('\nWord checks: SKIP (not Windows)');

const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length} passed, ${failed.length} failed. Output: ${outRoot}`);
process.exit(failed.length > 0 ? 1 : 0);
