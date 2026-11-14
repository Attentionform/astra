// Stands in for powershell.exe in tests. The runner starts it as `node fake-powershell.mjs` and passes the same
// environment variables the real scripts read. Behaviour is chosen by FAKE_PS_MODE.
import fs from 'node:fs';
import path from 'node:path';

const mode = process.env.FAKE_PS_MODE || 'ok';
const say = (s) => process.stdout.write(`${s}\n`);

function minimalPdf(pages = 1) {
  const objs = [];
  const kids = [];
  objs.push('<</Type/Catalog/Pages 2 0 R>>');
  for (let i = 0; i < pages; i++) kids.push(`${3 + i * 2} 0 R`);
  objs.push(`<</Type/Pages/Kids[${kids.join(' ')}]/Count ${pages}>>`);
  for (let i = 0; i < pages; i++) {
    objs.push(`<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 100]/Contents ${4 + i * 2} 0 R>>`);
    objs.push('<</Length 0>>\nstream\n\nendstream');
  }
  let out = '%PDF-1.4\n';
  const offsets = [];
  objs.forEach((o, i) => {
    offsets.push(out.length);
    out += `${i + 1} 0 obj\n${o}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n`;
  for (const off of offsets) out += `${String(off).padStart(10, '0')} 00000 n \n`;
  out += `trailer\n<</Size ${objs.length + 1}/Root 1 0 R>>\nstartxref\n${xref}\n%%EOF\n`;
  return out;
}

function bump(counterFile) {
  let n = 0;
  try { n = Number(fs.readFileSync(counterFile, 'utf8')) || 0; } catch { /* first call */ }
  fs.writeFileSync(counterFile, String(n + 1));
  return n + 1;
}

const log = process.env.FAKE_PS_LOG;
if (log) fs.appendFileSync(log, `${JSON.stringify({ mode, word: !!process.env.ASTRA_WORD_IN, pdf: !!process.env.ASTRA_PDF_IN, tagged: process.env.ASTRA_WORD_TAGGED, markup: process.env.ASTRA_WORD_MARKUP, argv: process.argv.slice(2) })}\n`);

if (process.env.ASTRA_WORD_IN) {
  const out = process.env.ASTRA_WORD_OUT;
  const counter = process.env.FAKE_PS_COUNTER;
  const writePdf = () => fs.writeFileSync(out, minimalPdf(Number(process.env.FAKE_PDF_PAGES || 1)));
  switch (mode) {
    case 'ok':
      say('CANDIDATE 4321'); say('VERSION 16.0'); say('PAGES 1');
      if (process.env.ASTRA_WORD_MARKUP === '1') { say('MARKUP 2'); for (const w of (process.env.FAKE_PS_WARN || '').split(',').filter(Boolean)) say(`WARN ${w}`); }
      writePdf();
      if (process.env.ASTRA_WORD_MARKUP === '1') say(process.env.FAKE_PS_RESTORE_FAIL ? `VIEWRESTORE failed ${process.env.FAKE_PS_RESTORE_FAIL}` : 'VIEWRESTORE ok');
      say('OK'); break;
    case 'markup-view-fail':
      say('CANDIDATE 4321'); say('ERR 23 0x00000000 MarkupMode (inline marks with comment balloons)'); process.exit(23); break;
    case 'markup-no-confirm':
      say('CANDIDATE 4321'); say('PAGES 1'); writePdf(); say('OK'); break;
    case 'ok-no-pdf':
      say('CANDIDATE 4321'); say('OK'); break;
    case 'ok-bad-pdf':
      say('CANDIDATE 4321'); fs.writeFileSync(out, 'not a pdf'); say('OK'); break;
    case 'busy-then-ok': {
      const n = bump(counter);
      if (n <= Number(process.env.FAKE_PS_BUSY || 1)) { say('ERR 13 0x00000000 instance predates this call'); process.exit(13); }
      say('CANDIDATE 4321'); writePdf(); say('OK'); break;
    }
    case 'always-busy': bump(counter); say('ERR 13 0x00000000 instance in use'); process.exit(13); break;
    case 'lock-busy': say('ERR 15 0x00000000 another render held the Word launch lock for 30 seconds'); process.exit(15); break;
    case 'ambiguous': bump(counter); say('ERR 14 0x00000000 several new Word processes'); process.exit(14); break;
    case 'password': say('CANDIDATE 4321'); say('ERR 20 0x800A13E9 password'); process.exit(20); break;
    case 'password-refused': say('CANDIDATE 4321'); say('ERR 20 0x800A1520 The password is incorrect. Word cannot open the document.'); process.exit(20); break;
    case 'corrupt': say('CANDIDATE 4321'); say('ERR 20 0x800A16A0 corrupt'); process.exit(20); break;
    case 'export-doc': say('CANDIDATE 4321'); say('ERR 22 0x800A1234 cannot export'); process.exit(22); break;
    case 'export-infra': say('CANDIDATE 4321'); say('ERR 22 0x80004005 unlicensed'); process.exit(22); break;
    case 'no-word': say('ERR 11 0x80040154 class not registered'); process.exit(11); break;
    case 'constrained': say('ERR 12 0x00000000 constrained language mode'); process.exit(12); break;
    case 'hang': say('CANDIDATE 4242'); setInterval(() => undefined, 1000); break;
    case 'garbage': say('hello'); say('PAGES x'); break;
    case 'crash': process.stderr.write('boom from powershell\n'); process.exit(1); break;
    case 'odd-exit': say('CANDIDATE 4321'); say('ERR 99 0x80000001 odd'); process.exit(99); break;
    default: say('ERR 11 0x00000000 unknown fake mode'); process.exit(11);
  }
} else if (process.env.ASTRA_PDF_IN) {
  const dir = process.env.ASTRA_IMG_DIR;
  const fmt = process.env.ASTRA_IMG_FORMAT;
  const total = Number(process.env.FAKE_TOTAL || 3);
  const max = Number(process.env.ASTRA_IMG_MAX);
  const raw = (process.env.ASTRA_IMG_PAGES || '').split(',').filter(Boolean).map(Number);
  switch (mode) {
    case 'ok': case 'missing-files': case 'zero-pages': {
      say(`TOTAL ${mode === 'zero-pages' ? 0 : total}`);
      if (mode === 'zero-pages') { say('OK'); break; }
      let sel;
      if (raw.length === 0) sel = Array.from({ length: Math.min(total, max) }, (_, i) => i + 1);
      else sel = [...new Set(raw.map((n) => Math.max(1, Math.min(total, n))))].sort((a, b) => a - b).slice(0, max);
      for (const n of sel) {
        if (mode !== 'missing-files') fs.writeFileSync(path.join(dir, `page-${n}.${fmt === 'jpeg' ? 'jpg' : 'png'}`), Buffer.from(`fake-${fmt}-${n}-dpi${process.env.ASTRA_IMG_DPI}-q${process.env.ASTRA_IMG_QUALITY}-w${process.env.ASTRA_IMG_MAXWIDTH}`));
        say(`PAGE ${n}`);
      }
      say('OK');
      break;
    }
    case 'err41': say('ERR 41 0x80131500 type not found'); process.exit(41); break;
    case 'err42': say('ERR 42 0x80070056 bad password'); process.exit(42); break;
    case 'err43': say(`TOTAL ${total}`); say('ERR 43 0x88982F50 render failed'); process.exit(43); break;
    case 'constrained': say('ERR 12 0x00000000 constrained language mode'); process.exit(12); break;
    case 'hang': say(`TOTAL ${total}`); setInterval(() => undefined, 1000); break;
    case 'crash': process.stderr.write('boom\n'); process.exit(1); break;
    default: say('ERR 41 0x00000000 unknown fake mode'); process.exit(41);
  }
} else {
  say('ERR 11 0x00000000 no inputs'); process.exit(11);
}
