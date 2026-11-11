// R1: cost of reading a document. Bytes/tokens of word/document.xml vs the ASTra outline.
import fs from 'node:fs';
import path from 'node:path';
import { encode } from 'gpt-tokenizer';
import { connect } from '../lib/mcp.mjs';
import { manifest, freshDir, jsonl, part, CORPUS } from '../lib/exp.mjs';

const out = jsonl('r1_representation');
const dir = freshDir('r1');
const s = await connect(dir);
for (const doc of manifest()) {
  const src = path.join(CORPUS, doc.file);
  fs.copyFileSync(src, path.join(dir, doc.file));
  const xml = part(src, 'word/document.xml');
  const rec = { file: doc.file, writer: doc.writer, kind: doc.kind, pages: doc.measured_pages,
    xml_bytes: Buffer.byteLength(xml), xml_tokens: encode(xml).length };
  const text = [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((m) => [...m[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join('')).filter(Boolean).join('\n');
  rec.text_floor_bytes = Buffer.byteLength(text); rec.text_floor_tokens = encode(text).length;
  const o = await s.call('open_document', { path: path.join(dir, doc.file) });
  if (!o.json?.ok) { out({ ...rec, error: o.text.slice(0, 300) }); continue; }
  const configs = { compact: { fidelity: 'compact' }, standard: { fidelity: 'standard' },
    'standard-fulltext': { fidelity: 'standard', maxTextLength: 10000 } };
  for (const [fidelity, params] of Object.entries(configs)) {
    let text = '', rows = 0, calls = 0, offset = 0, ms = 0, hasMore = true, truncated = false;
    while (hasMore && calls < 400) {
      const r = await s.call('inspect_document', { sessionId: o.json.sessionId, ...params, offset, limit: 200 });
      calls++; ms += r.ms;
      if (!r.json?.ok) { rec[fidelity] = { error: r.text.slice(0, 300) }; break; }
      text += r.json.toonText + '\n'; rows += r.json.returnedNodes;
      truncated ||= !!r.json.truncated;
      hasMore = r.json.hasMore; offset = r.json.nextOffset ?? offset + 200;
    }
    if (!rec[fidelity]) rec[fidelity] = { bytes: Buffer.byteLength(text), tokens: encode(text).length, rows, calls, ms, truncated };
  }
  await s.call('close_document', { sessionId: o.json.sessionId });
  out(rec);
  const std = rec.standard;
  console.log(rec.file.padEnd(28), 'xml', rec.xml_tokens, 'floor', rec.text_floor_tokens, 'std', std.tokens, 'stdFT', rec['standard-fulltext'].tokens, 'ratioFT', (rec.xml_tokens / rec['standard-fulltext'].tokens).toFixed(2));
}
await s.close();
