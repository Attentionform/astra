// Build results/<release>/REPORT.md and summary.json from the raw jsonl files. Nothing is dropped.
import fs from 'node:fs';
import path from 'node:path';
import { RESULTS, RELEASE, readJsonl } from '../lib/exp.mjs';
import { summarize, bootstrapMedianCI, wilson, fmtRate } from '../lib/stats.mjs';

const has = (n) => fs.existsSync(path.join(RESULTS, n + '.jsonl'));
const rd = (n) => { if (n === 'e5_history') return fs.readdirSync(RESULTS).filter((f) => /^e5_history(_w\d+)?\.jsonl$/.test(f)).flatMap((f) => readJsonl(f.replace('.jsonl', ''))); return has(n) ? readJsonl(n) : []; };
const f1 = (x) => (Number.isFinite(x) ? x.toFixed(1) : 'n/a');
const f2 = (x) => (Number.isFinite(x) ? x.toFixed(2) : 'n/a');
const ms = (x) => (Number.isFinite(x) ? (x >= 1000 ? (x / 1000).toFixed(2) + ' s' : x.toFixed(0) + ' ms') : 'n/a');
const table = (head, rows) => ['| ' + head.join(' | ') + ' |', '| ' + head.map(() => '---').join(' | ') + ' |', ...rows.map((r) => '| ' + r.join(' | ') + ' |')].join('\n');
const group = (arr, key) => { const m = new Map(); for (const x of arr) { const k = key(x); (m.get(k) || m.set(k, []).get(k)).push(x); } return m; };
const med = (a) => summarize(a).median;
const sec = [];
const summary = { release: RELEASE };
const hyp = [];

// ---------------- R1 ----------------
{
  const r = rd('r1_representation').filter((x) => x.standard && !x.error);
  if (r.length) {
    const ratio = (x, k) => x.xml_tokens / x[k].tokens;
    const all = r.map((x) => ratio(x, 'standard'));
    const ci = bootstrapMedianCI(all);
    const s = summarize(all);
    summary.r1 = { n: r.length, median_ratio_standard: s.median, ci, p25: s.p25, p75: s.p75, min: s.min, max: s.max };
    const rows = [];
    const bucket = (x) => (x.pages <= 2 ? '1-2' : x.pages <= 12 ? '3-12' : x.pages <= 60 ? '13-60' : '>60');
    for (const [k, g] of [...group(r, (x) => x.kind).entries()].sort()) rows.push([k, g.length, f2(med(g.map((x) => ratio(x, 'standard')))), f2(Math.min(...g.map((x) => ratio(x, 'standard')))), f2(Math.max(...g.map((x) => ratio(x, 'standard')))), f2(med(g.map((x) => x['standard'].tokens / x.text_floor_tokens)))]);
    sec.push(`## R1. Reading cost: document.xml versus the ASTra outline\n\nTokenizer: o200k_base (proxy, see PROTOCOL section 8). Ratio = tokens of \`word/document.xml\` divided by tokens of the outline (standard fidelity, every page of results). "vs text floor" is outline tokens divided by the tokens of the bare document text, i.e. the overhead the outline adds on top of the words themselves.\n\n**All ${r.length} documents: median ratio ${f2(s.median)}x (bootstrap 95% CI ${f2(ci[0])} to ${f2(ci[1])}), interquartile range ${f2(s.p25)} to ${f2(s.p75)}, range ${f2(s.min)} to ${f2(s.max)}.**\n\n` + table(['document kind', 'n', 'median ratio', 'min', 'max', 'outline / text floor (median)'], rows) +
      `\n\nBy writer: ` + [...group(r, (x) => x.writer).entries()].map(([k, g]) => `${k} median ${f2(med(g.map((x) => ratio(x, 'standard'))))}x (n=${g.length})`).join('; ') +
      `\n\nBy size: ` + [...group(r, bucket).entries()].map(([k, g]) => `${k} pages median ${f2(med(g.map((x) => ratio(x, 'standard'))))}x (n=${g.length})`).join('; ') +
      `\n\nCompact fidelity median ratio: ${f2(med(r.map((x) => ratio(x, 'compact'))))}x. Standard with full paragraph text (maxTextLength 10000): ${f2(med(r.map((x) => ratio(x, 'standard-fulltext'))))}x.`);
    hyp.push(['H1', 'median standard-outline ratio >= 3x', `${f2(s.median)}x`, s.median >= 3 ? 'met' : 'NOT met']);
  }
}

// ---------------- E1 ----------------
{
  const e = rd('e1_edit_locality');
  const edits = e.filter((x) => x.kind === 'edit' && !x.skipped);
  if (edits.length) {
    const ok = edits.filter((x) => x.save?.ok);
    const loc = ok.filter((x) => x.diff?.locality);
    const uniq = loc.filter((x) => x.diff.locality.target_unique);
    const nt0 = uniq.filter((x) => x.diff.locality.nontarget_changed === 0);
    const txt = uniq.filter((x) => x.diff.locality.nontarget_text_preserved);
    const intended = uniq.filter((x) => (x.diff.locality.intended_text_ok ?? true) && (x.diff.locality.intended_raw_ok ?? true) && x.diff.locality.target_changed !== false);
    const apply = edits.filter((x) => x.batch?.ok);
    summary.e1 = { trials: edits.length, batch_ok: apply.length, save_ok: ok.length, analysed: uniq.length, nontarget_zero: nt0.length, text_preserved: txt.length, intended_ok: intended.length };
    const byEdit = [...group(edits, (x) => x.edit).entries()].map(([k, g]) => {
      const gu = g.filter((x) => x.diff?.locality?.target_unique);
      const parts = g.filter((x) => x.diff).map((x) => x.diff.parts_changed_semantic.length + x.diff.parts_added.length);
      return [k, g.length, g.filter((x) => x.save?.ok).length, gu.length ? fmtRate(wilson(gu.filter((x) => x.diff.locality.nontarget_changed === 0).length, gu.length)) : 'n/a', gu.length ? fmtRate(wilson(gu.filter((x) => (x.diff.locality.intended_text_ok ?? true) && (x.diff.locality.intended_raw_ok ?? true) && x.diff.locality.target_changed !== false).length, gu.length)) : 'n/a', f1(med(parts))];
    });
    const fails = edits.filter((x) => !x.save?.ok).map((x) => `${x.doc} ${x.edit}/${x.pos}: ${x.batch?.ok === false ? 'batch ' + (x.batch.code || '') + ' ' + (x.batch.message || '') : 'save ' + (x.save?.code || '') + ' ' + (x.save?.message || '')}`);
    const noops = e.filter((x) => x.kind === 'first-save-normalization');
    const nrows = noops.map((x) => x.diff);
    const frac = (a, b) => (b ? a / b : 0);
    sec.push(`## E1. Edit correctness and locality\n\n${edits.length} edit trials (${[...new Set(edits.map((x) => x.doc))].length} documents x up to 8 edit types x 4 positions; types that need a table are skipped where a document has none: ${e.filter((x) => x.skipped).length} skipped). Locality compares N (document saved with no edit) and E (document saved with the edit): content view, identifier attributes stripped.\n\n` +
      `- Edit applied and saved: ${fmtRate(wilson(ok.length, edits.length))}\n- Targets that could be located uniquely and analysed: ${uniq.length}\n- **Zero non-target paragraphs changed:** ${fmtRate(wilson(nt0.length, uniq.length))}\n- Text of every non-target paragraph preserved: ${fmtRate(wilson(txt.length, uniq.length))}\n- Intended change present in the target: ${fmtRate(wilson(intended.length, uniq.length))}\n\n` +
      table(['edit type', 'trials', 'saved', 'zero non-target changes', 'intended change present', 'median parts changed'], byEdit) +
      (fails.length ? `\n\n**Trials that did not complete (${fails.length}):**\n\n` + fails.slice(0, 40).map((x) => '- ' + x).join('\n') : '') +
      `\n\n### First-save normalization (A versus N, ${noops.length} documents)\n\nOpening and saving a document written by another program, with no edit, changes it as follows. Median share of paragraphs whose raw XML changes: ${f1(100 * med(nrows.map((d) => frac(d.paragraphs_raw_changed_before, d.paragraphs_before))))}%; share whose content changes once identifier attributes are ignored: ${f1(100 * med(nrows.map((d) => frac(d.paragraphs_content_changed_before, d.paragraphs_before))))}%. Median number of parts with different bytes: ${f1(med(nrows.map((d) => d.parts_changed.length)))}; with different XML content: ${f1(med(nrows.map((d) => d.parts_changed_semantic.length)))}. Text identical after save in ${nrows.filter((d) => d.text_before === d.text_after).length}/${nrows.length} documents.\n`);
    hyp.push(['H2', '0 non-target paragraph content changes (N vs E)', `${nt0.length}/${uniq.length}`, nt0.length === uniq.length ? 'met' : 'NOT met']);
  }
  const c = rd('e1_control');
  if (c.length) sec.push(`### E1 control: plain parse-and-reserialize of every XML part (no product, ${c.length} documents)\n\nMedian parts with different bytes: ${f1(med(c.map((x) => x.parts_changed)))} of ${f1(med(c.map((x) => x.parts_total)))}; with different XML content: ${f1(med(c.map((x) => x.parts_changed_semantic)))}. Median share of paragraphs with different raw XML: ${f1(100 * med(c.map((x) => x.paragraphs_raw_changed / x.paragraphs)))}%; content-level: ${f1(100 * med(c.map((x) => x.paragraphs_content_changed / x.paragraphs)))}%.`);
}

// ---------------- oracles ----------------
for (const x of ['e1', 'c1', 'c2']) {
  const o = rd('oracles_' + x);
  if (!o.length) continue;
  const pass = o.filter((r) => r.pass_O1_O5);
  const per = ['O1', 'O2', 'O3', 'O4', 'O5'].map((k) => `${k} ${o.filter((r) => r[k]).length}/${o.length}`).join(', ');
  const fails = o.filter((r) => !r.pass_O1_O5).slice(0, 20).map((r) => `- ${path.basename(r.file)}: ${['O1', 'O2', 'O3', 'O4', 'O5'].filter((k) => !r[k]).map((k) => k + ' ' + JSON.stringify(r[k + '_err'] ?? '')).join('; ')}`);
  sec.push(`## Independent oracles on ${x.toUpperCase()} output\n\n${fmtRate(wilson(pass.length, o.length))} pass O1 to O5 (${per}).` + (fails.length ? '\n\nFailures:\n\n' + fails.join('\n') : ''));
  summary['oracles_' + x] = { n: o.length, pass: pass.length };
  if (x === 'e1') hyp.push(['H3', 'all produced files pass O1-O5 (E1)', `${pass.length}/${o.length}`, pass.length === o.length ? 'met' : 'NOT met']);
}

// ---------------- C2 ----------------
{
  const c = rd('c2_build');
  if (c.length) {
    const rows = c.map((r) => [r.recipe, r.batch?.ok ? 'yes' : 'NO ' + (r.batch?.code || ''), r.save?.ok ? 'yes' : (r.batch?.ok ? 'NO ' + (r.save?.code || '') : '-'), r.save?.valid === true ? 'valid' : String(r.save?.valid ?? '-'), (r.feature_missing || []).length ? 'MISSING ' + r.feature_missing.join(',') : (r.save?.ok ? 'present' : '-')]);
    const done = c.filter((r) => r.save?.ok && !(r.feature_missing || []).length);
    sec.push(`## C2. Building with typed operations (${c.length} recipes)\n\n${fmtRate(wilson(done.length, c.length))} recipes built, saved and contain the intended structure.\n\n` + table(['recipe', 'batch applied', 'saved', 'gate verdict', 'intended structure'], rows) + '\n\n' + c.filter((r) => !r.batch?.ok).map((r) => `- ${r.recipe}: ${r.batch?.code} ${r.batch?.message}`).join('\n'));
    summary.c2 = { recipes: c.length, complete: done.length };
  }
}

const DOC_ACCEPT = { 'table-without-separator': 'no separator row means plain text, as in Markdown', 'deep-nested-quote': 'documented: quote levels beyond 3 are written at level 3' };
// ---------------- C1 ----------------
{
  const c = rd('c1_create').filter((x) => x.kind === 'compose');
  if (c.length) {
    const ok = c.filter((x) => x.save?.ok);
    const fid = ok.filter((x) => x.fidelity && !x.fidelity.error);
    const tot = fid.reduce((a, x) => a + x.fidelity.expected, 0), found = fid.reduce((a, x) => a + x.fidelity.found_in_order, 0);
    const attr = fid.reduce((a, x) => a + x.fidelity.attr_checked, 0), attrOk = fid.reduce((a, x) => a + x.fidelity.attr_ok, 0);
    const ratios = ok.map((x) => x.write_cost_ratio);
    const rows = c.map((x) => [x.name, x.template, x.md_tokens, x.out_xml_tokens ?? '-', f1(x.write_cost_ratio), ms(x.compose?.ms), ms(x.save?.ms), x.save?.valid === true ? 'valid' : String(x.save?.valid ?? '-'), x.fidelity ? `${x.fidelity.found_in_order}/${x.fidelity.expected}` : '-']);
    sec.push(`## C1. Creating documents from Markdown (${c.length} trials)\n\nCompose, validate and save succeeded in ${fmtRate(wilson(ok.length, c.length))}. Content fidelity (independent reader, blocks found in order): **${fmtRate(wilson(found, tot))}**; heading level / list numbering / table shape / quote style correct on found blocks: ${fmtRate(wilson(attrOk, attr))}. Writing cost: the Markdown an agent writes is a median ${f1(med(ratios))}x fewer tokens than the \`word/document.xml\` that results (range ${f1(Math.min(...ratios))} to ${f1(Math.max(...ratios))}).\n\n` + table(['trial', 'template', 'md tokens', 'xml tokens', 'xml/md', 'compose', 'save', 'gate', 'fidelity'], rows));
    summary.c1 = { trials: c.length, saved: ok.length, blocks_expected: tot, blocks_found: found };
    hyp.push(['H8', 'C1 content fidelity >= 98% of blocks', `${(100 * found / tot).toFixed(2)}%`, found / tot >= 0.98 ? 'met' : 'NOT met']);
    const oos = rd('c1_create').filter((x) => x.kind === 'out-of-scope');
    if (oos.length) sec.push(`### C1b. Out-of-scope Markdown must fail closed (${oos.length} inputs)\n\n` + table(['input', 'result', 'code', 'document unchanged', 'note'], oos.map((x) => [x.name, x.accepted ? (DOC_ACCEPT[x.name] ? 'accepted (documented)' : 'ACCEPTED') : 'refused', x.code || '-', x.accepted ? '-' : (x.unchanged ? 'yes' : 'NO'), x.accepted ? (DOC_ACCEPT[x.name] || '') : (x.message || '').replace(/\|/g, '/').slice(0, 90)])) + `\n\nRefused with nothing written: ${oos.filter((x) => !x.accepted && x.unchanged).length} of ${oos.filter((x) => !DOC_ACCEPT[x.name]).length} inputs that are outside the documented subset.`);
  }
}

// ---------------- E3 ----------------
{
  const g = rd('e3_gate');
  const A = g.filter((x) => x.phase === 'A');
  if (A.length) {
    const clean = A.filter((x) => x.group === 'clean');
    const baseErr = Object.fromEntries(clean.map((x) => [x.base, x.validate?.errorCount ?? 0]));
    const preInvalid = clean.filter((x) => (x.validate?.errorCount ?? 0) > 0);
    const faults = A.filter((x) => x.group === 'fault');
    const ctrl = A.filter((x) => x.group === 'control');
    const baseOf = (x) => baseErr[x.base + '.docx'] ?? 0;
    const fu = new Map(rd('e3_followup').map((x) => [x.label, x]));
    const outcome = (x) => {
      if (!x.open.ok) return 'refused at open';
      if ((x.validate?.errorCount ?? 0) > baseOf(x)) return 'flagged invalid';
      if (x.validate?.pending_normalization?.length) return 'repaired at save';
      const f = fu.get(x.label);
      if (f && f.raw.errors > baseOf(x) && (f.default_output?.errors ?? 0) <= baseOf(x)) return 'repaired at save';
      return 'not flagged';
    };
    const caught = (x) => ['refused at open', 'flagged invalid'].includes(outcome(x));
    const safe = (x) => outcome(x) !== 'not flagged';
    const cats = [...group(faults, (x) => x.category).entries()].sort();
    const rows = cats.map(([k, gr]) => [k, gr.length, gr.filter((x) => outcome(x) === 'refused at open').length, gr.filter((x) => outcome(x) === 'flagged invalid').length, gr.filter((x) => outcome(x) === 'repaired at save').length, gr.filter((x) => outcome(x) === 'not flagged').length]);
    const schemaF = faults.filter((x) => x.kind !== 'semantic');
    const cleanBase = faults.filter((x) => baseOf(x) === 0 && x.strict);
    const strictRef = cleanBase.filter((x) => x.strict.ok === false && caught(x));
    const strictWritten = faults.filter((x) => x.strict?.file_changed && caught(x) && baseOf(x) === 0).length;
    const defaultWrote = faults.filter((x) => caught(x) && baseOf(x) === 0 && x.default?.file_changed);
    const falseAlarm = ctrl.filter((x) => !x.open.ok || (x.validate?.errorCount ?? 0) > baseOf(x) || (baseOf(x) === 0 && x.strict?.ok === false));
    const missed = faults.filter((x) => outcome(x) === 'not flagged');
    const upSilent = missed.filter((x) => (fu.get(x.label)?.raw.errors ?? 0) <= baseOf(x));
    const dOut = faults.filter((x) => fu.get(x.label)?.default_output);
    const dOutInvalid = dOut.filter((x) => fu.get(x.label).default_output.errors > baseOf(x));
    summary.e3 = { faults: faults.length, caught: faults.filter(caught).length, repaired: faults.filter((x) => outcome(x) === 'repaired at save').length, not_flagged: missed.length, controls: ctrl.length, false_alarms: falseAlarm.length };
    sec.push(`## E3. Schema gate: injected faults\n\n${faults.length} fault files (60 fault classes x 3 base documents; a class that does not apply to a base is skipped). Each file was opened, validated, then saved after a trivial edit in strict mode and in default mode. "Flagged invalid" means \`validate_document\` reports more errors than the same base document had before injection.\n\n` +
      `- **Caught (refused at open or flagged invalid):** ${fmtRate(wilson(faults.filter(caught).length, faults.length))}; excluding the semantic cross-reference class: ${fmtRate(wilson(schemaF.filter(caught).length, schemaF.length))}\n- **Caught or repaired by the save-time normalizer** (element-order faults are put back in schema order and the file is written valid): ${fmtRate(wilson(faults.filter(safe).length, faults.length))}; excluding the semantic class: ${fmtRate(wilson(schemaF.filter(safe).length, schemaF.length))}\n- Strict save refused the caught faults (bases with no pre-existing errors, ${cleanBase.length} files): ${fmtRate(wilson(strictRef.length, cleanBase.filter(caught).length))}; strict saves that changed the file on disk: ${strictWritten}\n- Default mode wrote the file for ${defaultWrote.length} of ${cleanBase.filter(caught).length} caught faults: default mode is documented to accept a save that adds no new error signature, and these faults were already in the input file; it refuses edits that add new errors (phase B below).\n- **False alarms** on valid-mutation controls (more errors than their base, or a strict refusal on an error-free base): ${fmtRate(wilson(falseAlarm.length, ctrl.length))}\n- Clean corpus documents that already carry a schema error before any injection: ${preInvalid.length} of ${clean.length}` + (preInvalid.length ? ` (${preInvalid.map((x) => x.base).join(', ')}; the validator reports ${preInvalid[0].validate.first[0]?.replace(/http:\/\/schemas[^ ']*/g, 'w:')} in the writer's numbering definitions; strict mode refuses to save them, default mode saves them and reports the error)` : '') + `\n\n` +
      table(['category', 'files', 'refused at open', 'flagged invalid', 'repaired at save', 'not flagged'], rows) + `\n\n### Not flagged (${missed.length})\n\nThese are published, not hidden. ${upSilent.length} of the ${missed.length} are also accepted by the standalone upstream validator run directly on the injected file (no ASTra code in the path), so they are limits of the upstream rules (or of how I labelled a fault as invalid), not something ASTra added. The semantic class (cross-references) is outside what an XSD validator checks.\n\nWhat default mode writes: ${dOutInvalid.length} of ${dOut.length} files written by default-mode saves of these fault files still contain the injected error, because default mode is documented to tolerate errors that were already in the input and to refuse only edits that add new ones. Use strict mode to refuse such inputs (strict refused 100% of the caught faults above).\n\n` + (missed.length ? table(['id', 'base', 'kind', 'fault', 'rule'], missed.map((x) => [x.id, x.base, x.kind, x.title, x.rule])) : 'none'));
    hyp.push(['H4', 'gate catches >= 95% of XSD/package faults, 0 false alarms', `${(100 * schemaF.filter(caught).length / schemaF.length).toFixed(1)}% caught (${(100 * schemaF.filter(safe).length / schemaF.length).toFixed(1)}% caught or repaired), ${falseAlarm.length} false alarms`, schemaF.filter(caught).length / schemaF.length >= 0.95 && falseAlarm.length === 0 ? 'met' : 'NOT met']);
  }
  const B = g.filter((x) => x.phase === 'B');
  if (B.length) {
    const o = rd('e3_phaseB_oracles');
    const oByFile = new Map(o.map((x) => [x.file, x]));
    const rows = B.map((x) => [x.group, x.name, x.outcome === 'saved' ? 'accepted and saved' : x.outcome === 'blocked-at-save' ? 'blocked at save' : 'rejected at edit', x.batch?.code || x.save?.code || '-', x.written ? (oByFile.get(x.file)?.pass_O1_O5 ? 'written, passes O1-O5' : 'written, FAILS oracles') : 'nothing written']);
    const invalidWritten = B.filter((x) => x.written && !(oByFile.get(x.file)?.pass_O1_O5));
    sec.push(`### E3 phase B. Edits that try to introduce invalid content (${B.length} edits on a valid document, default mode)\n\n` + table(['route', 'edit', 'outcome', 'code', 'file'], rows) + `\n\nOutcome counts: ` + [...group(B, (x) => x.outcome).entries()].map(([k, v]) => `${k} ${v.length}`).join(', ') + `. Every file that was written passed O1 to O5 (${B.filter((x) => x.written).length - invalidWritten.length}/${B.filter((x) => x.written).length}).\n\nEdits that were accepted and saved are schema-valid even when they are unwise (font size 0 or 99999, a negative margin, a chart with no series, a bookmark name with spaces, a style name the document does not define, a table with ragged rows). The gate answers "is this valid OOXML", not "is this sensible".`);
  }
}

// ---------------- E4 ----------------
{
  const a = rd('e4_atomicity');
  if (a.length) {
    const failed = a.filter((x) => !x.batch_ok);
    const accepted = a.filter((x) => x.batch_ok);
    const atomic = failed.filter((x) => x.atomic);
    sec.push(`## E4. Atomicity\n\n${a.length} trials: batches of 2, 5, 20 or 100 operations in which one operation (first, middle or last) is bad, in four ways. ${failed.length} batches were rejected and ${accepted.length} were accepted (the "unknown style" operation is not an error: ASTra applies a style name that the document does not define; see limits). Of the ${failed.length} rejected batches, **the session was unchanged afterwards (outline, revision, dirty flag, history length) in ${fmtRate(wilson(atomic.length, failed.length))}.** The same batch without the bad operation applied afterwards in ${fmtRate(wilson(a.filter((x) => x.recovery_ok).length, a.length))}.\n\n` + table(['failure kind', 'trials', 'rejected', 'rejected by', 'session unchanged after rejection'], [...group(a, (x) => x.kind).entries()].map(([k, g]) => [k, g.length, g.filter((x) => !x.batch_ok).length, [...new Set(g.filter((x) => !x.batch_ok).map((x) => x.code))].join(', ') || 'not rejected', g.filter((x) => !x.batch_ok).length ? `${g.filter((x) => x.atomic).length}/${g.filter((x) => !x.batch_ok).length}` : 'n/a'])) + `\n\nNote: "TOOL_INPUT_SCHEMA" means the MCP tool's own input schema refused the call before the engine ran, so those trials show the contract, not the engine's rollback. The engine-level rollback trials are the ${failed.filter((x) => x.code === 'E_VALIDATION' || x.code === 'E_NOT_FOUND').length} rejected with E_VALIDATION or E_NOT_FOUND.`);
    hyp.push(['H5', '100% atomic outcomes among rejected batches', `${atomic.length}/${failed.length}`, atomic.length === failed.length ? 'met' : 'NOT met']);
  }
}

// ---------------- E5 ----------------
{
  const h = rd('e5_history');
  if (h.length) {
    const rv = h.flatMap((x) => x.reverts.map((r) => ({ ...r, name: x.name })));
    const st = h.flatMap((x) => x.steps.map((r) => ({ ...r, name: x.name })));
    const eq = rv.filter((r) => r.all_parts_equal), fe = rv.filter((r) => r.file_equal), de = rv.filter((r) => r.doc_xml_equal);
    const rf = rv.filter((r) => r.redo?.all_parts_equal), pm = rv.filter((r) => r.preview_matches_actual);
    const seq = h.length;
    sec.push(`## E5. History, revert and redo\n\n${seq} sequences of 10 saved edits on 10 documents x 3 seeds.\n\n**Revert to a saved commit (by commit id), every depth 1 to 10, each followed by redo: ${rv.length} reverts.**\n\n- Restored package parts identical to the state saved at that commit (SHA-256 of every part): ${fmtRate(wilson(eq.length, rv.length))}\n- Restored \`word/document.xml\` identical: ${fmtRate(wilson(de.length, rv.length))}\n- **Parts reachable through relationships identical** (what Word or any reader loads): ${fmtRate(wilson(rv.filter((r) => r.reachable_equal).length, rv.length))}\n- Parts left behind or missing after the revert (name: reverts): ${(() => { const t = new Map(); for (const r of rv) for (const k of [...(r.extra_parts || []).map((x) => 'left behind ' + x), ...(r.missing_parts || []).map((x) => 'missing ' + x), ...(r.differing_parts || []).map((x) => 'differs ' + x)]) t.set(k, (t.get(k) || 0) + 1); return [...t.entries()].map(([k, v]) => k + ': ' + v).join('; ') || 'none'; })()}\n- Whole file byte-identical, including the ZIP container: ${fmtRate(wilson(fe.length, rv.length))}\n- Redo returns every part of the head state identically: ${fmtRate(wilson(rf.length, rv.length))}\n- \`preview_revert\` file list equals the parts that actually changed: ${fmtRate(wilson(pm.length, rv.length))}\n- Every save appears in the history listing: ${h.filter((x) => x.history.saves_all_listed).length}/${seq}; listing order matches save order: ${h.filter((x) => x.history.listed_order_matches).length}/${seq}\n\n**Revert by step count (\`steps\`), fresh history per depth 1, 2, 5, 9 (first seed of each document): ${st.length} reverts.** Every part identical: ${fmtRate(wilson(st.filter((r) => r.all_parts_equal).length, st.length))}; reachable parts identical: ${fmtRate(wilson(st.filter((r) => r.reachable_equal).length, st.length))}; preview names the intended commit: ${fmtRate(wilson(st.filter((r) => r.preview_target_message && r.preview_target_message.startsWith(r.expected_message)).length, st.length))}.\n`);
    hyp.push(['H6', '100% byte-exact restores', `document.xml ${de.length}/${rv.length}; every part ${eq.length}/${rv.length}; reachable parts ${rv.filter((r) => r.reachable_equal).length}/${rv.length}; whole file ${fe.length}/${rv.length}`, fe.length === rv.length ? 'met' : 'NOT met']);
    const orp = rd('e5_orphans'), orpO = rd('e5_orphans_oracles');
    if (orp.length) sec.push(`### E5 follow-up: files reverted past a comment (${orp.length} documents)\n\n${orp.some((x) => 'has_people_after_revert' in x) ? `After the revert, \`word/people.xml\` is still in the ZIP in ${orp.filter((x) => x.has_people_after_revert).length}/${orp.length}, \`word/commentsExtended.xml\` in ${orp.filter((x) => x.has_comments_extended_after_revert).length}/${orp.length} and \`word/comments.xml\` in ${orp.filter((x) => x.has_comments_after_revert).length}/${orp.length}; the revert target is the state saved before the comment was added.` : 'Parts such as \`commentsExtended.xml\` and \`people.xml\` may stay in the ZIP after a revert past a comment (not recorded by this older run).'} Independent checks on the reverted files: O1 to O5 pass ${orpO.filter((x) => x.pass_O1_O5).length}/${orpO.length}; standalone validator error count equals the count before the comment was added in ${orp.filter((x) => x.errors_before_comment === x.errors_after_revert).length}/${orp.length}.`);
    const gr = rd('e5_growth');
    if (gr.length) sec.push(`### History store growth\n\n` + table(['document', 'saves', 'docx bytes', 'history bytes', 'history / docx'], gr.map((x) => [x.base, x.saves, x.docx_bytes, x.history_bytes, f2(x.history_bytes / x.docx_bytes)])));
  }
}

// ---------------- E6/E7 ----------------
{
  const p = rd('e6_performance');
  if (p.length) {
    const timed = p.filter((x) => x.phase === 'timed' && !x.error);
    const rows = [...group(timed, (x) => x.base).entries()].map(([b, g]) => {
      const cell = (k) => { const s = summarize(g.map((x) => x[k])); return `${ms(s.median)} (p95 ${ms(s.p95)})`; };
      const mem = p.find((x) => x.base === b && x.phase === 'memory');
      return [b, g.length, cell('open_ms'), cell('inspect_ms'), cell('search_ms'), cell('edit1_ms'), cell('edit50_ms'), cell('validate_ms'), cell('save_ms'), f1(100 * med(g.map((x) => x.validator_share_of_save))) + '%', mem ? (mem.rss_peak_kb / 1024).toFixed(0) + ' MB' : '-'];
    });
    const cold = p.filter((x) => x.phase === 'cold').map((x) => [x.base, ms(x.open_ms), ms(x.save_ms)]);
    sec.push(`## E6. Latency and memory by document size\n\n15 timed runs after 2 warm-up runs per size, median (p95). One machine, otherwise idle. Memory is the peak resident size of the server process.\n\n` + table(['document', 'n', 'open', 'inspect', 'search', '1 edit', '50-op batch', 'validate', 'save', 'validator share of save', 'peak RSS'], rows) + `\n\n"Validator share of save" is the separately timed validate call divided by the save time. It is an estimate: a value near or above 100% means the schema check is essentially the whole cost of a save and the two timings differ only by noise.\n\nCold first run (new server process): ` + cold.map((c) => `${c[0]} open ${c[1]}, save ${c[2]}`).join('; ') + '.');
    const s100 = timed.filter((x) => x.base === 'pd-prose-100p').map((x) => x.save_ms);
    if (s100.length) hyp.push(['H7', 'median save < 5 s at 100 pages', ms(med(s100)), med(s100) < 5000 ? 'met' : 'NOT met']);
  }
  const b = rd('e7_batch_scaling');
  if (b.length) sec.push(`### E7. Batch scaling (100-page document, replace operations, 10 runs each)\n\n` + table(['operations', 'runs', 'batch median', 'batch p95', 'ms per operation', 'save median'], [...group(b, (x) => x.ops).entries()].map(([k, g]) => [k, g.length, ms(med(g.map((x) => x.batch_ms))), ms(summarize(g.map((x) => x.batch_ms)).p95), f1(med(g.map((x) => x.ms_per_op))), ms(med(g.map((x) => x.save_ms)))])));
}

let env = '';
if (fs.existsSync(path.join(RESULTS, 'env.json'))) { const e = JSON.parse(fs.readFileSync(path.join(RESULTS, 'env.json'), 'utf8')); env = '\n\n```json\n' + JSON.stringify(e, null, 1) + '\n```'; }
const head = `# ASTra benchmark report: release ${RELEASE}\n\n> **Work in progress.** ASTra is improved every day; these numbers describe one published release on one machine and will change. The protocol (\`PROTOCOL.md\`) was committed before the first measurement. Raw per-trial data sits next to this file as \`*.jsonl\`. Nothing was dropped.\n\n## Pre-registered expectations versus results\n\n` + table(['id', 'expectation', 'measured', 'verdict'], hyp) + '\n\n## Environment' + env;
fs.writeFileSync(path.join(RESULTS, 'REPORT.md'), head + '\n\n' + sec.join('\n\n') + '\n');
fs.writeFileSync(path.join(RESULTS, 'summary.json'), JSON.stringify(summary, null, 1));
console.log('wrote', path.join(RESULTS, 'REPORT.md'), 'sections', sec.length);
