#!/usr/bin/env node
/**
 * The visual workflows a user runs, end to end on the machine this runs on: the built command line
 * (`astra visual`, `astra inspect --visual`, `astra export --format pdf`) and the built MCP server
 * over stdio (`visual_inspect` on an open session, after an edit that is not saved). Each page image
 * must be a PNG with visible ink, so a blank render cannot pass.
 *
 * Run after building the engine, CLI and MCP server. Used by the `visual-render` job of ci.yml.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { countInk } from './lib/png-ink.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const cli = path.join(root, 'packages', 'cli', 'dist', 'cli.js');
const mcpServer = path.join(root, 'packages', 'mcp-server', 'dist', 'index.js');
const MIN_INK = 300;

function fail(message) {
  console.error(`visual-e2e-smoke: FAIL: ${message}`);
  process.exit(1);
}

function astra(args) {
  const run = spawnSync(process.execPath, [cli, ...args], { encoding: 'utf8', timeout: 240000 });
  if (run.error) fail(`astra ${args[0]}: ${run.error.message}`);
  return { status: run.status, stdout: run.stdout ?? '', stderr: run.stderr ?? '' };
}

function mustAstra(args) {
  const run = astra(args);
  if (run.status !== 0) fail(`astra ${args.join(' ')} exited ${run.status}\n${run.stdout}\n${run.stderr}`);
  return run;
}

function inkOf(label, file) {
  const bytes = fs.readFileSync(file);
  let result;
  try {
    result = countInk(bytes);
  } catch (error) {
    return fail(`${label}: ${file} is not a readable PNG: ${error.message}`);
  }
  if (result.ink < MIN_INK) fail(`${label}: page image has only ${result.ink} dark pixels (blank render?)`);
  console.log(`${label}: ok, ${result.width}x${result.height} PNG with ${result.ink} dark pixels`);
  return result.ink;
}

const work = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-visual-e2e-')));
const docx = path.join(work, 'doc.docx');
fs.writeFileSync(
  path.join(work, 'in.md'),
  '# ASTra Visual Smoke\n\nThis page must show readable text, so a blank render cannot pass this check.\n'
);

try {
  // Command line.
  mustAstra(['create', docx, '--from-md', path.join(work, 'in.md'), '--allowed-dir', work]);

  const pagePng = path.join(work, 'page.png');
  const visual = mustAstra(['visual', docx, '--out', pagePng]);
  if (!/VISUAL OK/.test(visual.stdout)) fail(`astra visual did not report VISUAL OK:\n${visual.stdout}`);
  inkOf('astra visual', pagePng);

  const inspect = mustAstra(['inspect', docx, '--visual']);
  const shown = /visual: page-(\d+) (.+) \(\d+B/.exec(inspect.stdout);
  if (!shown) fail(`astra inspect --visual printed no page image line:\n${inspect.stdout}`);
  inkOf('astra inspect --visual', shown[2]);

  const pdfOut = path.join(work, 'doc.pdf');
  mustAstra(['export', docx, '--format', 'pdf', '-o', pdfOut]);
  const pdf = fs.readFileSync(pdfOut);
  if (pdf.length < 500 || pdf.subarray(0, 5).toString('latin1') !== '%PDF-') fail('astra export --format pdf did not write a PDF');
  console.log(`astra export pdf: ok, ${pdf.length} bytes`);

  // A file whose only errors are element order is exported from a verified temporary copy; the file is not changed.
  const orderSource = path.join(work, 'word-order.docx');
  fs.copyFileSync(path.join(root, 'tests', 'fixtures', 'word-order.docx'), orderSource);
  const orderBefore = createHash('sha256').update(fs.readFileSync(orderSource)).digest('hex');
  const orderPdf = path.join(work, 'word-order.pdf');
  const orderRun = mustAstra(['export', orderSource, '--format', 'pdf', '-o', orderPdf]);
  if (!/note: the source has 1 element-order error/.test(orderRun.stdout + orderRun.stderr)) {
    fail(`astra export of word-order.docx printed no repair note:\n${orderRun.stdout}\n${orderRun.stderr}`);
  }
  const orderBytes = fs.readFileSync(orderPdf);
  if (orderBytes.length < 500 || orderBytes.subarray(0, 5).toString('latin1') !== '%PDF-') fail('astra export of word-order.docx did not write a PDF');
  if (createHash('sha256').update(fs.readFileSync(orderSource)).digest('hex') !== orderBefore) fail('astra export changed the source file');
  console.log(`astra export pdf of a file with element-order errors: ok, ${orderBytes.length} bytes, source unchanged`);

  // MCP server over stdio.
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [mcpServer, '--allowed-dir', work],
    env: { ...process.env },
  });
  const client = new Client({ name: 'visual-e2e-smoke', version: '1' });
  await client.connect(transport);
  try {
    const call = async (name, args) => {
      const result = await client.callTool({ name, arguments: args });
      const text = result.content?.[0]?.text ?? '';
      let body;
      try {
        body = JSON.parse(text);
      } catch {
        return fail(`${name} returned non-JSON: ${text.slice(0, 300)}`);
      }
      if (result.isError || body.ok === false) fail(`${name} failed: ${text.slice(0, 600)}`);
      return body;
    };

    const opened = await call('open_document', { path: docx });
    const before = await call('visual_inspect', { sessionId: opened.sessionId, withVisual: true });
    if (!before.images?.length) fail(`visual_inspect returned no images: ${JSON.stringify(before).slice(0, 400)}`);
    const inkBefore = inkOf('MCP visual_inspect', before.images[0].path);

    // Edit the open session, do not save: the render must show the session, not the file on disk.
    const edited = await call('execute_batch', {
      sessionId: opened.sessionId,
      operations: [{ action: 'insert', anchor: 'end', type: 'paragraph', text: 'EDITED IN THE SESSION AND NOT SAVED, WITH MUCH MORE TEXT TO DRAW' }],
    });
    if (edited.appliedCount !== 1) fail(`execute_batch applied ${edited.appliedCount} operations`);
    const after = await call('visual_inspect', { sessionId: opened.sessionId, withVisual: true });
    if (!after.dirty) fail('visual_inspect after an edit did not report the session as dirty');
    const inkAfter = inkOf('MCP visual_inspect (unsaved edit)', after.images[0].path);
    if (inkAfter === inkBefore) fail('the render after an unsaved edit has the same ink as before: the session state was not rendered');
    console.log(`MCP session state: ok, ink ${inkBefore} -> ${inkAfter}`);
  } finally {
    await client.close().catch(() => {});
  }
} finally {
  try {
    fs.rmSync(work, { recursive: true, force: true, maxRetries: 5, retryDelay: 500 });
  } catch {
    /* the runner cleans up */
  }
}
console.log('visual-e2e-smoke: PASS');
