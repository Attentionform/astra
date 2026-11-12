import { describe, it, expect, afterEach, vi } from 'vitest';
import * as fsp from 'fs/promises';
import os from 'os';
import path from 'path';
import { createHash } from 'crypto';
import * as engine from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS, GIT_HISTORY_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleExportOriginalDocument, handleRestoreOriginalDocument } from '../../src/lib/mcp/handlers/original';

const dirs: string[] = [];
const sha = (b: Buffer) => createHash('sha256').update(b).digest('hex');
afterEach(async () => { vi.restoreAllMocks(); await Promise.all(dirs.splice(0).map(p => fsp.rm(p, { recursive:true, force:true }))); });
async function setup() {
  const dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-original-mcp-')); dirs.push(dir);
  const file = path.join(dir, 'report.docx');
  const original = engine.createBlankDocxBuffer('blank'); await fsp.writeFile(file, original);
  const store = createMcpSessionStore();
  const opened = await LEAN6_HANDLERS.open_document(store, { path:file, validationTimeoutMs:0 });
  if (!opened.ok) throw new Error(JSON.stringify(opened));
  const sessionId = opened.sessionId;
  return { dir,file,original,store,sessionId };
}
async function edit(s: Awaited<ReturnType<typeof setup>>) {
  const h=s.store.get(s.sessionId); const id=h.ast.sections.flatMap(x=>x.nodes).find(n=>n.type==='paragraph')!.id;
  const result=await LEAN6_HANDLERS.execute_batch(s.store,{sessionId:s.sessionId,operations:[{action:'replace',target:id,text:'Changed after opening'}]});
  expect(result.ok,JSON.stringify(result)).toBe(true);
}
async function save(s: Awaited<ReturnType<typeof setup>>) {
 const r=await LEAN6_HANDLERS.save_document(s.store,{sessionId:s.sessionId});expect(r.ok,JSON.stringify(r)).toBe(true); return r;
}

describe('exact-original recovery through MCP',()=>{
 it('captures raw bytes at open and exports them after edits without changing the session', async()=>{
  const s=await setup();expect(s.store.get(s.sessionId).originalInfo?.available).toBe(true);
  await edit(s);await save(s);
  const before=await fsp.readFile(s.file);const h=s.store.get(s.sessionId);const rev=h.revision;
  const out=path.join(s.dir,'original.docx');
  const result=await handleExportOriginalDocument(s.store,{sessionId:s.sessionId,outPath:out});
  expect(result.ok,JSON.stringify(result)).toBe(true);
  expect(sha(await fsp.readFile(out))).toBe(sha(s.original));expect((await fsp.readFile(s.file)).equals(before)).toBe(true);expect(h.revision).toBe(rev);
  const history=await GIT_HISTORY_HANDLERS.get_document_history(s.store,{sessionId:s.sessionId});
  expect(history.ok && history.original.available).toBe(true);
 });
 it('restores exact bytes immediately after a legitimate save; normal revert and redo remain usable',async()=>{
  const s=await setup();await edit(s);await save(s);
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});
  expect(result.ok,JSON.stringify(result)).toBe(true);expect(sha(await fsp.readFile(s.file))).toBe(sha(s.original));
  const h=s.store.get(s.sessionId);expect(h.lastWrittenSha256).toBe(sha(s.original));expect(h.savedRevision).toBe(h.revision);
  const reverted=await GIT_HISTORY_HANDLERS.revert_document(s.store,{sessionId:s.sessionId,steps:1});expect(reverted.ok,JSON.stringify(reverted)).toBe(true);
  expect(s.store.get(s.sessionId).parser.readDocumentXml()).toContain('Changed after opening');
  const redo=await GIT_HISTORY_HANDLERS.redo_document(s.store,{sessionId:s.sessionId});expect(redo.ok,JSON.stringify(redo)).toBe(true);
  expect(s.store.get(s.sessionId).parser.readDocumentXml()).not.toContain('Changed after opening');
  await save(s);
 });
 it('refuses unsaved edits without discarding them, and accepts explicit discard',async()=>{
  const s=await setup();await edit(s);const rev=s.store.get(s.sessionId).revision;
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(result.ok).toBe(false);expect(s.store.get(s.sessionId).revision).toBe(rev);
  const confirmed=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId,discard:true});expect(confirmed.ok,JSON.stringify(confirmed)).toBe(true);expect(sha(await fsp.readFile(s.file))).toBe(sha(s.original));
 });
 it('refuses external modification after an ASTra save',async()=>{
  const s=await setup();await edit(s);await save(s);const external=Buffer.from('outside change');await fsp.writeFile(s.file,external);
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(result.ok).toBe(false);expect((await fsp.readFile(s.file)).equals(external)).toBe(true);
 });
 it('refuses export over the active document even with overwrite',async()=>{
  const s=await setup();const result=await handleExportOriginalDocument(s.store,{sessionId:s.sessionId,outPath:s.file,overwrite:true});expect(result.ok).toBe(false);
 });
 it('retains dirty session and disk bytes when original policy rejects during preparation',async()=>{
  const s=await setup();await edit(s);const parser=s.store.get(s.sessionId).parser;
  const read=vi.spyOn(engine,'readOriginal').mockRejectedValueOnce(new Error('corrupt original'));
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId,discard:true});expect(result.ok).toBe(false);expect(read).toHaveBeenCalled();
  expect(s.store.get(s.sessionId).parser).toBe(parser);expect((await fsp.readFile(s.file)).equals(s.original)).toBe(true);
 });
 it('rolls back disk and leaves session and history unchanged when the restore commit fails',async()=>{
  const s=await setup();await edit(s);await save(s);const before=await fsp.readFile(s.file);const h=s.store.get(s.sessionId);const parser=h.parser;const rev=h.revision;const hist=await engine.getHistory(s.file);
  vi.spyOn(engine,'commitChange').mockRejectedValueOnce(new Error('injected commit failure'));
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(result.ok).toBe(false);expect((await fsp.readFile(s.file)).equals(before)).toBe(true);expect(h.parser).toBe(parser);expect(h.revision).toBe(rev);expect((await engine.getHistory(s.file)).map(x=>x.oid)).toEqual(hist.map(x=>x.oid));
 });
 it('memory-label sessions have no original and cannot restore',async()=>{
  const store=createMcpSessionStore();const h=await store.openBuffer(engine.createBlankDocxBuffer('blank'));
  expect(h.originalInfo).toEqual({available:false,reason:'not-applicable'});
  const r=await handleRestoreOriginalDocument(store,{sessionId:h.sessionId});expect(r.ok).toBe(false);
 });
 it('explicit discard does not put unsaved text into redo',async()=>{
  const s=await setup();await save(s);await edit(s);
  const restored=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId,discard:true});expect(restored.ok,JSON.stringify(restored)).toBe(true);
  const redone=await GIT_HISTORY_HANDLERS.redo_document(s.store,{sessionId:s.sessionId});expect(redone.ok,JSON.stringify(redone)).toBe(true);
  expect(s.store.get(s.sessionId).parser.readDocumentXml()).not.toContain('Changed after opening');
 });
 it('a failed replacement leaves dirty session, original file and history untouched',async()=>{
  const s=await setup();await save(s);await edit(s);const h=s.store.get(s.sessionId);const parser=h.parser;const before=await fsp.readFile(s.file);const head=(await engine.getHistory(s.file))[0].oid;
  vi.spyOn(engine,'replaceFileAtomic').mockRejectedValueOnce(Object.assign(new Error('injected partial temporary write'),{outcome:'not-replaced'}));
  const restored=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId,discard:true});expect(restored.ok).toBe(false);
  expect(h.parser).toBe(parser);expect(h.revision).not.toBe(h.savedRevision);expect(h.parser.readDocumentXml()).toContain('Changed after opening');expect((await fsp.readFile(s.file)).equals(before)).toBe(true);expect((await engine.getHistory(s.file))[0].oid).toBe(head);
 });
 it('failed rollback retains a verified recovery backup and names it',async()=>{
  const s=await setup();await edit(s);await save(s);const before=await fsp.readFile(s.file);const h=s.store.get(s.sessionId);const parser=h.parser;
  const replace=engine.replaceFileAtomic;
  vi.spyOn(engine,'replaceFileAtomic').mockImplementationOnce(replace).mockRejectedValueOnce(new Error('rollback volume unavailable'));
  vi.spyOn(engine,'commitChange').mockRejectedValueOnce(new Error('commit volume unavailable'));
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(result.ok).toBe(false);
  const backups=(await fsp.readdir(s.dir)).filter(n=>n.includes('.astra-restore-backup-'));expect(backups).toHaveLength(1);
  expect((await fsp.readFile(path.join(s.dir,backups[0]))).equals(before)).toBe(true);expect(JSON.stringify(result)).toContain(backups[0]);expect(h.parser).toBe(parser);
 });
 it('corrupt retained original never replaces the working file',async()=>{
  const s=await setup();const info=await engine.getOriginalInfo(s.file);if(!info.available)throw new Error('capture unavailable');
  await fsp.writeFile(path.join(engine.getRepoDir(s.file),'.git','astra-original',`original-${info.sha256}.docx`),'corrupt');
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(result.ok).toBe(false);expect((await fsp.readFile(s.file)).equals(s.original)).toBe(true);
 });

});
