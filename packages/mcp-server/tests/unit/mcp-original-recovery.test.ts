import { describe, it, expect, afterEach, vi } from 'vitest';
import * as fsp from 'fs/promises';
import os from 'os';
import path from 'path';
import { createHash } from 'crypto';
import fs from 'node:fs';
import git from 'isomorphic-git';
import * as engine from '@attentionform/astra-engine';
import { createMcpSessionStore, resolveSessionOriginal } from '../../src/lib/session-store';
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
 it('a session that lost a conflicting first capture never exports or restores another opener’s original',async()=>{
  const s=await setup();const h=s.store.get(s.sessionId);
  h.originalInfo={available:false,reason:'capture-failed',detail:'Another opener recorded a different original for this history.'};h.originalPin=undefined;
  const history=await GIT_HISTORY_HANDLERS.get_document_history(s.store,{sessionId:s.sessionId});
  expect(history.ok && history.original.available).toBe(false);
  const exported=await handleExportOriginalDocument(s.store,{sessionId:s.sessionId,outPath:path.join(s.dir,'wrong.docx')});
  expect(exported.ok).toBe(false);
  const restored=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(restored.ok).toBe(false);
  await save(s);expect(h.originalInfo?.available).toBe(false);
  const again=await handleExportOriginalDocument(s.store,{sessionId:s.sessionId,outPath:path.join(s.dir,'wrong2.docx')});expect(again.ok).toBe(false);
 });
 it('does not let a session opened during initialization adopt another opener’s original later',async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-open-race-'));dirs.push(dir);
  const file=path.join(dir,'racing.docx');const firstBytes=engine.createBlankDocxBuffer('blank');const secondBytes=Buffer.concat([firstBytes,Buffer.from('other input bytes')]);await fsp.writeFile(file,secondBytes);
  const store=createMcpSessionStore();
  // Construct the exact observable state of a competing initializer: a Git
  // directory exists, but no initial commit or original snapshot is published.
  // The second session must pin this ambiguity even after another opener later
  // completes initialization for different input bytes.
  await engine.ensureAstraGitRepo(file);
  await git.init({fs,dir:engine.getRepoDir(file)});
  const second=await store.openBuffer(secondBytes,file);const early=second.originalInfo;
  expect(early?.available).toBe(false);
  await fsp.rm(engine.getRepoDir(file),{recursive:true,force:true});
  const first=await store.openBuffer(firstBytes,file);
  expect(first.originalInfo?.available).toBe(true);
  const history=await GIT_HISTORY_HANDLERS.get_document_history(store,{sessionId:second.sessionId});
  expect(history.ok && history.original.available).toBe(false);
  const exported=await handleExportOriginalDocument(store,{sessionId:second.sessionId,outPath:path.join(dir,'wrong.docx')});
  expect(exported.ok).toBe(false);
  const restored=await handleRestoreOriginalDocument(store,{sessionId:second.sessionId});expect(restored.ok).toBe(false);
 });
 it('a session whose open read "no history yet" before a competitor created it still pins the matching original',async()=>{
  const s=await setup();
  const real=engine.getOriginalInfo;let calls=0;
  vi.spyOn(engine,'getOriginalInfo').mockImplementation(async(...args)=>{calls++;return calls===1?{available:false,reason:'predates-feature'}:real(...args);});
  const second=createMcpSessionStore();const handle=await second.openDocument(s.file);
  expect(handle.originalInfo?.available,JSON.stringify(handle.originalInfo)).toBe(true);
  expect(handle.originalPin?.kind).toBe('retained');
 });
 it('restore is not blocked by a full redo stack and drops redo entries from before the restore',async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-redo-'));dirs.push(dir);
  const file=path.join(dir,'report.docx');const original=engine.createBlankDocxBuffer('blank');await fsp.writeFile(file,original);
  const store=createMcpSessionStore({redoMaxDepth:1});
  const opened=await LEAN6_HANDLERS.open_document(store,{path:file,validationTimeoutMs:0});if(!opened.ok)throw new Error(JSON.stringify(opened));
  const s={dir,file,original,store,sessionId:opened.sessionId};
  await edit(s);await save(s);
  const reverted=await GIT_HISTORY_HANDLERS.revert_document(store,{sessionId:s.sessionId,steps:1});expect(reverted.ok,JSON.stringify(reverted)).toBe(true);
  await save(s);
  const handle=store.get(s.sessionId);expect(handle.redoStack.length).toBe(1);
  const result=await handleRestoreOriginalDocument(store,{sessionId:s.sessionId});
  expect(result.ok,JSON.stringify(result)).toBe(true);expect(sha(await fsp.readFile(file))).toBe(sha(original));
  expect(handle.redoStack.length).toBeLessThanOrEqual(1);
 });
 it('exports the exact original when an allowed document contains a macro part', async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-macro-'));dirs.push(dir);
  const file=path.join(dir,'input.docx');const zip=new engine.AdmZip(engine.createBlankDocxBuffer('blank'));
  zip.addFile('word/vbaProject.bin',Buffer.from('fictional macro content'));const raw=zip.toBuffer();await fsp.writeFile(file,raw);
  const store=createMcpSessionStore();const handle=await store.openDocument(file,undefined,{allowMacros:true});
  const out=path.join(dir,'recovered.docx');const result=await handleExportOriginalDocument(store,{sessionId:handle.sessionId,outPath:out});
  expect(result.ok,JSON.stringify(result)).toBe(true);expect(sha(await fsp.readFile(out))).toBe(sha(raw));
 });
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
 it('can restore the same original twice without losing the original bytes or the session',async()=>{
  const s=await setup();await edit(s);await save(s);
  for(let i=0;i<2;i++){
   const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});
   expect(result.ok,JSON.stringify(result)).toBe(true);
   expect(sha(await fsp.readFile(s.file))).toBe(sha(s.original));
   expect(s.store.get(s.sessionId).lastWrittenSha256).toBe(sha(s.original));
  }
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
 it('leaves disk, session and history unchanged when backup creation fails',async()=>{
  const s=await setup();await edit(s);await save(s);const before=await fsp.readFile(s.file);const h=s.store.get(s.sessionId);const parser=h.parser;const revision=h.revision;const history=(await engine.getHistory(s.file)).map(x=>x.oid);
  vi.spyOn(engine,'createVerifiedBackup').mockRejectedValueOnce(Object.assign(new Error('injected backup creation failure'),{code:'EACCES'}));
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(result.ok).toBe(false);
  expect(await fsp.readFile(s.file)).toEqual(before);expect(h.parser).toBe(parser);expect(h.revision).toBe(revision);expect((await engine.getHistory(s.file)).map(x=>x.oid)).toEqual(history);
  expect((await fsp.readdir(s.dir)).some(name=>name.includes('astra-restore-backup'))).toBe(false);
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
  // The rollback itself is shared engine code; its real failure is injected for every surface in
  // tests/unit/original-restore-parity.test.ts. Here the handler is checked to keep and name the backup.
  vi.spyOn(engine,'rollbackReplacedFile').mockResolvedValueOnce({recovered:false,failure:'rollback volume unavailable'});
  vi.spyOn(engine,'commitChange').mockRejectedValueOnce(new Error('commit volume unavailable'));
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(result.ok).toBe(false);
  const backups=(await fsp.readdir(s.dir)).filter(n=>n.includes('.astra-restore-backup-'));expect(backups).toHaveLength(1);
  expect((await fsp.readFile(path.join(s.dir,backups[0]))).equals(before)).toBe(true);expect(JSON.stringify(result)).toContain(backups[0]);expect(h.parser).toBe(parser);
 });
 it('corrupt retained original never replaces the working file',async()=>{
  const s=await setup();const info=await engine.getOriginalInfo(s.file);if(!info.available)throw new Error('capture unavailable');
  await fsp.writeFile(path.join(engine.getRepoDir(s.file),'.git','astra-original','original.snapshot'),'corrupt');
  const result=await handleRestoreOriginalDocument(s.store,{sessionId:s.sessionId});expect(result.ok).toBe(false);expect((await fsp.readFile(s.file)).equals(s.original)).toBe(true);
 });

 // --- availability contract: what a session pins, and what it may later claim -----------------------------------
 const snapshotPath=(file:string)=>path.join(engine.getRepoDir(file),'.git','astra-original','original.snapshot');
 const originalOf=async(s:Awaited<ReturnType<typeof setup>>)=>{const r=await GIT_HISTORY_HANDLERS.get_document_history(s.store,{sessionId:s.sessionId});if(!r.ok)throw new Error(JSON.stringify(r));return r.original;};
 const initInputs=async(bytes:Buffer)=>{const p=new engine.DocxParser(bytes);await p.getAST();const extras:Record<string,Buffer>={};for(const[n,b]of p.capturePackageSnapshot())if(engine.isTrackedPackageFile(n))extras[n]=b;return{xml:p.readDocumentXml(),extras};};
 /** The observable state of a first initialization that another process has started: .git exists, a lock is held, nothing is committed. */
 const startCompetingInit=async(file:string)=>{await engine.ensureAstraGitRepo(file);await git.init({fs,dir:engine.getRepoDir(file)});const lock=path.join(engine.getRepoDir(file),'.astra-init.lock');await fsp.writeFile(lock,JSON.stringify({version:2,pid:process.pid,hostname:os.hostname(),token:'t',startedAt:new Date().toISOString()}));return lock;};
 it('reports corruption after open instead of a stale available status, with a recovery remedy for export',async()=>{
  const s=await setup();expect((await originalOf(s)).available).toBe(true);
  const file=snapshotPath(s.file);const saved=await fsp.readFile(file);const bytes=Buffer.from(saved);bytes[bytes.length-5]^=0xff;await fsp.writeFile(file,bytes);
  expect(await originalOf(s)).toMatchObject({available:false,reason:'corrupt'});
  const exported=await handleExportOriginalDocument(s.store,{sessionId:s.sessionId,outPath:path.join(s.dir,'x.docx')});
  expect(exported.ok).toBe(false);expect(JSON.stringify(exported)).toContain('cannot be trusted');expect(JSON.stringify(exported)).not.toContain('writable');
  expect(fs.existsSync(path.join(s.dir,'x.docx'))).toBe(false);
 });
 it('reports a deleted snapshot after open as unavailable',async()=>{
  const s=await setup();await fsp.unlink(snapshotPath(s.file));
  expect((await originalOf(s)).available).toBe(false);
 });
 it('never reports a different valid original as this session’s original (history recreated while open)',async()=>{
  const s=await setup();const other=Buffer.concat([s.original,Buffer.from('a different document')]);
  await fsp.rm(engine.getRepoDir(s.file),{recursive:true,force:true});
  const {xml,extras}=await initInputs(s.original);await engine.initAstraRepo(s.file,xml,extras,other);
  expect(await engine.getOriginalInfo(s.file)).toMatchObject({available:true});
  expect(await originalOf(s)).toMatchObject({available:false,reason:'capture-failed'});
  expect((await handleExportOriginalDocument(s.store,{sessionId:s.sessionId,outPath:path.join(s.dir,'y.docx')})).ok).toBe(false);
 });
 it('an ordinary reopen of an edited document stays eligible although its bytes differ from the original',async()=>{
  const s=await setup();await edit(s);await save(s);expect(sha(await fsp.readFile(s.file))).not.toBe(sha(s.original));
  await s.store.closeDocument(s.sessionId);
  const again=await LEAN6_HANDLERS.open_document(s.store,{path:s.file,validationTimeoutMs:0});if(!again.ok)throw new Error(JSON.stringify(again));
  const r=await GIT_HISTORY_HANDLERS.get_document_history(s.store,{sessionId:again.sessionId});
  expect(r.ok&&r.original).toMatchObject({available:true,sha256:sha(s.original)});
  const out=path.join(s.dir,'reopened-original.docx');
  expect((await handleExportOriginalDocument(s.store,{sessionId:again.sessionId,outPath:out})).ok).toBe(true);expect(sha(await fsp.readFile(out))).toBe(sha(s.original));
 });
 it('a losing opener with the same bytes reports initializing, then becomes eligible once the winner completes',async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-loser-'));dirs.push(dir);
  const file=path.join(dir,'race.docx');const bytes=engine.createBlankDocxBuffer('blank');await fsp.writeFile(file,bytes);
  const lock=await startCompetingInit(file);
  const store=createMcpSessionStore();const loser=await store.openBuffer(bytes,file);
  expect(loser.originalInfo).toMatchObject({available:false,reason:'initializing'});expect(loser.originalWarning).toContain('initialization');
  const sessionId=loser.sessionId;
  // Still running: nothing changes, nothing waits.
  const during=await GIT_HISTORY_HANDLERS.get_document_history(store,{sessionId});expect(during.ok&&during.original).toMatchObject({available:false,reason:'initializing'});
  expect((await handleExportOriginalDocument(store,{sessionId,outPath:path.join(dir,'early.docx')})).ok).toBe(false);
  // The winner finishes.
  await fsp.unlink(lock);const {xml,extras}=await initInputs(bytes);await engine.initAstraRepo(file,xml,extras,bytes);
  const after=await GIT_HISTORY_HANDLERS.get_document_history(store,{sessionId});expect(after.ok&&after.original).toMatchObject({available:true,sha256:sha(bytes)});
  const out=path.join(dir,'recovered.docx');expect((await handleExportOriginalDocument(store,{sessionId,outPath:out})).ok).toBe(true);expect(sha(await fsp.readFile(out))).toBe(sha(bytes));
 });
 it('a losing opener becomes eligible even when the winner has already saved edits',async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-loser-edit-'));dirs.push(dir);
  const file=path.join(dir,'race.docx');const bytes=engine.createBlankDocxBuffer('blank');await fsp.writeFile(file,bytes);
  const lock=await startCompetingInit(file);const store=createMcpSessionStore();const loser=await store.openBuffer(bytes,file);
  await fsp.unlink(lock);const {xml,extras}=await initInputs(bytes);await engine.initAstraRepo(file,xml,extras,bytes);
  await engine.commitChange(file,xml.replace('<w:body>','<w:body><!--winner edit-->'),'winner edit',extras);
  const r=await GIT_HISTORY_HANDLERS.get_document_history(store,{sessionId:loser.sessionId});
  expect(r.ok&&r.original).toMatchObject({available:true,sha256:sha(bytes)});
 });
 it('a losing opener with different bytes never receives the winner’s original',async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-loser-diff-'));dirs.push(dir);
  const file=path.join(dir,'race.docx');const mine=engine.createBlankDocxBuffer('blank');const theirs=Buffer.concat([mine,Buffer.from('different input')]);await fsp.writeFile(file,mine);
  const lock=await startCompetingInit(file);const store=createMcpSessionStore();const loser=await store.openBuffer(mine,file);
  await fsp.unlink(lock);const {xml,extras}=await initInputs(mine);await engine.initAstraRepo(file,xml,extras,theirs);
  const r=await GIT_HISTORY_HANDLERS.get_document_history(store,{sessionId:loser.sessionId});
  expect(r.ok&&r.original).toMatchObject({available:false,reason:'capture-failed'});
  expect((await handleExportOriginalDocument(store,{sessionId:loser.sessionId,outPath:path.join(dir,'no.docx')})).ok).toBe(false);
 });
 it('a winner killed after publishing the snapshot and before its first commit never makes a session eligible',async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-killed-'));dirs.push(dir);
  const file=path.join(dir,'race.docx');const bytes=engine.createBlankDocxBuffer('blank');await fsp.writeFile(file,bytes);
  await engine.ensureAstraGitRepo(file);await git.init({fs,dir:engine.getRepoDir(file)});
  expect((await engine.captureOriginal(file,bytes)).available).toBe(true); // snapshot published, no commit, no lock
  const store=createMcpSessionStore();const opened=await store.openBuffer(bytes,file);
  expect(opened.originalInfo).toMatchObject({available:false,reason:'initializing'});
  const r=await GIT_HISTORY_HANDLERS.get_document_history(store,{sessionId:opened.sessionId});expect(r.ok&&r.original).toMatchObject({available:false,reason:'initializing'});
  expect((await handleExportOriginalDocument(store,{sessionId:opened.sessionId,outPath:path.join(dir,'no.docx')})).ok).toBe(false);
 });
 it('refuses to save while an initialization lock is held, before touching the file, and saves once it is removed',async()=>{
  const s=await setup();await edit(s);const before=await fsp.readFile(s.file);
  const lock=path.join(engine.getRepoDir(s.file),'.astra-init.lock');await fsp.writeFile(lock,JSON.stringify({version:1,pid:999999,token:'old'}));
  const write=vi.spyOn(engine,'writeFileNoFollow');
  const r=await LEAN6_HANDLERS.save_document(s.store,{sessionId:s.sessionId});
  expect(r.ok).toBe(false);expect(JSON.stringify(r)).toContain('E_CONFLICT');expect(JSON.stringify(r)).toContain(lock);expect(JSON.stringify(r)).toContain('cannot tell');
  expect((await fsp.readFile(s.file)).equals(before)).toBe(true);expect(write).not.toHaveBeenCalled();
  await fsp.unlink(lock);await save(s);
 });
 it('does not commit an initial baseline when the package parts cannot be collected',async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-snapshot-fail-'));dirs.push(dir);
  const file=path.join(dir,'broken.docx');await fsp.writeFile(file,engine.createBlankDocxBuffer('blank'));
  vi.spyOn(engine.DocxParser.prototype,'capturePackageSnapshot').mockImplementation(()=>{throw new Error('injected snapshot collection failure');});
  const opened=await LEAN6_HANDLERS.open_document(createMcpSessionStore(),{path:file,validationTimeoutMs:0});
  // Opening still works, but history is not created and recovery is not offered.
  expect(opened.ok,JSON.stringify(opened)).toBe(true);
  if(opened.ok) expect(opened.original).toMatchObject({available:false});
  await expect(engine.getHistory(file)).rejects.toThrow(/no history exists/);
  expect(fs.existsSync(path.join(engine.getRepoDir(file),'.git','astra-original','original.snapshot'))).toBe(false);
  // A later save must not create the history from an empty package either.
  vi.restoreAllMocks();
  const store=createMcpSessionStore();const again=await LEAN6_HANDLERS.open_document(store,{path:file,validationTimeoutMs:0});
  expect(again.ok).toBe(true);if(!again.ok)return;
  vi.spyOn(engine.DocxParser.prototype,'capturePackageSnapshot').mockImplementation(()=>{throw new Error('injected snapshot collection failure');});
  const saved=await LEAN6_HANDLERS.save_document(store,{sessionId:again.sessionId});
  expect(saved.ok).toBe(false);
 });
 it('an opener that read a published original before the history was complete keeps the input-hash rule',async()=>{
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-original-publish-window-'));dirs.push(dir);
  const file=path.join(dir,'window.docx');const firstBytes=engine.createBlankDocxBuffer('blank');const otherBytes=Buffer.concat([firstBytes,Buffer.from('different input')]);
  const parser=new engine.DocxParser(firstBytes);await parser.getAST();
  const extras:Record<string,Buffer>={};for(const[name,buf]of parser.capturePackageSnapshot())if(engine.isTrackedPackageFile(name))extras[name]=buf;
  // State a winner leaves between publishing its original and committing: .git exists, original published, no commit.
  await engine.ensureAstraGitRepo(file);await git.init({fs,dir:engine.getRepoDir(file)});await engine.captureOriginal(file,firstBytes);
  // The winner finishes right after this opener read the original's status.
  const real=engine.getOriginalInfo;let first=true;
  vi.spyOn(engine,'getOriginalInfo').mockImplementation(async(...args)=>{const result=await real(...args);if(first){first=false;await engine.initAstraRepo(file,parser.readDocumentXml(),extras,firstBytes);}return result;});
  await fsp.writeFile(file,otherBytes);
  const store=createMcpSessionStore();const handle=await store.openDocument(file);
  const status=await resolveSessionOriginal(handle);
  expect(status.available).toBe(false);
  const exported=await handleExportOriginalDocument(store,{sessionId:handle.sessionId,outPath:path.join(dir,'export.docx')});
  expect(exported.ok).toBe(false);
 });
 it('a save whose package parts cannot be collected changes nothing: not the file, the history or the session',async()=>{
  const s=await setup();await edit(s);const h=s.store.get(s.sessionId);
  const diskBefore=await fsp.readFile(s.file);const historyBefore=(await engine.getHistory(s.file)).map(e=>e.oid);
  const marks={revision:h.revision,savedRevision:h.savedRevision,byteSize:h.byteSize,filePath:h.filePath,lastWritten:h.lastWrittenSha256};
  expect(h.revision).not.toBe(h.savedRevision);
  const snapshot=vi.spyOn(engine.DocxParser.prototype,'capturePackageSnapshot');
  snapshot.mockImplementation(()=>{throw new Error('injected snapshot collection failure');});
  const saved=await LEAN6_HANDLERS.save_document(s.store,{sessionId:s.sessionId});
  expect(saved.ok).toBe(false);
  expect((await fsp.readFile(s.file)).equals(diskBefore)).toBe(true);
  expect((await engine.getHistory(s.file)).map(e=>e.oid)).toEqual(historyBefore);
  expect({revision:h.revision,savedRevision:h.savedRevision,byteSize:h.byteSize,filePath:h.filePath,lastWritten:h.lastWrittenSha256}).toEqual(marks);
  // Save As to a new path: nothing is created either.
  const copy=path.join(s.dir,'copy.docx');
  const savedAs=await LEAN6_HANDLERS.save_document(s.store,{sessionId:s.sessionId,outPath:copy});
  expect(savedAs.ok).toBe(false);expect(fs.existsSync(copy)).toBe(false);expect(h.filePath).toBe(marks.filePath);
  // With the fault gone the same session saves normally.
  snapshot.mockRestore();
  const again=await LEAN6_HANDLERS.save_document(s.store,{sessionId:s.sessionId});expect(again.ok,JSON.stringify(again)).toBe(true);
  expect((await engine.getHistory(s.file)).length).toBe(historyBefore.length+1);
 });
});
