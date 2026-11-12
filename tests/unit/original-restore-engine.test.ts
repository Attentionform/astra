import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fsp from 'node:fs/promises';
import * as fs from 'node:fs';
import git from 'isomorphic-git';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { DocxParser, createBlankDocxBuffer, initAstraRepo, commitChange, getHistory, isTrackedPackageFile, restoreOriginal } from '@attentionform/astra-engine';
const fault = vi.hoisted(() => ({ commit: false, afterCommit: false, stageThenFail: false, failRollback: false, replaceCalls: 0 }));
vi.mock('../../packages/engine/src/git', async (importOriginal) => {
 const actual = await importOriginal<typeof import('../../packages/engine/src/git')>();
 return { ...actual, commitChange: async (...args: Parameters<typeof actual.commitChange>) => {
  if (fault.commit) throw new Error('injected commit failure');
  if (fault.stageThenFail) { const repo=actual.getRepoDir(args[0]); await fsp.writeFile(path.join(repo,'injected-stage'),'unexpected'); await git.add({fs,dir:repo,filepath:'injected-stage'}); throw new Error('injected staged failure'); }
  const result = await actual.commitChange(...args);
  if (fault.afterCommit) throw new Error('injected failure after commit');
  return result;
 } };
});
vi.mock('../../packages/engine/src/original', async (importOriginal) => {
 const actual = await importOriginal<typeof import('../../packages/engine/src/original')>();
 return { ...actual, replaceFileAtomic: async (...args: Parameters<typeof actual.replaceFileAtomic>) => {
  fault.replaceCalls++;
  if (fault.failRollback && fault.replaceCalls === 2) throw new Error('injected rollback failure');
  return actual.replaceFileAtomic(...args);
 } };
});
let dir: string;
const hash=(b:Buffer)=>createHash('sha256').update(b).digest('hex');
beforeEach(async()=>{dir=await fsp.mkdtemp(path.join(os.tmpdir(),'astra-engine-restore-'));});
afterEach(async()=>{fault.commit=false;fault.afterCommit=false;fault.stageThenFail=false;fault.failRollback=false;fault.replaceCalls=0;await fsp.rm(dir,{recursive:true,force:true});});
async function setup(){
 const file=path.join(dir,'report.docx');const original=createBlankDocxBuffer('blank');await fsp.writeFile(file,original);
 const initial=new DocxParser(original);await initial.getAST();const first=initial.capturePackageSnapshot();const firstExtras=Object.fromEntries([...first].filter(([n])=>isTrackedPackageFile(n)));
 await initAstraRepo(file,initial.readDocumentXml(),firstExtras,original);
 const xml=initial.readDocumentXml().replace('</w:body>','<w:p><w:r><w:t>later change</w:t></w:r></w:p></w:body>');initial.reloadDocumentXml(xml);const edited=initial.rezip();await fsp.writeFile(file,edited);
 const parser=new DocxParser(edited);await parser.getAST();const extras=Object.fromEntries([...parser.capturePackageSnapshot()].filter(([n])=>isTrackedPackageFile(n)));
 await commitChange(file,parser.readDocumentXml(),'edit',extras);
 return {file,original,edited};
}
describe('engine exact-original restore',()=>{
 it('restores the original file byte-for-byte and appends history',async()=>{
  const s=await setup();const before=(await getHistory(s.file)).length;const r=await restoreOriginal(s.file,{expectedSha256:hash(s.edited)});
  expect(hash(await fsp.readFile(s.file))).toBe(hash(s.original));expect(r.sha256).toBe(hash(s.original));expect((await getHistory(s.file)).length).toBe(before+1);
 });
 it('refuses a changed file without touching its bytes',async()=>{
  const s=await setup();const changed=Buffer.from('external');await fsp.writeFile(s.file,changed);
  await expect(restoreOriginal(s.file,{expectedSha256:hash(s.edited)})).rejects.toMatchObject({code:'E_CONFLICT'});expect(await fsp.readFile(s.file)).toEqual(changed);
 });
 it('rolls back disk and preserves history when the commit fails',async()=>{
  const s=await setup();const hist=(await getHistory(s.file)).map(x=>x.oid);fault.commit=true;
  await expect(restoreOriginal(s.file)).rejects.toThrow('injected commit failure');
  expect(await fsp.readFile(s.file)).toEqual(s.edited);expect((await getHistory(s.file)).map(x=>x.oid)).toEqual(hist);
  expect((await fsp.readdir(dir)).filter(x=>x.includes('astra-restore-backup'))).toEqual([]);
 });
 it('reports that history changed if a commit fails after advancing head',async()=>{
  const s=await setup();const hist=(await getHistory(s.file)).map(x=>x.oid);fault.afterCommit=true;
  await expect(restoreOriginal(s.file)).rejects.toThrow('History changed despite failure');
  expect(await fsp.readFile(s.file)).toEqual(s.edited);expect((await getHistory(s.file)).length).toBe(hist.length+1);
 });
 it('detects a changed Git index even if the branch head did not advance',async()=>{
  const s=await setup();const hist=(await getHistory(s.file)).map(x=>x.oid);fault.stageThenFail=true;
  await expect(restoreOriginal(s.file)).rejects.toThrow('History changed despite failure');
  expect(await fsp.readFile(s.file)).toEqual(s.edited);expect((await getHistory(s.file)).map(x=>x.oid)).toEqual(hist);
 });
 it('retains and reports the recovery backup when rollback fails',async()=>{
  const s=await setup();fault.commit=true;fault.failRollback=true;
  let caught: unknown;try{await restoreOriginal(s.file);}catch(e){caught=e;}
  expect(caught).toMatchObject({code:'E_IO',backupPath:expect.any(String)});
  const backup=(caught as {backupPath:string}).backupPath;expect(await fsp.readFile(backup)).toEqual(s.edited);
  expect(await fsp.readFile(s.file)).toEqual(s.original);
 });
 it('refuses a corrupt original before changing the working file',async()=>{
  const s=await setup();const info=await import('@attentionform/astra-engine');const file=path.join(info.getRepoDir(s.file),'.git','astra-original',`original-${hash(s.original)}.docx`);await fsp.writeFile(file,'bad');
  await expect(restoreOriginal(s.file)).rejects.toMatchObject({code:'E_STATE'});expect(await fsp.readFile(s.file)).toEqual(s.edited);
 });
});
