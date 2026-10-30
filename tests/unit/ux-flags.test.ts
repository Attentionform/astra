import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { DocumentManager } from '@attentionform/astra-engine';

/**
 * UX-C-flags (OX-09 / UX-01 P1-3): explicit opts:{untitled} replaces the
 * Drafts-substring heuristic as the source of truth; the heuristic stays as
 * a fallback safety net for direct-manager callers that pass no opts.
 */
describe('UX-C-flags: explicit untitled flag (OX-09)', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-ux-flags-'));
  });

  afterEach(() => {
    if (tempDir && fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  });

  it('explicit untitled:true wins over heuristic on a normal (titled) path', async () => {
    const docManager = new DocumentManager();
    const normalPath = path.join(tempDir, 'MyNormalDoc.docx');

    await docManager.createNewDocument('blank', normalPath, { untitled: true });

    expect(docManager.getActiveDocPath()).toBe(normalPath);
    expect(docManager.getIsUntitled()).toBe(true);
  });

  it('explicit untitled:false wins over Drafts heuristic on a silent-Drafts path', async () => {
    const docManager = new DocumentManager();
    const draftsPath = path.join(tempDir, 'Drafts', `ASTra_Document_${Date.now()}.docx`);

    await docManager.createNewDocument('blank', draftsPath, { untitled: false });

    expect(docManager.getActiveDocPath()).toBe(draftsPath);
    expect(docManager.getIsUntitled()).toBe(false);
  });

  it('fallback (no opts) still classifies the silent Drafts path as untitled', async () => {
    const docManager = new DocumentManager();
    const draftsPath = path.join(tempDir, 'Drafts', `ASTra_Document_${Date.now()}.docx`);

    await docManager.createNewDocument('blank', draftsPath);

    expect(docManager.getActiveDocPath()).toBe(draftsPath);
    expect(docManager.getIsUntitled()).toBe(true);
  });

  it('user-chosen Drafts-like path + explicit false stays titled', async () => {
    const docManager = new DocumentManager();
    // Same shape as the silent fallback (Drafts/ + ASTra_Document_ prefix) but
    // the caller explicitly marks it as a user-chosen (titled) location.
    const userDraftsLike = path.join(tempDir, 'Drafts', `ASTra_Document_userchosen.docx`);

    await docManager.createNewDocument('blank', userDraftsLike, { untitled: false });

    expect(docManager.getActiveDocPath()).toBe(userDraftsLike);
    expect(docManager.getIsUntitled()).toBe(false);
    expect(docManager.getIsDirty()).toBe(false);
  });

  it('fallback (no opts) classifies a normal explicit path as titled', async () => {
    const docManager = new DocumentManager();
    const normalPath = path.join(tempDir, 'ExplicitTarget.docx');

    const result = await docManager.createNewDocument('blank', normalPath);

    expect(result.astraDocPath).toBe(normalPath);
    expect(docManager.getIsUntitled()).toBe(false);
    expect(docManager.getIsDirty()).toBe(false);
  });
});
