import { describe, it, expect, vi } from 'vitest';

vi.mock('../../backend/git', () => ({
  undo: vi.fn().mockRejectedValue(new Error('Cannot undo: already at the initial state.')),
  initAstraRepo: vi.fn(),
  commitChange: vi.fn(),
  getHistory: vi.fn()
}));

vi.mock('../../packages/engine/src/git', () => ({
  undo: vi.fn().mockRejectedValue(new Error('Cannot undo: already at the initial state.')),
  initAstraRepo: vi.fn(),
  commitChange: vi.fn(),
  getHistory: vi.fn()
}));

import { DocumentManager } from '@attentionform/astra-engine';

describe('Sprint 2: Stream Cancellation & Graceful Undo', () => {
  it('document-manager.ts: undoLastAction returns atInitialState: true when only initial commit exists', async () => {
    const docManager = new DocumentManager();
    const mockParser = {
      readDocumentXml: () => '<w:document/>',
      reloadDocumentXml: vi.fn(),
      rezip: () => Buffer.from('mock-zip'),
      getStructuralSummary: () => ({ nodes: [], sections: [] })
    } as any;

    (docManager as any).parser = mockParser;
    (docManager as any).activeDocPath = 'mock-test-path.docx';

    const result = await docManager.undoLastAction();
    expect(result.success).toBe(false);
    expect(result.atInitialState).toBe(true);
  });
});
