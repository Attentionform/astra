import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { validateDocxBuffer } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS, GIT_HISTORY_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { useReproducibleOutput } from '../../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpDirs: string[] = [];
afterEach(async () => {
  while (tmpDirs.length > 0) {
    await fs.promises.rm(tmpDirs.pop() as string, { recursive: true, force: true }).catch(() => undefined);
  }
});

/** A package that holds no docProps/core.xml, as many Word-made files do. */
function withoutCoreProperties(): Buffer {
  const zip = new AdmZip(path.resolve(__dirname, '../../../../test-gallery/T3_proposal.docx'));
  return zip.toBuffer();
}

describe('revert after the core properties part appears', () => {
  it('restores a baseline that has no core properties part and keeps the package valid', async () => {
    const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'astra-revert-core-'));
    tmpDirs.push(dir);
    const file = path.join(dir, 'doc.docx');
    const source = withoutCoreProperties();
    expect(new AdmZip(source).getEntry('docProps/core.xml')).toBeNull();
    await fs.promises.writeFile(file, source);

    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: file });
    if (!opened.ok) throw new Error('open failed in test setup');
    const sessionId = opened.sessionId;

    const base = await GIT_HISTORY_HANDLERS.get_document_history(store, { sessionId });
    if (!base.ok) throw new Error('history failed in test setup');
    const baselineOid = base.history[0].oid;

    // Writing a title creates docProps/core.xml with its content type and package relationship.
    const meta = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'metadata', operation: 'setProperties', title: 'Probe title' }],
    });
    expect(meta.ok).toBe(true);
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId, message: 'with title' });
    expect(saved.ok).toBe(true);
    expect(new AdmZip(fs.readFileSync(file)).getEntry('docProps/core.xml')).not.toBeNull();

    const reverted = await GIT_HISTORY_HANDLERS.revert_document(store, { sessionId, commitOid: baselineOid, message: 'back to baseline' });
    expect(reverted.ok).toBe(true);

    const bytes = fs.readFileSync(file);
    const strict = await validateDocxBuffer(bytes);
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
    const types = new AdmZip(bytes).getEntry('[Content_Types].xml')?.getData().toString('utf8') ?? '';
    const hasPart = new AdmZip(bytes).getEntry('docProps/core.xml') !== null;
    expect(hasPart ? types.includes('/docProps/core.xml') : true).toBe(true);
  });
});
