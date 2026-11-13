import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Every doubt ends in a refusal: an injected verifier failure and an injected validator failure each stop the
 * export before anything is rendered.
 */
const control = vi.hoisted(() => ({ verifier: 'real' as 'real' | 'false' | 'throw', validator: 'real' as 'real' | 'throw' }));

vi.mock('../../packages/engine/src/ooxml/normalize', async (importActual) => {
  const actual = await importActual<typeof import('../../packages/engine/src/ooxml/normalize')>();
  return {
    ...actual,
    verifyOrderOnlyRepair: (...args: Parameters<typeof actual.verifyOrderOnlyRepair>) => {
      if (control.verifier === 'false') return false;
      if (control.verifier === 'throw') throw new Error('verifier exploded');
      return actual.verifyOrderOnlyRepair(...args);
    },
  };
});

vi.mock('../../packages/engine/src/ooxml/validator-gate', async (importActual) => {
  const actual = await importActual<typeof import('../../packages/engine/src/ooxml/validator-gate')>();
  let calls = 0;
  return {
    ...actual,
    validateDocxBuffer: async (...args: Parameters<typeof actual.validateDocxBuffer>) => {
      calls += 1;
      // The first call validates the original; the injected failure hits the validation of the copy.
      if (control.validator === 'throw' && calls % 2 === 0) throw new Error('validator unavailable');
      return actual.validateDocxBuffer(...args);
    },
  };
});

import { prepareDocxForPdfExport } from '../../packages/engine/src/export/pdf-source';

const FIXTURE = path.resolve(__dirname, '../fixtures/word-order.docx');

describe('prepareDocxForPdfExport under injected failures', () => {
  beforeEach(() => {
    control.verifier = 'real';
    control.validator = 'real';
  });

  it('control: with nothing injected the fixture is repaired', async () => {
    const prepared = await prepareDocxForPdfExport(fs.readFileSync(FIXTURE));
    expect(prepared.repaired).toBe(true);
  });

  it('a verifier that says no refuses the export', async () => {
    control.verifier = 'false';
    await expect(prepareDocxForPdfExport(fs.readFileSync(FIXTURE))).rejects.toMatchObject({ code: 'E_VALIDATION' });
  });

  it('a verifier that throws refuses the export', async () => {
    control.verifier = 'throw';
    const error = (await prepareDocxForPdfExport(fs.readFileSync(FIXTURE)).catch((e: unknown) => e)) as { code?: string; cause?: unknown };
    expect(error.code).toBe('E_VALIDATION');
    expect((error.cause as Error).message).toContain('verifier exploded');
  });

  it('a validator that fails on the copy stops the export with the validator error', async () => {
    control.validator = 'throw';
    await expect(prepareDocxForPdfExport(fs.readFileSync(FIXTURE))).rejects.toThrow(/validator unavailable/);
  });
});
