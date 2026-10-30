import { describe, it, expect } from 'vitest';
import { LEAN6_DESCRIPTIONS } from '../src/lib/mcp/descriptions';

describe('mcp execute_batch description discoverability', () => {
  it('documents section, page-setup, headerFooter, and image EMU operations', () => {
    const desc = LEAN6_DESCRIPTIONS.execute_batch;
    expect(desc).toContain('setColumns');
    expect(desc).toContain('pageSetup');
    expect(desc).toContain('sectionIndex');
    expect(desc).toContain('{PAGE}');
    expect(desc).toContain('EMU');
  });
});
