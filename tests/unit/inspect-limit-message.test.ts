import { describe, expect, it } from 'vitest';
import { InspectDocumentSchema } from '../../packages/mcp-server/src/lib/mcp/schemas';

describe('inspect_document paging limits tell the caller how to page', () => {
  it('limit above 200 says to use offset and nextOffset', () => {
    const res = InspectDocumentSchema.safeParse({ sessionId: 's', limit: 500 });
    expect(res.success).toBe(false);
    const messages = res.success ? [] : res.error.issues.map((i) => i.message);
    expect(messages).toContain('limit is at most 200 per call; use offset and nextOffset to read more');
  });

  it('reviewLimit above 200 says to use reviewOffset', () => {
    const res = InspectDocumentSchema.safeParse({ sessionId: 's', reviewLimit: 500 });
    const messages = res.success ? [] : res.error.issues.map((i) => i.message);
    expect(messages).toContain('reviewLimit is at most 200 per call; use reviewOffset to read more');
  });

  it('200 is still accepted', () => {
    expect(InspectDocumentSchema.safeParse({ sessionId: 's', limit: 200 }).success).toBe(true);
  });
});
