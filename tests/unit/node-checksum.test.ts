import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, calculateNodeChecksum, calculateNodeChecksumSync } from '@attentionform/astra-engine';

/**
 * The block checksum used `crypto.subtle.digest`, one separate job per block (about 125 ms for 3,000 blocks, paid on
 * every open and every refresh of the model). It now uses a synchronous SHA-256. The string must not change: a stored
 * or compared checksum from an older build has to equal the new one.
 */
async function viaSubtle(text: string): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('');
}

describe('block checksum', () => {
  const strings = [
    '', ' ', 'a', '<w:p/>', 'é', 'ü ß', '日本語', '😀 emoji', 'a\u0000b', 'tab\tnew\nline\r\n', '\ud800 lone high', 'lone low \udc00',
    '<w:p w14:paraId="0A1B2C3D"><w:r><w:t xml:space="preserve">Paragraph number 7 lorem ipsum</w:t></w:r></w:p>',
    'x'.repeat(100_000),
  ];

  it('is the same string as the one the digest of the old implementation gave, for odd inputs too', async () => {
    for (const text of strings) {
      expect(calculateNodeChecksumSync(text), JSON.stringify(text.slice(0, 20))).toBe(await viaSubtle(text));
    }
  });

  it('the async function returns the same value as the synchronous one', async () => {
    for (const text of strings) expect(await calculateNodeChecksum(text)).toBe(calculateNodeChecksumSync(text));
  });

  it('is 64 lower-case hex characters', () => {
    expect(calculateNodeChecksumSync('anything')).toMatch(/^[0-9a-f]{64}$/);
  });

  it('every block of every .docx in the repository carries the checksum the old implementation computed', async () => {
    const files = execFileSync('git', ['ls-files', '*.docx'], { encoding: 'utf8' }).split('\n').filter(Boolean);
    expect(files.length).toBeGreaterThan(20);
    let blocks = 0;
    for (const file of files) {
      let parser: DocxParser;
      let ast;
      try {
        parser = new DocxParser(readFileSync(file));
        ast = await parser.getAST();
      } catch {
        continue; // hostile fixtures that are refused on open
      }
      for (const node of ast.sections.flatMap(section => section.nodes)) {
        expect(node.checksum, `${file} ${node.id}`).toBe(await viaSubtle(new XMLSerializer().serializeToString(node.xml as Element)));
        blocks++;
      }
    }
    expect(blocks).toBeGreaterThan(200);
  });
});
