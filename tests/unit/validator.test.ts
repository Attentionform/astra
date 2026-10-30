import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  CitationValidationError,
  DocxParser,
  EquationValidationError,
  FormattingValidationError,
  NamespaceValidationError,
  RunCountValidationError,
  TrackedChangeValidationError,
  validateNode,
} from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';

describe('Part 3: The 6-Layer Validation Engine', () => {
  const validOriginal = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
    <w:r><w:t>A</w:t></w:r>
    <m:oMath><m:r><m:t>E=mc2</m:t></m:r></m:oMath>
    <w:fldSimple w:instr="CITATION"><w:r><w:t>[1]</w:t></w:r></w:fldSimple>
    <w:ins><w:r><w:t>Tracked</w:t></w:r></w:ins>
  </w:p>`;

  it('should return true for a completely happy path valid edit', () => {
    const validEdit = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
    <w:r><w:t>A revised sentence here.</w:t></w:r>
    <m:oMath><m:r><m:t>E=mc2</m:t></m:r></m:oMath>
    <w:fldSimple w:instr="CITATION"><w:r><w:t>[1]</w:t></w:r></w:fldSimple>
    <w:ins><w:r><w:t>Tracked</w:t></w:r></w:ins>
  </w:p>`;
    expect(validateNode(validOriginal, validEdit)).toBe(true);
  });

  it('Layer 1: should reject if <m:oMath> is deleted', () => {
    const edit = validOriginal.replace('<m:oMath><m:r><m:t>E=mc2</m:t></m:r></m:oMath>', '');
    expect(() => validateNode(validOriginal, edit)).toThrow(EquationValidationError);
  });

  it('Layer 2: should reject if <w:fldSimple> is deleted', () => {
    const edit = validOriginal.replace('<w:fldSimple w:instr="CITATION"><w:r><w:t>[1]</w:t></w:r></w:fldSimple>', '');
    expect(() => validateNode(validOriginal, edit)).toThrow(CitationValidationError);
  });

  it('Layer 3: should reject if <w:pPr> formatting is altered', () => {
    const edit = validOriginal.replace('Heading1', 'Normal');
    expect(() => validateNode(validOriginal, edit)).toThrow(FormattingValidationError);
  });

  it('Layer 4: should reject if run count changes drastically', () => {
    // Original has 4 w:r tags. Adding 4 more to break the max threshold of 3 difference.
    const edit = validOriginal.replace('</w:p>', '<w:r><w:t>1</w:t></w:r><w:r><w:t>2</w:t></w:r><w:r><w:t>3</w:t></w:r><w:r><w:t>4</w:t></w:r></w:p>');
    expect(() => validateNode(validOriginal, edit, { maxRunCountDelta: 3 })).toThrow(RunCountValidationError);
  });

  it('Layer 5: should reject if new namespace is introduced', () => {
    const edit = validOriginal.replace('<w:p ', '<w:p xmlns:custom="http://evil.com" ');
    expect(() => validateNode(validOriginal, edit)).toThrow(NamespaceValidationError);
  });

  it('Layer 6: should reject if <w:ins> tracked change is deleted', () => {
    const edit = validOriginal.replace('<w:ins><w:r><w:t>Tracked</w:t></w:r></w:ins>', '');
    expect(() => validateNode(validOriginal, edit)).toThrow(TrackedChangeValidationError);
  });
});

describe('validate command on damaged containers', () => {
  const tmpFiles: string[] = [];

  function writeTemp(suffix: string, data: Buffer): string {
    const file = path.join(os.tmpdir(), `astra-validate-${process.pid}-${Date.now()}-${tmpFiles.length}-${suffix}.docx`);
    tmpFiles.push(file);
    fs.writeFileSync(file, data);
    return file;
  }

  function stdoutText(): string {
    return vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
  }

  function stderrText(): string {
    return vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
  }

  beforeEach(() => {
    vi.spyOn(console, 'log').mockImplementation(() => undefined);
    vi.spyOn(console, 'error').mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    while (tmpFiles.length > 0) {
      const f = tmpFiles.pop() as string;
      try {
        fs.unlinkSync(f);
      } catch {
        /* best-effort */
      }
    }
  });

  it('leaves a well-formed document valid', async () => {
    const file = writeTemp('wellformed', DocxParser.createNewDocument('blank').rezip());
    await expect(run(['validate', file])).resolves.toBe(0);
    expect(stdoutText()).toContain(`VALID ${file}`);
  });

  it('names a byte-damaged container without leaking allocation internals', async () => {
    const base = DocxParser.createNewDocument('blank').rezip();
    let damaged = '';
    for (const offset of [150, 500, 1000, 1200, 1500]) {
      if (offset >= base.length) continue;
      const probe = Buffer.from(base);
      probe[offset] = 0xff;
      const candidate = writeTemp(`damaged-${offset}`, probe);
      vi.mocked(console.error).mockClear();
      const code = await run(['validate', candidate]);
      if (code === 1 && stderrText().includes('(unreadable container)')) {
        damaged = candidate;
        break;
      }
    }
    expect(damaged).not.toBe('');
    vi.mocked(console.log).mockClear();
    vi.mocked(console.error).mockClear();
    await expect(run(['validate', damaged])).resolves.toBe(1);
    const err = stderrText();
    expect(err).toContain('[E_VALIDATION]');
    expect(err).toContain('not a valid .docx package (unreadable container)');
    // Error-hygiene contract (decided, not drift): absolute input paths
    // redact to "<path>" in the one-liner so machine layout (tmp dir, user
    // home) never leaves the box, while relative names echo literally so
    // the file is still named. This probe feeds an absolute tmp path, so
    // it pins the redacted shape; the relative-echo half is pinned by the
    // CLI exit-contract suite.
    expect(err).toContain('"<path>" is not a valid .docx package (unreadable container)');
    expect(err).not.toContain(damaged);
    expect(err).not.toContain(path.basename(damaged));
    expect(err).not.toContain('Buffer');
    expect(err).toContain('restore the file from source');
    expect(err).not.toContain('astra inspect');
  });
});
