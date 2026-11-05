import { describe, it, expect, vi } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import * as engine from '@attentionform/astra-engine';
import { cmdValidate } from '../src/cli/validate-cmd';

// Unlike a subprocess assertion, this observes the expensive entry points directly.
describe('validate advice does not run normalization', () => {
  it('returns the invalid verdict without invoking forecast or repair', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-val-work-'));
    const zip = new engine.AdmZip(engine.DocxParser.createNewDocument('blank').rezip());
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace('</w:body>', '<w:p><w:unknown/></w:p></w:body>')));
    const file = path.join(dir, 'invalid.docx');
    fs.writeFileSync(file, zip.toBuffer());
    const forecast = vi.spyOn(engine, 'forecastNormalize');
    const repair = vi.spyOn(engine, 'repairDocxBuffer');
    const output = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    try {
      expect(await cmdValidate({ json: true, verbose: false }, [file])).toBe(1);
      expect(forecast).not.toHaveBeenCalled();
      expect(repair).not.toHaveBeenCalled();
      const data = JSON.parse(output.mock.calls[0][0]);
      expect(data.normalizeHint).toContain('attempt supported repairs');
    } finally {
      vi.restoreAllMocks();
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});
