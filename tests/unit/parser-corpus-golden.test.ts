/**
 * Parser corpus golden (parser/ decomposition witness).
 *
 * Gap filled: no suite pinned byte-identical parse outputs across fixture
 * files. Pins sha256 of `getStructuralSummary()` canonical JSON + `rezip()`
 * bytes for the stable test-gallery corpus. Any behavior change in the
 * parser split (or a future refactor) breaks this loudly.
 *
 * Hashes verified HEAD-vs-split identical (27/27 tracked .docx, summary +
 * rezip) and rerun-deterministic. Runs under default reproducible mode only.
 */
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { describe, expect, it } from 'vitest';
import { DocxParser } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

const GOLDEN: Record<string, { summary: string; rezip: string }> = {
    "test-gallery/T1_coffee_onepager.docx": { summary: "0c6a95a3f4a41185b276094e858f3659b717d0d99645dc3d780fda0db47d3c6e", rezip: "1f75a9d2494b8fcbcdf739f52cd5b03d70a1cb5bf86ff8ac91d87b5c663c7451" },
    "test-gallery/T2_edit_target.docx": { summary: "e28e4bd458b74363745d4aa1d940c4e4bd118c6cb3604c1ba705729777e33485", rezip: "4d1ab8903092c44609d2df71b994d95ae828c86f9217f478699b9d1039550c15" },
    "test-gallery/T3_proposal.docx": { summary: "28ac8edcca529108bf61e4206153fe895c9856adba9a634b60597356032a29be", rezip: "a62dabb932fcca861c2b24809bdda2045ea8b7fed14d4bdeffe102a1d1476b01" },
    "test-gallery/T4_budget.docx": { summary: "df13adbb1ccd5290708f00a4731256c4f5b0e6017f960c26d355a72931fd738e", rezip: "1005879f11c8b8bba945b26007c8266b81c7b60fdb7a9645eea87b02433bfc3f" },
    "test-gallery/T5_image_doc.docx": { summary: "6082e96492425947ef11f8eba807e3ba2cfcad6e7bf1c0ab290fb41dc42a7a9e", rezip: "4ddb552344327f07a7581c9e15791d302ab1b38aa8c6ec5382962b64a8f06a7e" },
    "test-gallery/T6_honesty.docx": { summary: "3f89ce01141136a57833c946d082e7371c64bf81fafd9fe83f48a01382e33c25", rezip: "bbbe3695865b253f1367bb5e10567313f0a0058a4e4e8b7f912063c3af40777c" },
    "test-gallery/T7_visual_probe.docx": { summary: "fa0f5da780c8712eaf98fcbb16ac3a4654d7d750919a760298685531d6d541c5", rezip: "c1f122496fe60ed86c7cc636a589801134e3475270a039c979e3a17767a9927c" },
    "test-gallery/T8_footnotes_endnotes.docx": { summary: "c01cb9644c18b13e340068f7f2fef49bbe0ca6a9c3b24519bc4397d0a7561752", rezip: "0871bb3f20c8831c020626ce2a7693515273c877aaed8aa558be808025874463" },
    "test-gallery/T9_postmigration.docx": { summary: "7ff18d62b7173e2e94e6212dca7483fad10e97ade529d74c3c432a9fabf39a1e", rezip: "c6c5f25a3913468546696008516ef7b9cb307cf8586f776f51d5b21d3d446369" },
};

const REPO_ROOT = path.resolve(__dirname, '../..');
const OPT_OUT =
  ['1', 'true', 'yes', 'on'].includes(String(process.env.ASTRA_NO_REPRODUCIBLE || '').trim().toLowerCase()) ||
  ['0', 'false', 'no', 'off'].includes(String(process.env.ASTRA_REPRODUCIBLE || '').trim().toLowerCase());

useReproducibleOutput();

describe.runIf(!OPT_OUT)('parser corpus golden', () => {
  for (const [rel, want] of Object.entries(GOLDEN)) {
    it(`byte-identical summary+rezip: ${rel}`, () => {
      const buf = fs.readFileSync(path.join(REPO_ROOT, rel));
      const parser = new DocxParser(buf);
      const summarySha = crypto.createHash('sha256').update(JSON.stringify(parser.getStructuralSummary())).digest('hex');
      const rezipSha = crypto.createHash('sha256').update(parser.rezip()).digest('hex');
      expect(summarySha).toBe(want.summary);
      expect(rezipSha).toBe(want.rezip);
    });
  }
});
