import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleInspectDocument } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function xmlEscape(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function normalizeStyle(s: string): string {
  return s.toLowerCase().replace(/\s+/g, '');
}

interface TemplateCase {
  name: string;
  title: string;
  body: string;
  /** Exact styleId expected from the palette (verbatim, no hallucination). */
  expected: string;
}

// H1 (reworked): discriminating cases across 5 palette entries.
// Corporate-cue phrases ("Key finding", "Board notice", "Takeaway", ...) are
// DELIBERATELY planted in distractor bodies (stat/disclaimer/quote) so a
// naive cue-regex or always-corporate picker FAILS. The correct blind picker
// must rank specific signals (disclaimer > quote > stat) above the generic
// corporate cue, and must return Normal when no signal fires (negatives).
const TEMPLATE_CASES: TemplateCase[] = [
  // Corporate Callout positives (callout cue, NO stat/quote/disclaimer signals).
  { name: 'corp-expansion', title: 'Expansion Review', body: 'Key finding: enterprise adoption expanded steadily across all segments this quarter.', expected: 'CorporateCallout' },
  { name: 'corp-risk', title: 'Risk Review', body: 'Strategic insight: supply concentration remains the top residual risk this fiscal year.', expected: 'CorporateCallout' },
  { name: 'corp-launch', title: 'Launch Readiness', body: 'Takeaway: launch readiness is green across documentation, engineering, and support.', expected: 'CorporateCallout' },
  // Stat Callout distractors (corporate cue + hard numeric signals -> Stat wins).
  { name: 'stat-retention', title: 'Retention Report', body: 'Key finding: net revenue retention reached 128% on expansion alone.', expected: 'StatCallout' },
  { name: 'stat-trial', title: 'Trial Readout', body: 'Takeaway: the treatment arm met its primary endpoint with 94% efficacy.', expected: 'StatCallout' },
  { name: 'stat-deals', title: 'Deal Velocity', body: 'Strategic insight: multi-threaded deals close 2.4x faster than single-threaded ones.', expected: 'StatCallout' },
  // Disclaimer distractors (corporate cue + fine-print signals -> Disclaimer wins).
  { name: 'disclaimer-privilege', title: 'Legal Notice', body: 'Board notice: this analysis is for informational purposes only and does not constitute legal advice.', expected: 'Disclaimer' },
  { name: 'disclaimer-forward', title: 'Forward Statements', body: 'Highlight: forward-looking statements are subject to change without notice and carry no warranty.', expected: 'Disclaimer' },
  { name: 'disclaimer-creds', title: 'Credential Handling', body: 'Board notice: credentials described here are for authorized testing at your own risk only.', expected: 'Disclaimer' },
  // Quote distractors (corporate cue + quotation marks -> Quote wins).
  { name: 'quote-ceo', title: 'CEO Remarks', body: 'Takeaway: as the CEO noted, \u201Cwe will double down on onboarding\u201D in her remarks.', expected: 'Quote' },
  { name: 'quote-customer', title: 'Customer Voice', body: 'Highlight: one customer said, \u201Conboarding took ten minutes, not ten days\u201D.', expected: 'Quote' },
  { name: 'quote-board', title: 'Board Minutes', body: 'Conclusion: the chair stated, \u201Cthe plan delivers without headcount growth\u201D.', expected: 'Quote' },
  // Negatives: plain prose (no cue, no stat, no quote, no disclaimer -> Normal).
  { name: 'normal-roadmap', title: 'Weekly Notes', body: 'The team met on Tuesday to discuss the roadmap and assign owners.', expected: 'Normal' },
  { name: 'normal-recap', title: 'Meeting Recap', body: 'Attendees reviewed the calendar and agreed to reconvene next week.', expected: 'Normal' },
  { name: 'normal-style', title: 'Style Guide', body: 'Use short paragraphs and plain headings for internal drafts.', expected: 'Normal' },
];

function buildTemplateDocx(c: TemplateCase): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
      'utf8'
    )
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      'utf8'
    )
  );
  const docXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}">` +
    `<w:body>` +
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${xmlEscape(c.title)}</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>${xmlEscape(c.body)}</w:t></w:r></w:p>` +
    `<w:sectPr/>` +
    `</w:body></w:document>`;
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  const stylesXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="${W_NS}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:color w:val="0F172A"/><w:sz w:val="48"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="CorporateCallout"><w:name w:val="Corporate Callout"/><w:rPr><w:b/><w:color w:val="3730A3"/><w:sz w:val="24"/></w:rPr><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="EEF2FF"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="StatCallout"><w:name w:val="Stat Callout"/><w:rPr><w:b/><w:color w:val="065F46"/><w:sz w:val="36"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Disclaimer"><w:name w:val="Disclaimer"/><w:rPr><w:color w:val="64748B"/><w:sz w:val="18"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:rPr><w:i/><w:color w:val="475569"/><w:sz w:val="22"/></w:rPr></w:style>` +
    `</w:styles>`;
  zip.addFile('word/styles.xml', Buffer.from(stylesXml, 'utf8'));
  return zip.toBuffer();
}

/**
 * Blind agent: picks a style using ONLY standard TOON text + palette.
 * Specific signals outrank the generic corporate cue (disclaimer > quote >
 * stat > corporate cue > Normal). Takes a `tools` bag containing a
 * visual_review stub that THROWS — the agent must succeed without calling it.
 */
function blindPickStyle(
  toonText: string,
  availableStyles: string[],
  tools: { visual_review: () => never }
): string {
  void tools;
  const lower = toonText.toLowerCase();
  const byNorm = new Map(availableStyles.map((s) => [normalizeStyle(s), s]));
  const pick = (norm: string): string | undefined => byNorm.get(norm);
  // 1. Disclaimer fine-print wins even when a corporate cue is present.
  if (/does not constitute|informational purposes|no warranty|subject to change without notice|at your own risk/.test(lower)) {
    const d = pick('disclaimer');
    if (d) return d;
  }
  // 2. Quotation marks win (quoted speech is a Quote, not a callout).
  if (toonText.includes('\u201C') || toonText.includes('\u201D') || toonText.includes('"')) {
    const q = pick('quote');
    if (q) return q;
  }
  // 3. Hard numeric signals win (%, $, Nx multipliers, bps).
  if (/[%$]|\d+\.\d+x|\bbps\b/.test(toonText)) {
    const s = pick('statcallout');
    if (s) return s;
  }
  // 4. Generic callout cue -> Corporate Callout.
  if (/key finding|strategic insight|board notice|takeaway|conclusion|highlight|recommendation/.test(lower)) {
    const c = pick('corporatecallout');
    if (c) return c;
  }
  // 5. Negatives: plain prose -> Normal (never Corporate).
  const n = pick('normal');
  if (n) return n;
  return availableStyles[0] ?? 'Normal';
}

describe('TOON-blind style selection (discriminating, real inspect path, no visual_review)', () => {
  it('covers all five palette entries with negatives', () => {
    expect(TEMPLATE_CASES.length).toBe(15);
    const kinds = new Set(TEMPLATE_CASES.map((c) => normalizeStyle(c.expected)));
    expect(kinds).toEqual(new Set(['corporatecallout', 'statcallout', 'disclaimer', 'quote', 'normal']));
    // Distractors exist: a majority of cases are NOT Corporate (naive picker fails).
    const nonCorporate = TEMPLATE_CASES.filter((c) => normalizeStyle(c.expected) !== 'corporatecallout');
    expect(nonCorporate.length).toBe(12);
  });

  it('inspect_document palette passthrough carries every distractor style verbatim', async () => {
    const store = createMcpSessionStore();
    const buf = buildTemplateDocx(TEMPLATE_CASES[0]);
    const opened = await store.openBuffer(buf, 'memory.docx');
    const res = await handleInspectDocument(store, { sessionId: opened.sessionId, fidelity: 'standard' });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.fidelity).toBe('standard');
    const norms = res.palette.availableStyles.map(normalizeStyle);
    for (const want of ['corporatecallout', 'statcallout', 'disclaimer', 'quote', 'normal']) {
      expect(norms).toContain(want);
    }
    // Brand colors passthrough (hex without #, from styles.xml).
    for (const hex of ['3730A3', '065F46', '64748B', '475569']) {
      expect(res.palette.brandColors.map((c) => c.toUpperCase())).toContain(hex);
    }
    expect(res.palette.primaryFont.length).toBeGreaterThan(0);
    await store.closeDocument(opened.sessionId).catch(() => false);
  });

  it('blind agent resolves every case from standard TOON without visual_review', async () => {
    // Wired blindness proof: the stub THROWS if called. It is PASSED to the
    // picker on every case, so a vision-dependent picker fails loudly.
    let visualReviewCalls = 0;
    const tools = {
      visual_review: (): never => {
        visualReviewCalls += 1;
        throw new Error('visual_review must not be called in the blind harness');
      },
    };

    const store = createMcpSessionStore();
    let correct = 0;
    for (const c of TEMPLATE_CASES) {
      const buf = buildTemplateDocx(c);
      // Real inspect path: openBuffer -> inspect_document (TOON + palette).
      const opened = await store.openBuffer(buf, 'memory.docx');
      const res = await handleInspectDocument(store, { sessionId: opened.sessionId, fidelity: 'standard' });
      expect(res.ok).toBe(true);
      if (!res.ok) throw new Error(`inspect failed for ${c.name}`);
      expect(res.fidelity).toBe('standard');
      expect(res.toonText.split('\n')[0]).toMatch(/fidelity: lossy.*standard/);
      expect(res.toonText).toContain(c.body.slice(0, 24));
      expect(res.rowCount).toBeGreaterThan(0);

      // Palette passthrough: expected style offered verbatim (no hallucination).
      expect(res.palette.availableStyles.map(normalizeStyle)).toContain(normalizeStyle(c.expected));

      const picked = blindPickStyle(res.toonText, res.palette.availableStyles, tools);
      // Picked style must come from the palette verbatim.
      expect(res.palette.availableStyles).toContain(picked);
      expect(normalizeStyle(picked)).toBe(normalizeStyle(c.expected));
      correct += 1;
      // Best-effort session cleanup (never masks a pick failure above).
      await store.closeDocument(opened.sessionId).catch(() => false);
    }

    expect(correct).toBe(TEMPLATE_CASES.length);
    expect(visualReviewCalls).toBe(0);
  });

  it('is discriminating: always-corporate and cue-only pickers fail', () => {
    // Prove the suite is non-vacuous: degenerate strategies score < 100%.
    const alwaysCorporate = TEMPLATE_CASES.filter((c) => normalizeStyle(c.expected) === 'corporatecallout').length;
    expect(alwaysCorporate).toBe(3);
    expect(alwaysCorporate).toBeLessThan(TEMPLATE_CASES.length);

    const cueOnly = (body: string): string => {
      const lower = body.toLowerCase();
      if (/key finding|strategic insight|board notice|takeaway|conclusion|highlight|recommendation/.test(lower)) {
        return 'CorporateCallout';
      }
      return 'Normal';
    };
    const cueScore = TEMPLATE_CASES.filter((c) => normalizeStyle(cueOnly(c.body)) === normalizeStyle(c.expected)).length;
    // Cue-only gets the 3 corporate + 3 negatives right, misses all 9 distractors.
    expect(cueScore).toBe(6);
    expect(cueScore).toBeLessThan(TEMPLATE_CASES.length);
  });

  it('inspect path is code-blind: no visual-review import in the MCP core', () => {
    // Static blindness: the real inspect path must not depend on screenshots.
    // If a future change wires visual_review into inspect, this fails.
    for (const rel of ['packages/mcp-server/src/lib/mcp-tool-adapter.ts', 'packages/mcp-server/src/lib/session-store.ts']) {
      const abs = path.resolve(__dirname, '..', '..', rel);
      const src = fs.readFileSync(abs, 'utf8');
      expect(src).not.toMatch(/from\s+['"]\.\.?\/.*visual-review['"]/);
      expect(src).not.toMatch(/require\(['"].*visual-review['"]\)/);
    }
  });
});
