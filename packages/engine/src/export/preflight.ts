import AdmZip from '../adm-zip-safe';

export type PreflightSeverity = 'blocker' | 'warning';

export interface PreflightIssue {
  check: string;
  severity: PreflightSeverity;
  code: string;
  message: string;
}

export interface PreflightReport {
  blockers: PreflightIssue[];
  warnings: PreflightIssue[];
}

export interface PreflightContext {
  names: Set<string>;
  readText: (name: string) => string | null;
}

export interface PreflightCheck {
  id: string;
  run: (ctx: PreflightContext) => PreflightIssue[];
}

export class PreflightError extends Error {
  code = 'E_PREFLIGHT';
  blockers: PreflightIssue[];
  constructor(blockers: PreflightIssue[]) {
    super(
      `[E_PREFLIGHT] PDF export blocked by ${blockers.length} preflight blocker(s): ` +
        blockers.map((b) => `${b.code} ${b.message}`).join('; ')
    );
    this.name = 'PreflightError';
    this.blockers = blockers;
  }
}

function collectAttrRefs(xml: string, tag: string, attr: string): string[] {
  const out: string[] = [];
  const tagRe = new RegExp(`<${tag}\\b[^>]*>`, 'g');
  let m: RegExpExecArray | null;
  while ((m = tagRe.exec(xml)) !== null) {
    const attrRe = new RegExp(`${attr}="([^"]+)"`);
    const a = attrRe.exec(m[0]);
    if (a) out.push(a[1]);
  }
  return out;
}

function parseRels(xml: string): Map<string, string> {
  const map = new Map<string, string>();
  const re = /<Relationship\b[^>]*>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    const id = /Id="([^"]+)"/.exec(m[0]);
    const target = /Target="([^"]+)"/.exec(m[0]);
    if (id && target) map.set(id[1], target[1]);
  }
  return map;
}

function resolvePart(baseDir: string, target: string): string {
  const clean = target.split('#')[0];
  if (clean.startsWith('/')) return clean.slice(1);
  return `${baseDir}/${clean}`;
}

function extractTextRuns(xml: string): string[] {
  const out: string[] = [];
  const re = /<(?:w:t|w:delText)(?:\s[^>]*)?>([\s\S]*?)<\/(?:w:t|w:delText)>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) out.push(m[1]);
  return out;
}

/**
 * Decodes numeric XML character references (`&#61623;`, `&#xF0B7;`) so
 * downstream scans see the real code points. Named entities and malformed
 * references pass through untouched; out-of-range values are left as-is.
 */
function unescapeNumericEntities(s: string): string {
  if (!s.includes('&#')) return s;
  return s.replace(/&#(?:(\d+)|[xX]([0-9a-fA-F]+));/g, (whole, dec?: string, hex?: string) => {
    const cp = dec !== undefined ? parseInt(dec, 10) : parseInt(hex ?? '', 16);
    if (!Number.isFinite(cp) || cp < 0 || cp > 0x10ffff) return whole;
    try {
      return String.fromCodePoint(cp);
    } catch {
      return whole;
    }
  });
}

const checkNotesCompanions: PreflightCheck = {
  id: 'notes-companions',
  run: (ctx) => {
    const issues: PreflightIssue[] = [];
    const doc = ctx.readText('word/document.xml');
    if (!doc) return issues;
    if (/<w:footnoteReference\b/.test(doc) && !ctx.names.has('word/footnotes.xml')) {
      issues.push({
        check: 'notes-companions',
        severity: 'blocker',
        code: 'E_NOTES_PART',
        message: 'document references footnotes but word/footnotes.xml is missing from the package',
      });
    }
    if (/<w:endnoteReference\b/.test(doc) && !ctx.names.has('word/endnotes.xml')) {
      issues.push({
        check: 'notes-companions',
        severity: 'blocker',
        code: 'E_NOTES_PART',
        message: 'document references endnotes but word/endnotes.xml is missing from the package',
      });
    }
    return issues;
  },
};

const checkChartCompanions: PreflightCheck = {
  id: 'chart-companions',
  run: (ctx) => {
    const issues: PreflightIssue[] = [];
    const doc = ctx.readText('word/document.xml');
    if (!doc) return issues;
    const relsText = ctx.readText('word/_rels/document.xml.rels');
    const rels = relsText ? parseRels(relsText) : new Map<string, string>();
    const re = /<c:chart\b[^>]*?r:id="([^"]+)"/g;
    let m: RegExpExecArray | null;
    const seen = new Set<string>();
    while ((m = re.exec(doc)) !== null) {
      const rid = m[1];
      if (seen.has(rid)) continue;
      seen.add(rid);
      const target = rels.get(rid);
      if (!target) {
        issues.push({
          check: 'chart-companions',
          severity: 'blocker',
          code: 'E_CHART_PART',
          message: `chart reference "${rid}" has no relationship entry in word/_rels/document.xml.rels`,
        });
        continue;
      }
      const part = resolvePart('word', target);
      if (!ctx.names.has(part)) {
        issues.push({
          check: 'chart-companions',
          severity: 'blocker',
          code: 'E_CHART_PART',
          message: `chart reference "${rid}" points at "${part}" which is missing from the package`,
        });
      }
    }
    return issues;
  },
};

const checkNumberingRefs: PreflightCheck = {
  id: 'numbering-refs',
  run: (ctx) => {
    const issues: PreflightIssue[] = [];
    const doc = ctx.readText('word/document.xml');
    if (!doc) return issues;
    const refs = new Set<string>();
    for (const v of collectAttrRefs(doc, 'w:numId', 'w:val')) {
      if (v !== '0') refs.add(v);
    }
    if (refs.size === 0) return issues;
    const numbering = ctx.readText('word/numbering.xml');
    if (!numbering) {
      issues.push({
        check: 'numbering-refs',
        severity: 'warning',
        code: 'E_NUMBERING_PART',
        message: `document uses numbering (${refs.size} numId(s)) but word/numbering.xml is missing; lists render unnumbered`,
      });
      return issues;
    }
    const defined = new Set(collectAttrRefs(numbering, 'w:num', 'w:numId'));
    const missing = [...refs].filter((r) => !defined.has(r));
    if (missing.length > 0) {
      issues.push({
        check: 'numbering-refs',
        severity: 'warning',
        code: 'E_NUMBERING_REF',
        message: `numId(s) ${missing.slice(0, 5).join(', ')}${missing.length > 5 ? ` (+${missing.length - 5} more)` : ''} have no definition in word/numbering.xml`,
      });
    }
    return issues;
  },
};

const checkFonts: PreflightCheck = {
  id: 'fonts',
  run: (ctx) => {
    const issues: PreflightIssue[] = [];
    // Collect rFonts references before touching the table so the no-table
    // path stays sound: word/fontTable.xml is optional, and without declared
    // fonts no reference can be verified. Any referenced font with no table
    // warns as unknown (E_FONT_REF) instead of reporting false-clean, while
    // documents carrying no rFonts references stay clean.
    const doc = ctx.readText('word/document.xml') ?? '';
    const styles = ctx.readText('word/styles.xml') ?? '';
    const refs = new Set<string>();
    for (const xml of [doc, styles]) {
      const tagRe = /<w:rFonts\b[^>]*>/g;
      let m: RegExpExecArray | null;
      while ((m = tagRe.exec(xml)) !== null) {
        for (const attr of ['w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia']) {
          const a = new RegExp(`${attr}="([^"]+)"`).exec(m[0]);
          if (a) refs.add(a[1]);
        }
      }
    }
    const table = ctx.readText('word/fontTable.xml');
    if (!table) {
      if (refs.size > 0) {
        const missing = [...refs];
        issues.push({
          check: 'fonts',
          severity: 'warning',
          code: 'E_FONT_REF',
          message: `font(s) ${missing.slice(0, 5).join(', ')}${missing.length > 5 ? ` (+${missing.length - 5} more)` : ''} are referenced but word/fontTable.xml is absent; declared fonts unknown, renderer substitutes`,
        });
      }
      return issues;
    }
    const defined = new Set(collectAttrRefs(table, 'w:font', 'w:name').map((n) => n.toLowerCase()));
    const missing = [...refs].filter((r) => !defined.has(r.toLowerCase()));
    if (missing.length > 0) {
      issues.push({
        check: 'fonts',
        severity: 'warning',
        code: 'E_FONT_REF',
        message: `font(s) ${missing.slice(0, 5).join(', ')}${missing.length > 5 ? ` (+${missing.length - 5} more)` : ''} are referenced but absent from word/fontTable.xml; renderer substitutes`,
      });
    }
    const relsText = ctx.readText('word/_rels/fontTable.xml.rels');
    const rels = relsText ? parseRels(relsText) : new Map<string, string>();
    const embedRe = /<w:embed(?:Regular|Bold|Italic|BoldItalic)\b[^>]*?r:id="([^"]+)"/g;
    let m: RegExpExecArray | null;
    const dangling: string[] = [];
    while ((m = embedRe.exec(table)) !== null) {
      const target = rels.get(m[1]);
      const part = target ? resolvePart('word', target) : null;
      if (!part || !ctx.names.has(part)) dangling.push(m[1]);
    }
    if (dangling.length > 0) {
      issues.push({
        check: 'fonts',
        severity: 'warning',
        code: 'E_FONT_EMBED',
        message: `embedded font relationship(s) ${dangling.slice(0, 5).join(', ')} point at missing font data; text falls back to system fonts`,
      });
    }
    return issues;
  },
};

const checkPendingRedlines: PreflightCheck = {
  id: 'pending-redlines',
  run: (ctx) => {
    const issues: PreflightIssue[] = [];
    const doc = ctx.readText('word/document.xml');
    if (!doc) return issues;
    const all: Array<[string, number]> = [
      ['w:ins', (doc.match(/<w:ins\b/g) || []).length],
      ['w:del', (doc.match(/<w:del\b/g) || []).length],
      ['w:moveFrom', (doc.match(/<w:moveFrom\b/g) || []).length],
      ['w:moveTo', (doc.match(/<w:moveTo\b/g) || []).length],
    ];
    const counts = all.filter(([, n]) => n > 0);
    if (counts.length > 0) {
      issues.push({
        check: 'pending-redlines',
        severity: 'warning',
        code: 'E_REDLINES',
        message: `document carries unaccepted tracked revisions (${counts.map(([k, n]) => `${k}=${n}`).join(', ')}); export renders markup or accepted text depending on renderer`,
      });
    }
    return issues;
  },
};

const checkRiskyGlyphs: PreflightCheck = {
  id: 'risky-glyphs',
  run: (ctx) => {
    const issues: PreflightIssue[] = [];
    const doc = ctx.readText('word/document.xml');
    if (!doc) return issues;
    const found = new Set<string>();
    for (const run of extractTextRuns(doc)) {
      for (const ch of unescapeNumericEntities(run)) {
        const cp = ch.codePointAt(0) ?? 0;
        if (cp === 0xfffd || (cp >= 0xe000 && cp <= 0xf8ff) || cp > 0xffff) {
          found.add(`U+${cp.toString(16).toUpperCase().padStart(4, '0')}`);
          if (found.size >= 8) break;
        }
      }
      if (found.size >= 8) break;
    }
    if (found.size > 0) {
      issues.push({
        check: 'risky-glyphs',
        severity: 'warning',
        code: 'E_RISKY_GLYPH',
        message: `text uses glyphs outside common font coverage (${[...found].slice(0, 5).join(', ')}); renderer may substitute or drop them`,
      });
    }
    return issues;
  },
};

export const preflightChecks: PreflightCheck[] = [
  checkNotesCompanions,
  checkChartCompanions,
  checkNumberingRefs,
  checkFonts,
  checkPendingRedlines,
  checkRiskyGlyphs,
];

export function runPreflightChecks(docx: Buffer): PreflightReport {
  const blockers: PreflightIssue[] = [];
  const warnings: PreflightIssue[] = [];
  let ctx: PreflightContext;
  try {
    const zip = new AdmZip(docx);
    const entries = zip.getEntries();
    const names = new Set<string>();
    const cache = new Map<string, string | null>();
    for (const e of entries) names.add(e.entryName);
    const readText = (name: string): string | null => {
      if (cache.has(name)) return cache.get(name) ?? null;
      try {
        const entry = zip.getEntry(name);
        if (!entry) {
          cache.set(name, null);
          return null;
        }
        const text = entry.getData().toString('utf8');
        cache.set(name, text);
        return text;
      } catch {
        cache.set(name, null);
        return null;
      }
    };
    ctx = { names, readText };
  } catch {
    return {
      blockers: [
        {
          check: 'container',
          severity: 'blocker',
          code: 'E_CONTAINER',
          message: 'package is not a readable zip container',
        },
      ],
      warnings: [],
    };
  }
  for (const check of preflightChecks) {
    let found: PreflightIssue[];
    try {
      found = check.run(ctx);
    } catch {
      found = [
        {
          check: check.id,
          severity: 'warning',
          code: 'E_CHECK_FAILED',
          message: `check "${check.id}" failed internally; skipped`,
        },
      ];
    }
    for (const issue of found) {
      if (issue.severity === 'blocker') blockers.push(issue);
      else warnings.push(issue);
    }
  }
  return { blockers, warnings };
}
