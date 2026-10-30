/**
 * H3 guard extensions: XXE-adjacent vectors + SVG active content + external
 * rels (split per D5 — strip for sanitize paths, refuse for compose ingest).
 *
 * - xxe-guard predicates flag xi:include, xsl:, and xml-stylesheet PIs while
 *   leaving the `<?xml ...?>` declaration, comments, and escaped math alone.
 * - SVG predicates flag event-handler attributes, foreignObject, script, and
 *   handler/listener elements while letting quoted prose (`title="turn
 *   on=off"`) and plain shapes pass.
 * - stripSvgActiveContent removes each class and reports kinds (loud,
 *   counts-by-kind — the sanitize-path contract).
 * - External-rel targets classify against the scheme allowlist and every
 *   decision lands on the audit trail.
 * - Compose ingest refuses hostile SVGs (file + inline data-URL) and keeps
 *   legit SVGs/images green.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  assertNoSvgActiveContent,
  assertNoXxeAdjacent,
  auditExternalRelTarget,
  classifyExternalRelTarget,
  clearExternalRelAuditTrail,
  containsSvgActiveContent,
  containsSvgDangerousHref,
  containsSvgEventHandler,
  containsSvgStyleElement,
  containsStylesheetPI,
  containsXInclude,
  containsXslt,
  containsXxeAdjacent,
  getExternalRelAuditTrail,
  stripSvgActiveContent,
} from '../../packages/engine/src/ooxml/xxe-guard';
import { compileMarkdown } from '../../packages/engine/src/compose/markdown';

describe('XXE-adjacent predicates', () => {
  const EVIL = [
    '<xi:include href="evil.xml" parse="xml"/>',
    '<XI:INCLUDE href="evil.xml"/>',
    '<xsl:stylesheet version="1.0"/>',
    '<xsl:value-of select="/"/>',
    '<doc xmlns:xsl="http://www.w3.org/1999/XSL/Transform"/>',
    '<?xml-stylesheet href="evil.xsl" type="text/xsl"?>',
    '<?XML-STYLESHEET href="evil.xsl"?>',
  ];
  const LEGIT = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<!-- a plain comment -->',
    '<w:t>profit &lt; forecast, a &gt; b</w:t>',
    '<root><child>text</child></root>',
    '',
  ];

  it('flags every adjacent vector', () => {
    for (const evil of EVIL) {
      expect(containsXxeAdjacent(evil), evil).toBe(true);
    }
    expect(containsXInclude('<xi:include href="x"/>')).toBe(true);
    expect(containsXslt('<xsl:template match="/"/>')).toBe(true);
    expect(containsStylesheetPI('<?xml-stylesheet href="x"?>')).toBe(true);
  });

  it('leaves declarations, comments, and math alone', () => {
    for (const legit of LEGIT) {
      expect(containsXxeAdjacent(legit), legit).toBe(false);
      expect(() => assertNoXxeAdjacent(legit, 'probe')).not.toThrow();
    }
    expect(containsXxeAdjacent(undefined)).toBe(false);
    expect(containsXxeAdjacent(null)).toBe(false);
  });

  it('fails closed with E_VALIDATION plus a remedy', () => {
    for (const evil of EVIL) {
      expect(() => assertNoXxeAdjacent(evil, 'probe'), evil).toThrowError(/E_VALIDATION/);
      expect(() => assertNoXxeAdjacent(evil, 'probe'), evil).toThrowError(/Remedy/);
    }
  });
});

describe('SVG active-content predicates', () => {
  const EVIL = [
    '<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"></svg>',
    '<svg><circle ONCLICK="evil()"/></svg>',
    '<svg><g onmouseover="x"><rect/></g></svg>',
    '<svg><foreignObject><body>hi</body></foreignObject></svg>',
    '<svg><FOREIGNOBJECT/></svg>',
    '<svg><script>alert(1)</script></svg>',
    '<svg><handler>evil</handler></svg>',
    '<svg><listener event="click"/></svg>',
  ];
  const LEGIT = [
    '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>',
    '<svg viewBox="0 0 10 10"><circle cx="5" cy="5" r="4"/></svg>',
    '<svg><a title="turn on=off today"><text>hi</text></a></svg>',
    '<svg><text>call onload=maybe later</text></svg>',
  ];

  it('flags every script-capable shape', () => {
    for (const evil of EVIL) {
      expect(containsSvgActiveContent(evil), evil).toBe(true);
      expect(() => assertNoSvgActiveContent(evil, 'probe'), evil).toThrowError(/E_VALIDATION/);
    }
    expect(containsSvgEventHandler('<svg onload="x"/>')).toBe(true);
  });

  it('leaves plain shapes and quoted prose alone', () => {
    for (const legit of LEGIT) {
      expect(containsSvgActiveContent(legit), legit).toBe(false);
      expect(() => assertNoSvgActiveContent(legit, 'probe')).not.toThrow();
    }
  });
});

describe('stripSvgActiveContent (sanitize-path contract)', () => {
  it('strips each class and reports kinds loudly', () => {
    const hostile =
      '<?xml-stylesheet href="e.xsl"?>' +
      '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" onload="e()">' +
      '<script>alert(1)</script>' +
      '<foreignObject><body>x</body></foreignObject>' +
      '<xi:include href="e.xml"/>' +
      '<rect width="10" height="10"/>' +
      '<xsl:value-of select="/"/>' +
      '</svg>';
    const { cleaned, removed } = stripSvgActiveContent(hostile);
    expect(cleaned).not.toMatch(/onload/i);
    expect(cleaned).not.toMatch(/<script/i);
    expect(cleaned).not.toMatch(/foreignObject/i);
    expect(cleaned).not.toMatch(/xi:include/i);
    expect(cleaned).not.toMatch(/xsl:/i);
    expect(cleaned).not.toMatch(/xml-stylesheet/i);
    expect(cleaned).toContain('<rect');
    for (const kind of ['script', 'foreignObject', 'event-handler-attrs', 'xi:include', 'xsl:stylesheet', 'xml-stylesheet-PI'] as const) {
      expect(removed).toContain(kind);
    }
    // Stripped output is itself clean.
    expect(containsSvgActiveContent(cleaned)).toBe(false);
    expect(containsXxeAdjacent(cleaned)).toBe(false);
  });

  it('leaves legit SVGs byte-identical with an empty report', () => {
    const legit = '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>';
    expect(stripSvgActiveContent(legit)).toEqual({ cleaned: legit, removed: [] });
  });

  it('never throws on malformed input', () => {
    expect(() => stripSvgActiveContent('<svg><unclosed')).not.toThrow();
    expect(stripSvgActiveContent('')).toEqual({ cleaned: '', removed: [] });
  });
});

describe('external-rel allowlist + audit trail', () => {
  beforeEach(() => clearExternalRelAuditTrail());
  afterEach(() => clearExternalRelAuditTrail());

  it('allows allowlisted schemes and relative targets', () => {
    for (const target of ['https://example.com/x', 'http://example.com', 'mailto:a@b.c', '/docs/x', '#frag', '']) {
      expect(classifyExternalRelTarget(target).allowed, target).toBe(true);
    }
  });

  it('blocks dangerous schemes', () => {
    for (const target of ['javascript:alert(1)', 'JaVaScRiPt:x', 'data:text/html,x', 'file:///etc/passwd', 'vbscript:x']) {
      const decision = classifyExternalRelTarget(target);
      expect(decision.allowed, target).toBe(false);
      expect(decision.reason).toBe('blocked-scheme');
    }
  });

  it('every audited decision lands on the bounded trail', () => {
    const ok = auditExternalRelTarget('https://example.com/x');
    const bad = auditExternalRelTarget('javascript:alert(1)');
    expect(ok.allowed).toBe(true);
    expect(bad.allowed).toBe(false);
    const trail = getExternalRelAuditTrail();
    expect(trail.length).toBe(2);
    expect(trail[0].reason).toBe('allowlisted-scheme');
    expect(trail[1].reason).toBe('blocked-scheme');
    expect(trail[1].scheme).toBe('javascript:');
  });
});

describe('compose ingest refuses hostile SVGs, keeps legit assets green', () => {
  const LEGIT_SVG =
    '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50">' +
    '<rect width="100" height="50" fill="blue"/></svg>';
  const HOSTILE_SVGS: Array<[string, string]> = [
    ['event attr', '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50" onload="alert(1)"><rect width="1" height="1"/></svg>'],
    ['foreignObject', '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"><foreignObject><body>hi</body></foreignObject></svg>'],
    ['script', '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"><script>alert(1)</script></svg>'],
    ['xi:include', '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"><xi:include href="e.xml" xmlns:xi="http://www.w3.org/2001/XInclude"/></svg>'],
    ['stylesheet PI', '<?xml-stylesheet href="e.xsl"?><svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"><rect width="1" height="1"/></svg>'],
    ['xsl', '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"><xsl:value-of select="/"/></svg>'],
  ];

  function assetDir(files: Record<string, string>): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-svg-'));
    for (const [name, content] of Object.entries(files)) {
      fs.writeFileSync(path.join(dir, name), content);
    }
    return dir;
  }

  it('refuses every hostile SVG file with E_VALIDATION + line number', () => {
    for (const [label, svg] of HOSTILE_SVGS) {
      const dir = assetDir({ 'evil.svg': svg });
      try {
        let message = '';
        try {
          compileMarkdown('![Evil](evil.svg)\n', { allowedDir: dir });
        } catch (err: unknown) {
          message = err instanceof Error ? err.message : String(err);
        }
        expect(message, label).toContain('E_VALIDATION');
        expect(message, label).toContain('line 1');
      } finally {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    }
  });

  it('refuses hostile inline data-URL SVGs', () => {
    const hostile = HOSTILE_SVGS[0][1];
    const b64 = Buffer.from(hostile, 'utf8').toString('base64');
    let message = '';
    try {
      compileMarkdown(`![Evil](data:image/svg+xml;base64,${b64})\n`);
    } catch (err: unknown) {
      message = err instanceof Error ? err.message : String(err);
    }
    expect(message).toContain('E_VALIDATION');
  });

  it('keeps legit SVGs and PNGs green with fitted sizes', () => {
    const png1x1 = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      'base64'
    );
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-svg-legit-'));
    try {
      fs.writeFileSync(path.join(dir, 'ok.svg'), LEGIT_SVG);
      fs.writeFileSync(path.join(dir, 'ok.png'), png1x1);
      const ops = compileMarkdown('![Ok](ok.svg)\n\n![Png](ok.png)\n', { allowedDir: dir });
      const images = ops.filter((op) => (op as { type?: string }).type === 'image');
      expect(images.length).toBe(2);
      const svgOp = images[0] as unknown as { width?: number; height?: number; fileData?: string };
      expect(svgOp.fileData).toContain('data:image/svg+xml;base64,');
      expect(svgOp.width).toBeGreaterThan(0);
      expect(svgOp.height).toBeGreaterThan(0);
      // Legit inline SVG data URL also passes.
      const inlineB64 = Buffer.from(LEGIT_SVG, 'utf8').toString('base64');
      const inlineOps = compileMarkdown(`![Ok](data:image/svg+xml;base64,${inlineB64})\n`);
      expect(inlineOps.filter((op) => (op as { type?: string }).type === 'image').length).toBe(1);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe('SVG dangerous hrefs (R1 refined: refuse-only js/vbs/data:text/html)', () => {
  // Mirrors /tmp evidence f4-live.ts transcript shapes: tag-scoped href scan,
  // prose never trips, https + benign data pass. Note (documented best-effort
  // miss, same class as the event-handler scan): a quoted `>` inside an href
  // value ends the tag scan early, so data:text/html payloads carrying literal
  // `<...>` are pinned via base64 / bracket-free shapes here; bracket-carrying
  // payloads still refuse via the script predicate, pinned separately below.
  const cases: Array<[string, string, boolean]> = [
    ['js href', '<svg xmlns="http://www.w3.org/2000/svg"><a xlink:href="javascript:alert(1)"><text>x</text></a></svg>', true],
    ['JS upper href', '<svg xmlns="http://www.w3.org/2000/svg"><a href="JaVaScRiPt:alert(1)"><text>x</text></a></svg>', true],
    ['vbs href', '<svg xmlns="http://www.w3.org/2000/svg"><a href="vbscript:msgbox(1)"><text>x</text></a></svg>', true],
    ['data html href', '<svg xmlns="http://www.w3.org/2000/svg"><image href="data:text/html,hello"/></svg>', true],
    ['data html base64', '<svg xmlns="http://www.w3.org/2000/svg"><image href="data:text/html;base64,PHNjcmlwdA=="/></svg>', true],
    ['DATA HTML upper', '<svg xmlns="http://www.w3.org/2000/svg"><image href="DATA:TEXT/HTML,hello"/></svg>', true],
    ['prose js text', '<svg xmlns="http://www.w3.org/2000/svg"><text>see javascript:foo for details</text></svg>', false],
    ['prose data text', '<svg xmlns="http://www.w3.org/2000/svg"><text>data:foo is fine here</text></svg>', false],
    ['legit https href', '<svg xmlns="http://www.w3.org/2000/svg"><a href="https://example.com"><text>x</text></a></svg>', false],
    ['legit data-image raster', '<svg xmlns="http://www.w3.org/2000/svg"><image href="data:image/png;base64,iVBORw0KGgo="/></svg>', false],
    ['legit data-font', '<svg xmlns="http://www.w3.org/2000/svg"><image href="data:font/woff2;base64,AAAA"/></svg>', false],
    ['legit title prose', '<svg><a title="turn on=off today"><text>hi</text></a></svg>', false],
    ['legit plain', '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>', false],
  ];

  it('flags refuse-only hrefs, allows benign data + https + prose', () => {
    for (const [label, svg, want] of cases) {
      expect(containsSvgDangerousHref(svg), label).toBe(want);
      expect(containsSvgActiveContent(svg), `active:${label}`).toBe(want);
    }
  });

  it('refuse path throws E_VALIDATION + Remedy on hostile hrefs only', () => {
    for (const [label, svg, want] of cases) {
      if (want) {
        expect(() => assertNoSvgActiveContent(svg, 'probe'), label).toThrowError(/E_VALIDATION/);
        expect(() => assertNoSvgActiveContent(svg, 'probe'), label).toThrowError(/Remedy/);
      } else {
        expect(() => assertNoSvgActiveContent(svg, 'probe'), label).not.toThrow();
      }
    }
  });

  it('strip drops the hostile href attr, keeps element + benign hrefs, reports loudly', () => {
    const hostile =
      '<svg xmlns="http://www.w3.org/2000/svg">' +
      '<a xlink:href="javascript:alert(1)"><text>x</text></a>' +
      '<image href="data:text/html;base64,PHNjcmlwdA=="/>' +
      '<image href="data:image/png;base64,AAAA"/>' +
      '<a href="https://example.com"><text>y</text></a>' +
      '</svg>';
    const { cleaned, removed } = stripSvgActiveContent(hostile);
    expect(cleaned).not.toContain('javascript:');
    expect(cleaned).not.toContain('data:text/html');
    expect(cleaned).toContain('data:image/png;base64,AAAA');
    expect(cleaned).toContain('https://example.com');
    expect(cleaned).toContain('<text>x</text>');
    expect(removed).toContain('dangerous-href');
    expect(containsSvgActiveContent(cleaned)).toBe(false);
  });

  it('bracket-carrying data:text/html still refuses (script predicate covers the tag-split miss)', () => {
    // Documented best-effort miss: quoted `>` ends the tag scan early, so the
    // href predicate alone misses `data:text/html,<script>…`. The embedded
    // `<script>` still trips the refuse path — never silently allowed.
    const bracketed =
      '<svg xmlns="http://www.w3.org/2000/svg"><image href="data:text/html,<script>alert(1)</script>"/></svg>';
    expect(containsSvgActiveContent(bracketed)).toBe(true);
    expect(() => assertNoSvgActiveContent(bracketed, 'probe')).toThrowError(/E_VALIDATION/);
  });
});

describe('SVG style-element (R1 strip-only: ingest allows, sanitize strips loudly)', () => {
  it('detects <style> but does NOT refuse it on ingest', () => {
    const styled =
      '<svg xmlns="http://www.w3.org/2000/svg"><style>.c{fill:red}</style><rect width="1" height="1"/></svg>';
    expect(containsSvgStyleElement(styled)).toBe(true);
    expect(containsSvgActiveContent(styled)).toBe(false);
    expect(() => assertNoSvgActiveContent(styled, 'probe')).not.toThrow();
    expect(containsSvgStyleElement('<STYLE>.c{fill:red}</STYLE>')).toBe(true);
    // spreadsheetML <styleSheet> is a different layer, never flagged.
    expect(
      containsSvgStyleElement(
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>'
      )
    ).toBe(false);
  });

  it('sanitize strips <style> loudly and keeps shapes', () => {
    const hostile =
      '<svg xmlns="http://www.w3.org/2000/svg"><style>.c{fill:red}</style><rect width="1" height="1"/></svg>';
    const { cleaned, removed } = stripSvgActiveContent(hostile);
    expect(cleaned).not.toMatch(/<\s*style/i);
    expect(cleaned).toContain('<rect');
    expect(removed).toContain('style-element');
  });
});

describe('Illustrator-style SVG (R1 proof: ingest pass + sanitize raster intact)', () => {
  const ILLUSTRATOR =
    '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50">' +
    '<style>.st0{fill:#ff0000}</style>' +
    '<image href="data:image/png;base64,iVBORw0KGgo=" width="10" height="10"/>' +
    '<a href="https://example.com"><text>hi</text></a>' +
    '<rect width="10" height="10" class="st0"/>' +
    '</svg>';

  function assetDir(files: Record<string, string>): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-svg-ill-'));
    for (const [name, content] of Object.entries(files)) {
      fs.writeFileSync(path.join(dir, name), content);
    }
    return dir;
  }

  it('passes ingest (file + inline data-URL) with style + raster + https', () => {
    expect(() => assertNoSvgActiveContent(ILLUSTRATOR, 'probe')).not.toThrow();
    const dir = assetDir({ 'ill.svg': ILLUSTRATOR });
    try {
      const ops = compileMarkdown('![Ill](ill.svg)\n', { allowedDir: dir });
      expect(ops.filter((op) => (op as { type?: string }).type === 'image').length).toBe(1);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
    const b64 = Buffer.from(ILLUSTRATOR, 'utf8').toString('base64');
    const inlineOps = compileMarkdown(`![Ill](data:image/svg+xml;base64,${b64})\n`);
    expect(inlineOps.filter((op) => (op as { type?: string }).type === 'image').length).toBe(1);
  });

  it('sanitizes with raster + https intact and style-strip reported', () => {
    const { cleaned, removed } = stripSvgActiveContent(ILLUSTRATOR);
    expect(cleaned).toContain('data:image/png;base64,iVBORw0KGgo=');
    expect(cleaned).toContain('https://example.com');
    expect(cleaned).toContain('<rect');
    expect(cleaned).not.toMatch(/<\s*style/i);
    expect(removed).toContain('style-element');
    expect(removed).not.toContain('dangerous-href');
  });

  it('hostile script / js-href still refused on ingest', () => {
    for (const svg of [
      '<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>',
      '<svg xmlns="http://www.w3.org/2000/svg"><a href="javascript:alert(1)"><text>x</text></a></svg>',
    ]) {
      expect(() => assertNoSvgActiveContent(svg, 'probe')).toThrowError(/E_VALIDATION/);
      const dir = assetDir({ 'evil.svg': svg });
      try {
        let message = '';
        try {
          compileMarkdown('![Evil](evil.svg)\n', { allowedDir: dir });
        } catch (err: unknown) {
          message = err instanceof Error ? err.message : String(err);
        }
        expect(message).toContain('E_VALIDATION');
      } finally {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    }
  });
});
