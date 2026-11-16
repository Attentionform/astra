/**
 * DN3 Phase 0 spike: the prefix canonicaliser and its independent verifier. Not wired into any code path.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocxParser } from '@attentionform/astra-engine';
import {
  CANONICAL_PREFIXES, PrefixRefusal, canonicalisePart, canonicalisePrefixes, ingestPrefixes, scanPart, verifyEquivalent,
} from '../../packages/engine/src/ooxml/prefix-canonicaliser';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const UNKNOWN = 'urn:example:extension';

const doc = (rootAttrs: string, body: string, p = 'w'): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${p}:document ${rootAttrs}><${p}:body>${body}</${p}:body></${p}:document>`;

const refusalOf = (fn: () => unknown): PrefixRefusal => {
  try { fn(); } catch (e) { if (e instanceof PrefixRefusal) return e; throw e; }
  throw new Error('expected a PrefixRefusal');
};

describe('scanPart', () => {
  it('is null for a canonical part (no rewrite, no cost beyond the scan)', () => {
    const xml = doc(`xmlns:w="${W}" xmlns:r="${R}"`, '<w:p><w:r><w:t>x</w:t></w:r></w:p>');
    expect(scanPart(xml)).toBeNull();
    expect(canonicalisePrefixes(xml).xml).toBe(xml);
  });
  it('reports a renamed prefix, the default namespace and a local redeclaration', () => {
    expect(scanPart(doc(`xmlns:x="${W}"`, '', 'x'))!.nonCanonical).toEqual([{ from: 'x', to: 'w', uri: W }]);
    expect(scanPart(`<document xmlns="${W}"/>`)!.nonCanonical[0]).toMatchObject({ from: '', to: 'w' });
    expect(scanPart(`<w:document xmlns:w="${W}"><x:tbl xmlns:x="${W}"/></w:document>`)!.nonCanonical[0]).toMatchObject({ from: 'x' });
  });
});

describe('scanPart: a retained alias is not work (idempotence)', () => {
  it('a declaration that no tag uses needs no rewrite', () => {
    expect(scanPart(`<w:document xmlns:w="${W}" xmlns:x="${W}"><w:body/></w:document>`)).toBeNull();
  });
  it('the same declaration used by an element name, a closing tag or an attribute does', () => {
    expect(scanPart(`<w:document xmlns:w="${W}" xmlns:x="${W}"><x:body/></w:document>`)).not.toBeNull();
    expect(scanPart(`<w:document xmlns:w="${W}" xmlns:x="${W}"><w:body x:val="1"></w:body></w:document>`)).not.toBeNull();
  });
  it('text, comments and values that merely contain "x:name" do not count as use', () => {
    expect(scanPart(`<w:document xmlns:w="${W}" xmlns:x="${W}"><w:t>see x:body here</w:t><!-- <x:body/> --><w:p e="a x:b"/></w:document>`)).toBeNull();
  });
  it('a default-namespace declaration that no unprefixed element uses needs no rewrite; one that is used does', () => {
    expect(scanPart(`<w:document xmlns="${W}" xmlns:w="${W}"><w:body/></w:document>`)).toBeNull();
    expect(scanPart(`<document xmlns="${W}"><body/></document>`)).not.toBeNull();
  });
  it('canonicalising twice changes nothing the second time, and the second scan finds nothing', () => {
    const xml = doc(`xmlns:x="${W}"`, '<x:p x:rsidR="1"><x:r/></x:p>', 'x');
    const once = canonicalisePart(xml).xml;
    expect(scanPart(once)).toBeNull();
    expect(canonicalisePart(once)).toEqual({ xml: once, renames: [] });
  });
});

describe('deeply nested parts (no stack overflow)', () => {
  it('canonicalises and verifies a part nested 100,000 deep', () => {
    const depth = 100000;
    const xml = `<x:document xmlns:x="${W}">${'<x:sdt>'.repeat(depth)}<x:p/>${'</x:sdt>'.repeat(depth)}</x:document>`;
    const out = canonicalisePart(xml, 'deep.xml');
    expect(out.xml.startsWith('<w:document ')).toBe(true);
    expect(out.xml.endsWith('</w:document>')).toBe(true);
    expect(verifyEquivalent(xml, out.xml).ok).toBe(true);
  });
  it('a part that declares a new prefix at every one of 20,000 nesting levels finishes in useful time or is refused cleanly', () => {
    const depth = 20000;
    let open = '';
    for (let i = 0; i < depth; i++) open += `<x:sdt xmlns:p${i}="urn:p${i}">`;
    const xml = `<x:document xmlns:x="${W}">${open}<x:p/>${'</x:sdt>'.repeat(depth)}</x:document>`;
    const started = Date.now();
    let outcome = 'ok';
    try { canonicalisePart(xml, 'wide.xml'); } catch (e) { outcome = e instanceof PrefixRefusal ? 'refused' : `crash:${(e as Error).message}`; }
    expect(outcome).not.toMatch(/^crash/);
    expect(Date.now() - started).toBeLessThan(20000);
  }, 60000);

  it('a document nested 20,000 deep opens with another prefix as it does with the usual one', async () => {
    const { randomBytes } = await import('node:crypto');
    const { createBlankDocxBuffer: blank } = await import('@attentionform/astra-engine');
    for (const p of ['w', 'x']) {
      const zip = new AdmZip(blank('blank'));
      const open = Array.from({ length: 20000 }, () => `<${p}:sdt ${p}:val="${randomBytes(8).toString('hex')}">`).join('');
      zip.updateFile('word/document.xml', Buffer.from(`<${p}:document xmlns:${p}="${W}"><${p}:body>${open}<${p}:p/>${`</${p}:sdt>`.repeat(20000)}</${p}:body></${p}:document>`));
      expect(() => new DocxParser(zip.toBuffer())).not.toThrow();
    }
  }, 120000);
});

describe('a namespace written with character references', () => {
  const spelled = W.replace('http', '&#104;ttp').replace('wordprocessingml', 'wordprocessing&#x6D;l');
  it('is the same namespace: scanned, rewritten and verified like the plain spelling', () => {
    const xml = `<x:document xmlns:x="${spelled}"><x:body><x:p/></x:body></x:document>`;
    expect(scanPart(xml)).not.toBeNull();
    const out = canonicalisePart(xml, 'spelled.xml');
    expect(out.xml).toContain('<w:body>');
    expect(out.xml).toContain(`xmlns:w="${W}"`);
    expect(verifyEquivalent(xml, out.xml).ok).toBe(true);
  });
});

describe('a part written by a real producer', () => {
  it('pandoc writes word/webSettings.xml with the prefix ns0: it is rewritten and a second pass leaves it alone', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>\n<ns0:webSettings xmlns:ns0="${W}">\n  <ns0:allowPNG />\n  <ns0:doNotSaveAsSingleFile />\n</ns0:webSettings>`;
    const once = canonicalisePart(xml, 'word/webSettings.xml');
    expect(once.renames).toEqual([{ from: 'ns0', to: 'w', uri: W }]);
    expect(once.xml).toContain('<w:webSettings ');
    expect(once.xml).toContain('<w:allowPNG/>');
    expect(scanPart(once.xml)).toBeNull();
  });
});

describe('canonicalisePrefixes: what it rewrites', () => {
  it('renames elements and attributes, keeps the old declaration and adds the canonical one', () => {
    const xml = doc(`xmlns:x="${W}" xmlns:rel="${R}"`, '<x:p><x:r x:rsidR="00AB"><x:t>hi</x:t></x:r><x:drawing rel:id="rId1"/></x:p>', 'x');
    const out = canonicalisePart(xml).xml;
    expect(out).toContain('<w:document ');
    expect(out).toContain('xmlns:x="' + W + '"');
    expect(out).toContain('xmlns:w="' + W + '"');
    expect(out).toContain('w:rsidR="00AB"');
    expect(out).toContain('r:id="rId1"');
    expect(out).toContain('xmlns:r="' + R + '"');
    expect(out).toMatch(/<\/w:document>$/);
  });

  it('keeps text, CDATA, comments, entities, whitespace and xml:space byte for byte', () => {
    const body = '<x:p><!-- a <x:p> in a comment --><x:r><x:t xml:space="preserve">  a &amp; b &#x2014; <![CDATA[<x:t>not a tag</x:t>]]>  </x:t></x:r></x:p>\n  ';
    const xml = doc(`xmlns:x="${W}"`, body, 'x');
    const out = canonicalisePart(xml).xml;
    const strip = (s: string): string => s.replace(/<(?!!--|!\[CDATA\[|\?)[^>]*>/g, '<>');
    // outside of tags nothing changed: same text, same comment, same CDATA
    expect(strip(out).replace(/<>/g, '')).toBe(strip(xml).replace(/<>/g, ''));
    expect(out).toContain('<!-- a <x:p> in a comment -->');
    expect(out).toContain('<![CDATA[<x:t>not a tag</x:t>]]>');
    expect(out).toContain('xml:space="preserve"');
  });

  it('copes with > and quotes of both kinds inside attribute values', () => {
    const xml = doc(`xmlns:x="${W}"`, `<x:p x:val="a>b" x:other='say "hi"'/>`, 'x');
    const out = canonicalisePart(xml).xml;
    expect(out).toContain('w:val="a>b"');
    expect(out).toContain(`w:other='say "hi"'`);
  });

  it('handles the default namespace', () => {
    const xml = `<document xmlns="${W}" xmlns:r="${R}"><body><p><r r:id="rId1"/></p></body></document>`;
    const out = canonicalisePart(xml).xml;
    expect(out).toContain('<w:document ');
    expect(out).toContain('<w:r r:id="rId1"/>');
    expect(out).toContain(`xmlns="${W}"`); // kept
    expect(verifyEquivalent(xml, out).ok).toBe(true);
  });

  it('leaves unknown namespaces, their declarations and unprefixed attributes alone', () => {
    const xml = doc(`xmlns:x="${W}" xmlns:e="${UNKNOWN}"`, `<x:p><e:thing e:a="1" plain="2"/></x:p>`, 'x');
    const out = canonicalisePart(xml).xml;
    expect(out).toContain(`xmlns:e="${UNKNOWN}"`);
    expect(out).toContain('<e:thing e:a="1" plain="2"/>');
  });

  it('is idempotent', () => {
    const xml = doc(`xmlns:x="${W}"`, '<x:p/>', 'x');
    const once = canonicalisePart(xml).xml;
    expect(canonicalisePart(once).xml).toBe(once);
  });

  it('a renamed part and its canonical twin have the same resolved form', () => {
    const canonical = doc(`xmlns:w="${W}" xmlns:r="${R}" xmlns:mc="${MC}" xmlns:w14="${W14}" mc:Ignorable="w14"`, '<w:p w14:paraId="1"><w:r><w:t>x</w:t></w:r></w:p>');
    const renamed = doc(`xmlns:q="${W}" xmlns:r="${R}" xmlns:m="${MC}" xmlns:z="${W14}" m:Ignorable="z"`, '<q:p z:paraId="1"><q:r><q:t>x</q:t></q:r></q:p>', 'q');
    const out = canonicalisePart(renamed).xml;
    expect(out).toContain('mc:Ignorable="w14"'); // the list was rewritten to the canonical prefix
    expect(out).toContain('w14:paraId="1"');
    expect(verifyEquivalent(renamed, out).ok).toBe(true);
    // the canonical twin and the output agree on every node
    const a = verifyEquivalent(canonical, out);
    expect(a.ok, a.problem).toBe(true);
  });
});

describe('markup compatibility values', () => {
  const root = (extra = '') => `xmlns:q="${W}" xmlns:m="${MC}" xmlns:z="${W14}" xmlns:y="${W15}" xmlns:e="${UNKNOWN}" ${extra}`;
  it('rewrites Ignorable, MustUnderstand, ProcessContent and Requires by the namespace each token names', () => {
    const xml = doc(
      root('m:Ignorable="z y e" m:MustUnderstand="z" m:ProcessContent="q:p z:foo e:bar"'),
      `<m:AlternateContent><m:Choice Requires="z y"><q:p/></m:Choice><m:Fallback><q:p/></m:Fallback></m:AlternateContent>`,
      'q'
    );
    const out = canonicalisePart(xml).xml;
    expect(out).toContain('mc:Ignorable="w14 w15 e"');
    expect(out).toContain('mc:MustUnderstand="w14"');
    expect(out).toContain('mc:ProcessContent="w:p w14:foo e:bar"');
    expect(out).toContain('<mc:Choice Requires="w14 w15">');
    expect(verifyEquivalent(xml, out).ok).toBe(true);
  });

  it('an Ignorable that names an undeclared prefix is left as written (it was readable before and stays so)', () => {
    const xml = doc(`xmlns:q="${W}" xmlns:m="${MC}" m:Ignorable="ghost"`, '<q:p/>', 'q');
    expect(canonicalisePart(xml).xml).toContain('Ignorable="ghost"');
  });

  it('a prefix used only by Ignorable keeps its declaration', () => {
    const xml = doc(`xmlns:q="${W}" xmlns:m="${MC}" xmlns:z="${W14}" m:Ignorable="z"`, '<q:p/>', 'q');
    const out = canonicalisePart(xml).xml;
    expect(out).toContain(`xmlns:z="${W14}"`);
    expect(out).toContain(`xmlns:w14="${W14}"`);
  });
});

describe('values the module does not understand keep resolving', () => {
  it('an extension value that names the old prefix still resolves because the old binding is kept', () => {
    const xml = doc(`xmlns:x="${W}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:e="${UNKNOWN}"`, '<x:p><e:t xsi:type="x:CT_P" e:ref="x:name"/></x:p>', 'x');
    const out = canonicalisePart(xml).xml;
    expect(out).toContain('xsi:type="x:CT_P"');
    expect(out).toContain('e:ref="x:name"');
    const b = verifyEquivalent(xml, out);
    expect(b.ok, b.problem).toBe(true);
  });
});

describe('conflicts are refused, nothing is written', () => {
  it('the canonical prefix is bound to a different namespace in scope', () => {
    const xml = `<x:document xmlns:x="${W}" xmlns:w="${UNKNOWN}"><x:body/></x:document>`;
    expect(refusalOf(() => canonicalisePart(xml)).reason).toBe('conflict');
  });
  it('the canonical prefix is shadowed in a subtree that holds a table-namespace element', () => {
    const xml = `<x:document xmlns:x="${W}"><e xmlns:w="${UNKNOWN}"><x:body/></e></x:document>`;
    expect(refusalOf(() => canonicalisePart(xml)).reason).toBe('conflict');
  });
  it('two table namespaces that swap their conventional prefixes', () => {
    const xml = `<w:document xmlns:w="${A}" xmlns:a="${W}"><a:body/></w:document>`;
    expect(refusalOf(() => canonicalisePart(xml)).reason).toBe('conflict');
  });
  it('a document type declaration and broken markup', () => {
    expect(refusalOf(() => canonicalisePart(`<!DOCTYPE x [<!ENTITY e "v">]><x:a xmlns:x="${W}"/>`)).reason).toBe('malformed');
    expect(refusalOf(() => canonicalisePart(`<x:a xmlns:x="${W}"><x:b></x:a>`)).reason).toBe('malformed');
  });
  it('an undeclared prefix on an element', () => {
    expect(refusalOf(() => canonicalisePart(`<x:a xmlns:x="${W}"><q:b/></x:a>`)).reason).toBe('unresolved');
  });
});

describe('local rebinding', () => {
  it('the same prefix names a table namespace outside and an unknown one inside', () => {
    const xml = `<x:document xmlns:x="${W}"><x:body><x:p/><e x:foo="1" xmlns:x="${UNKNOWN}"><x:thing/></e><x:p/></x:body></x:document>`;
    const out = canonicalisePart(xml).xml;
    expect(out).toContain('<x:thing/>'); // unknown namespace: not renamed
    expect(out).toContain('<w:p/>');
    expect(verifyEquivalent(xml, out).ok).toBe(true);
  });
  it('a table namespace declared on an inner element only (one prefixed table in a canonical part)', () => {
    const xml = `<w:document xmlns:w="${W}"><w:body><x:tbl xmlns:x="${W}"><x:tr/></x:tbl></w:body></w:document>`;
    const out = canonicalisePart(xml).xml;
    expect(out).toContain('<w:tbl ');
    expect(out).toContain('<w:tr/>');
    expect(out).toContain('</w:tbl>');
    expect(verifyEquivalent(xml, out).ok).toBe(true);
  });
  it('the same URI under two prefixes on one element', () => {
    const xml = `<a:document xmlns:a="${W}" xmlns:b="${W}"><b:body/></a:document>`;
    const out = canonicalisePart(xml).xml;
    expect(out).toContain('<w:document ');
    expect(out).toContain('<w:body/>');
    expect(verifyEquivalent(xml, out).ok).toBe(true);
  });
});

describe('verifyEquivalent catches the losses it exists to catch', () => {
  const before = doc(`xmlns:q="${W}" xmlns:m="${MC}" xmlns:z="${W14}" xmlns:e="${UNKNOWN}" m:Ignorable="z"`, '<q:p z:paraId="1" e:note="x:y"><q:r><q:t>text</q:t></q:r></q:p>', 'q');
  const good = canonicalisePart(before).xml;
  it('accepts the real output', () => expect(verifyEquivalent(before, good).ok).toBe(true));

  const mutations: Array<[string, (s: string) => string]> = [
    ['drops the old alias declaration', (s) => s.replace(` xmlns:q="${W}"`, '')],
    ['drops a declaration that only an Ignorable needed', (s) => s.replace(` xmlns:z="${W14}"`, '').replace(` xmlns:w14="${W14}"`, ' ')],
    ['renames an element to the wrong namespace', (s) => s.replace('<w:t>', '<m:t>').replace('</w:t>', '</m:t>')],
    ['changes an attribute value', (s) => s.replace('w14:paraId="1"', 'w14:paraId="2"')],
    ['changes an unknown-namespace value that names the old prefix', (s) => s.replace('e:note="x:y"', 'e:note="w:y"')],
    ['changes text', (s) => s.replace('>text<', '>texT<')],
    ['drops a token from Ignorable', (s) => s.replace('mc:Ignorable="w14"', 'mc:Ignorable=""')],
    ['drops an element', (s) => s.replace('<w:r><w:t>text</w:t></w:r>', '')],
    ['binds the canonical prefix to another namespace', (s) => s.replace(`xmlns:w="${W}"`, `xmlns:w="${UNKNOWN}"`)],
  ];
  for (const [label, mutate] of mutations) {
    it(`rejects an output that ${label}`, () => {
      const bad = mutate(good);
      expect(bad).not.toBe(good);
      let result = { ok: true } as { ok: boolean; problem?: string };
      try { result = verifyEquivalent(before, bad); } catch { result = { ok: false }; }
      expect(result.ok, `mutation "${label}" went unnoticed`).toBe(false);
    });
  }
});

// ---------------------------------------------------------------------------------------------------------------
// Real Word-authored parts, renamed by an independent DOM-based renamer (not by the code under test)
// ---------------------------------------------------------------------------------------------------------------

const XMLNS = 'http://www.w3.org/2000/xmlns/';

/** Renames every table prefix in `xml` to `${prefix}${suffix}` using a DOM copy; mc lists and declarations follow. */
function renameWithDom(xml: string, suffix: string): string {
  const parser = new DOMParser();
  const src = parser.parseFromString(xml, 'text/xml');
  const dst = new DOMParser().parseFromString('<root/>', 'text/xml');
  const newPrefixOf = (uri: string, old: string): string => (CANONICAL_PREFIXES[uri] !== undefined ? CANONICAL_PREFIXES[uri] + suffix : old);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const scopeUri = (el: any, prefix: string): string | null => el.lookupNamespaceURI(prefix || null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const copy = (node: any): any => {
    if (node.nodeType === 3) return dst.createTextNode(node.data);
    if (node.nodeType === 4) return dst.createCDATASection(node.data);
    if (node.nodeType === 8) return dst.createComment(node.data);
    if (node.nodeType !== 1) return null;
    const uri = node.namespaceURI as string | null;
    const el = uri ? dst.createElementNS(uri, `${newPrefixOf(uri, node.prefix ?? '')}:${node.localName}`) : dst.createElement(node.tagName);
    for (let i = 0; i < node.attributes.length; i++) {
      const a = node.attributes[i];
      if (a.name === 'xmlns' || a.name.startsWith('xmlns:')) {
        const u = a.value;
        const old = a.name === 'xmlns' ? '' : a.name.slice(6);
        const np = CANONICAL_PREFIXES[u] !== undefined ? CANONICAL_PREFIXES[u] + suffix : old;
        if (np === '' ) el.setAttributeNS(XMLNS, 'xmlns', u);
        else el.setAttributeNS(XMLNS, `xmlns:${np}`, u);
        continue;
      }
      if (a.namespaceURI) {
        let value = a.value;
        if (a.namespaceURI === MC && ['Ignorable', 'MustUnderstand'].includes(a.localName)) {
          value = value.split(/\s+/).filter(Boolean).map((p: string) => newPrefixOf(scopeUri(node, p) ?? '', p)).join(' ');
        }
        el.setAttributeNS(a.namespaceURI, `${newPrefixOf(a.namespaceURI, a.prefix ?? '')}:${a.localName}`, value);
      } else {
        let value = a.value;
        if (a.name === 'Requires' && node.namespaceURI === MC) {
          value = value.split(/\s+/).filter(Boolean).map((p: string) => newPrefixOf(scopeUri(node, p) ?? '', p)).join(' ');
        }
        el.setAttribute(a.name, value);
      }
    }
    for (let c = node.firstChild; c; c = c.nextSibling) {
      const k = copy(c);
      if (k) el.appendChild(k);
    }
    return el;
  };
  const root = copy(src.documentElement);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(root);
}

const CORPUS = join(__dirname, '..', '..', '..', 'corpus', 'vendor', 'astra', 'documents');
const FIXTURES = join(__dirname, '..', 'fixtures');
const WORD_FILES: string[] = [];
for (const dir of [FIXTURES, join(FIXTURES, 'word-revisions'), CORPUS]) {
  if (!existsSync(dir)) continue;
  const names = readdirSync(dir).filter((n) => n.endsWith('.docx') && !n.startsWith('encrypted')).sort();
  const step = dir === CORPUS ? Math.max(1, Math.floor(names.length / 40)) : 1; // a deterministic sample of the corpus
  names.forEach((n, i) => { if (i % step === 0) WORD_FILES.push(join(dir, n)); });
}

describe('real documents: independent renamer, then canonicalise, then compare with the original', () => {
  it('has real documents to work with', () => expect(WORD_FILES.length).toBeGreaterThan(10));

  it('every part of every sampled document: unchanged when canonical; a renamed twin canonicalises to the original meaning', () => {
    let parts = 0;
    let canonicalUntouched = 0;
    let renamedTwins = 0;
    const problems: string[] = [];
    for (const file of WORD_FILES) {
      let zip: AdmZip;
      try { zip = new AdmZip(readFileSync(file)); } catch { continue; }
      for (const entry of zip.getEntries()) {
        if (entry.isDirectory || !/\.xml$/i.test(entry.entryName) || entry.entryName.startsWith('[') || /^xl\//.test(entry.entryName)) continue;
        let text: string;
        try { text = entry.getData().toString('utf8'); } catch { continue; }
        if (text.length > 2_000_000) continue;
        parts++;
        // 1. a canonical part is never rewritten
        const direct = canonicalisePrefixes(text, entry.entryName);
        if (direct.xml === text && direct.renames.length === 0) canonicalUntouched++;
        else problems.push(`${file}:${entry.entryName} was rewritten although canonical (${JSON.stringify(direct.renames)})`);
        // 2. rename with the independent tool and canonicalise
        let twin: string;
        try { twin = renameWithDom(text, '9'); } catch { continue; }
        if (scanPart(twin) === null) continue;
        renamedTwins++;
        try {
          const out = canonicalisePart(twin, entry.entryName).xml;
          // original vs result: every node equal after resolution, every original binding still in scope
          const same = verifyEquivalent(text, out);
          if (!same.ok) problems.push(`${file}:${entry.entryName}: ${same.problem}`);
        } catch (e) {
          problems.push(`${file}:${entry.entryName}: refused ${(e as Error).message.slice(0, 140)}`);
        }
      }
    }
    // eslint-disable-next-line no-console
    console.log(`SPIKE parts=${parts} canonicalUntouched=${canonicalUntouched} renamedTwins=${renamedTwins} problems=${problems.length}`);
    for (const p of problems.slice(0, 15)) console.log('SPIKE problem', p);
    expect(problems).toEqual([]);
  }, 300000);
});

describe('real documents opened whole: renamed twin gives the same document as the original', () => {
  it('readDocumentXml and structural summary are equal for every sampled document', () => {
    let compared = 0;
    const problems: string[] = [];
    for (const file of WORD_FILES.slice(0, 80)) {
      let buf: Buffer;
      try { buf = readFileSync(file); new AdmZip(buf); } catch { continue; }
      let original: DocxParser;
      try { original = new DocxParser(buf); } catch { continue; } // encrypted, macro, invalid: not this test's business
      const zip = new AdmZip(buf);
      for (const e of zip.getEntries()) {
        if (e.isDirectory || !/^(word|customXml)\/.*\.xml$/.test(e.entryName)) continue;
        try { zip.updateFile(e.entryName, Buffer.from(renameWithDom(e.getData().toString('utf8'), '9'), 'utf8')); } catch { /* leave the part as it is */ }
      }
      let twin: DocxParser;
      try { twin = new DocxParser(zip.toBuffer()); } catch (err) { problems.push(`${file}: twin refused: ${(err as Error).message.slice(0, 120)}`); continue; }
      compared++;
      const a = original.getStructuralSummary();
      const b = twin.getStructuralSummary();
      if (a.nodes.length !== b.nodes.length) problems.push(`${file}: ${a.nodes.length} nodes became ${b.nodes.length}`);
      else if (JSON.stringify(a.nodes.map((n) => n.text)) !== JSON.stringify(b.nodes.map((n) => n.text))) problems.push(`${file}: node text differs`);
    }
    // eslint-disable-next-line no-console
    console.log(`SPIKEWHOLE compared=${compared} problems=${problems.length}`);
    for (const p of problems.slice(0, 10)) console.log('SPIKEWHOLE problem', p);
    expect(compared).toBeGreaterThan(10);
    expect(problems).toEqual([]);
  }, 300000);
});

describe('review fixes (second self-review)', () => {
  it('a default-namespace rewrite is idempotent even when a child sits in another default namespace', () => {
    const xml = `<root xmlns="${W}"><child xmlns="urn:other"><leaf/></child></root>`;
    const once = canonicalisePart(xml, 'p.xml');
    expect(once.renames.length).toBe(1);
    expect(scanPart(once.xml)).toBeNull();
    expect(canonicalisePrefixes(once.xml).renames).toEqual([]);
  });
  it('a declaration that exists only inside a comment neither triggers nor reports a rewrite', () => {
    const xml = `<r xmlns:x="urn:other"><!-- xmlns:x="${W}" --><x:p/></r>`;
    expect(scanPart(xml)).toBeNull();
    const out = ingestPrefixesFor(xml);
    expect(out.changed.size).toBe(0);
    expect(out.report).toEqual([]);
  });
  it('a name outside ASCII does not make the part unreadable', () => {
    const xml = `<x:document xmlns:x="${W}"><é/><x:p x:val="1"/></x:document>`;
    const out = canonicalisePart(xml, 'u.xml');
    expect(out.xml).toContain('<é/>');
    expect(out.xml).toContain('<w:p w:val="1"/>');
  });
  it('two attributes that would both become w:val are refused as a conflict with the standard text', () => {
    const xml = `<x:document xmlns:x="${W}" xmlns:y="${W}"><x:p x:val="1" y:val="2"/></x:document>`;
    const r = refusalOf(() => canonicalisePart(xml, 'd.xml'));
    expect(r.reason).toBe('conflict');
    expect(r.message).toMatch(/^E_VALIDATION: d\.xml cannot be canonicalised safely/);
    expect(r.message).toMatch(/Open the file in Word/);
  });
  it('mc:Ignorable that names a prefix declared nowhere is left as written, not refused', () => {
    const xml = `<x:document xmlns:x="${W}" xmlns:mc="${MC}" mc:Ignorable="w14 wp14"><x:body/></x:document>`;
    const out = canonicalisePart(xml, 'i.xml');
    expect(out.xml).toContain('mc:Ignorable="w14 wp14"');
  });
  it('a > inside an attribute value does not hide a tag from the scan', () => {
    expect(scanPart(`<x:document xmlns:x="${W}" title="a>b"><x:p/></x:document>`)).not.toBeNull();
  });
});

function ingestPrefixesFor(xml: string) {
  return ingestPrefixes([['p.xml', xml]]);
}
