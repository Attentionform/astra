import { createWordElement } from '../ooxml/ns';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST } from '../ast';
import { DocxParser, DocxZipAdapter, SECT_PR_CHILD_ORDER, upsertSectPrChild } from '../parser';
import { RelationshipsManipulator } from './relationships';
import { getOOXMLSnippet } from '../ooxml/guideline-bank';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { toError } from '../errors';

const HEADER_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const FOOTER_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const HEADER_PART_RE = /(?:^|\/)header\d+\.xml$/;
const FOOTER_PART_RE = /(?:^|\/)footer\d+\.xml$/;

function assertHeaderFooterTarget(
  targetType: string,
  pageType: string
): asserts targetType is 'header' | 'footer' {
  // BX-10 (manipulator scope): unknown targets/page types must fail loudly.
  // The old ternaries silently coerced ANY unknown targetType to 'footer'.
  if (targetType !== 'header' && targetType !== 'footer') {
    throw new Error(
      `Unknown header/footer targetType '${String(targetType)}'. Expected 'header' or 'footer'.`
    );
  }
  if (pageType !== 'default' && pageType !== 'first' && pageType !== 'even') {
    throw new Error(
      `Unknown header/footer pageType '${String(pageType)}'. Expected 'default', 'first' or 'even'.`
    );
  }
}

export interface SetHeaderFooterOptions {
  targetType: 'header' | 'footer';
  pageType?: 'default' | 'first' | 'even';
  text?: string;
  xmlContent?: string;
  sectionIndex?: number;
  alignment?: 'left' | 'center' | 'right';
  numberFormat?: string;
}

export interface RemoveHeaderFooterOptions {
  targetType: 'header' | 'footer';
  pageType?: 'default' | 'first' | 'even';
  sectionIndex?: number;
}

export class HeaderFooterManipulator {
  private ast: DocumentAST;
  private parser?: DocxParser;
  private domParser: DOMParser;
  private serializer: XMLSerializer;

  constructor(ast: DocumentAST, parser?: DocxParser) {
    this.ast = ast;
    this.parser = parser;
    this.domParser = new DOMParser();
    this.serializer = new XMLSerializer();
  }

  /**
   * Sets or updates a header or footer for a section, managing relationship entries,
   * content types, AST maps, and section properties (<w:sectPr>).
   */
  public setHeaderFooter(options: SetHeaderFooterOptions): { rId: string; partPath: string } {
    if (options.xmlContent) assertNoXxe(options.xmlContent, 'header/footer xmlContent');
    const targetType = options.targetType;
    const pageType = options.pageType || 'default';
    assertHeaderFooterTarget(targetType, pageType);
    const sectionIdx = options.sectionIndex || 0;

    const relType = targetType === 'header' ? HEADER_REL_TYPE : FOOTER_REL_TYPE;

    const contentType = targetType === 'header'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
      : 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

    const prefix = targetType === 'header' ? 'H' : 'F';
    void prefix;
    const relsManipulator = new RelationshipsManipulator(this.ast);

    // BX-04: match by relationship type AND pageType. The old lookup matched on
    // the part target alone, so requesting 'first'/'even' reused the 'default'
    // part's rId and relabeled it — destroying the default header. Now a
    // pageType only reuses a part whose AST map entry already carries that
    // pageType; otherwise a fresh part + rId is allocated below.
    const targetMapForLookup = targetType === 'header' ? this.ast.headers : this.ast.footers;
    void targetMapForLookup;
    const partRe = targetType === 'header' ? HEADER_PART_RE : FOOTER_PART_RE;
    let rId = '';
    let partName = '';

    // Scope lookup strictly to THIS section's sectPr to prevent cross-section clobbering
    let thisSectionExistingRId: string | null = null;
    try {
      const sectPrs = this.parser?.getSectPrElements() ?? [];
      const thisSectPr = sectPrs.length > sectionIdx ? sectPrs[sectionIdx] : null;
      if (thisSectPr) {
        const refTagName = targetType === 'header' ? 'w:headerReference' : 'w:footerReference';
        const refs = Array.from(thisSectPr.getElementsByTagName(refTagName));
        for (const ref of refs) {
          if ((ref.getAttribute('w:type') || 'default') === pageType) {
            thisSectionExistingRId = ref.getAttribute('r:id') || ref.getAttribute('id');
            break;
          }
        }
      }
    } catch { /* ignore */ }

    if (thisSectionExistingRId && this.ast.relationships[thisSectionExistingRId]) {
      const rel = this.ast.relationships[thisSectionExistingRId];
      if (rel.type === relType && partRe.test(rel.target)) {
        rId = thisSectionExistingRId;
        partName = rel.target.replace(/^word\//, '');
      }
    }

    // Word field format switches keyed by OOXML ST_NumberFormat values.
    const FIELD_SWITCH: Record<string, string> = {
      decimal: 'Arabic', lowerroman: 'roman', upperroman: 'ROMAN',
      lowerletter: 'alphabetic', upperletter: 'ALPHABETIC'
    };

    // In-place update: when the target part already exists and the caller only
    // tweaks alignment/numberFormat, edit the existing part so field codes
    // (page numbers) and custom content survive instead of being wiped.
    const existingPartPath = `word/${partName}`;
    const parser = this.parser;
    const canEditInPlace =
      !!rId && options.text === undefined && options.xmlContent === undefined &&
      !!(options.alignment || options.numberFormat) &&
      !!parser && parser.getZipAdapter().hasFile(existingPartPath);

    if (canEditInPlace) {
      const zipAdapter = parser.getZipAdapter();
      const existingXml = zipAdapter.readText(existingPartPath);
      assertNoXxe(existingXml, existingPartPath);
      const partDom = this.domParser.parseFromString(existingXml, 'text/xml');

      if (options.alignment) {
        for (const p of Array.from(partDom.getElementsByTagName('w:p'))) {
          let pPr = p.getElementsByTagName('w:pPr')[0];
          if (!pPr) {
            pPr = createWordElement(partDom, 'w:pPr');
            p.insertBefore(pPr, p.firstChild);
          }
          let jc = pPr.getElementsByTagName('w:jc')[0];
          if (!jc) {
            jc = createWordElement(partDom, 'w:jc');
            pPr.appendChild(jc);
          }
          jc.setAttribute('w:val', options.alignment);
        }
      }

      if (options.numberFormat) {
        const switchName = FIELD_SWITCH[options.numberFormat.toLowerCase()];
        if (switchName) {
          let patched = false;
          for (const instr of Array.from(partDom.getElementsByTagName('w:instrText'))) {
            const raw = instr.textContent || '';
            const field = /NUMPAGES/i.test(raw) ? 'NUMPAGES' : /PAGE/i.test(raw) ? 'PAGE' : null;
            if (field) {
              instr.textContent = ` ${field} \\* ${switchName} `;
              patched = true;
            }
          }
          if (!patched) {
            const root = partDom.documentElement;
            const fld = createWordElement(partDom, 'w:fldSimple');
            fld.setAttribute('w:instr', ` PAGE \\* ${switchName} `);
            root.appendChild(fld);
          }
        }
      }

      zipAdapter.writeText(existingPartPath, new XMLSerializer().serializeToString(partDom));
      const targetMap = targetType === 'header' ? this.ast.headers : this.ast.footers;
      targetMap[rId] = { type: pageType, xml: partDom.documentElement };
      this.updateSectPrReference(sectionIdx, targetType, pageType, rId);
      return { rId, partPath: existingPartPath };
    }

    if (!rId) {
      // Relationships may exist only in the package (e.g. the numbering rel
      // written by a rezip the session AST never saw). Adopt them first so the
      // allocation below cannot hand out an Id the zip already uses.
      if (this.parser) this.adoptZipRelationships(this.parser.getZipAdapter());
      // Find highest index for part naming
      let maxNum = 1;
      for (const rel of Object.values(this.ast.relationships)) {
        if (rel.type === relType) {
          const match = rel.target.match(/(?:header|footer)(\d+)\.xml/);
          if (match) {
            const num = parseInt(match[1], 10);
            if (num >= maxNum) maxNum = num + 1;
          }
        }
      }
      partName = `${targetType}${maxNum}.xml`;
      rId = relsManipulator.addRelationship(relType, partName);
    }

    // Determine XML content to use
    let partXmlStr = options.xmlContent;
    if (!partXmlStr) {
      if (options.text) {
        const safeText = options.text
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&apos;');

        const jcVal = options.alignment || (targetType === 'header' ? 'right' : 'center');
        const switchName = options.numberFormat ? FIELD_SWITCH[options.numberFormat.toLowerCase()] : undefined;
        const numFormatStr = switchName ? ` \\* ${switchName}` : '';

        const fieldRun = (instr: string): string =>
          `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr}${numFormatStr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
        const hasTokens = /\{(PAGE|NUMPAGES)\}/.test(safeText);
        const isPageLabel = /^\s*page\s*(\d+\s*(of\s*\d+)?)?\s*$/i.test(options.text);
        const rootTag = targetType === 'header' ? 'w:hdr' : 'w:ftr';

        if (hasTokens) {
          const runs = safeText
            .split(/(\{PAGE\}|\{NUMPAGES\})/)
            .filter((part) => part !== '')
            .map((part) =>
              part === '{PAGE}' ? fieldRun('PAGE')
              : part === '{NUMPAGES}' ? fieldRun('NUMPAGES')
              : `<w:r><w:t xml:space="preserve">${part}</w:t></w:r>`)
            .join('');
          partXmlStr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${rootTag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:jc w:val="${jcVal}"/></w:pPr>
    ${runs}
  </w:p>
</${rootTag}>`;
        } else if (targetType === 'footer' && (isPageLabel || !!switchName)) {
          partXmlStr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:jc w:val="${jcVal}"/></w:pPr>
    <w:r><w:t xml:space="preserve">${isPageLabel ? 'Page' : safeText} </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE${numFormatStr} </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> of </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMPAGES${numFormatStr} </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>`;
        } else if (targetType === 'header') {
          partXmlStr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:jc w:val="${jcVal}"/></w:pPr>
    <w:r><w:t>${safeText}</w:t></w:r>
  </w:p>
</w:hdr>`;
        } else {
          partXmlStr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:jc w:val="${jcVal}"/></w:pPr>
    <w:r><w:t>${safeText}</w:t></w:r>
  </w:p>
</w:ftr>`;
        }
      } else if (this.parser && this.parser.getZipAdapter().hasFile(`word/${partName}`)) {
        partXmlStr = this.parser.getZipAdapter().readText(`word/${partName}`);
      } else {
        const snippetName = targetType === 'header' ? 'headerPart' : 'footerPart';
        partXmlStr = getOOXMLSnippet('headersFooters', snippetName) || (
          targetType === 'header'
            ? `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`
            : `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>`
        );
      }
    }

    // OP-18: when xmlContent is a bare content fragment (e.g. a watermark
    // <w:p>) rather than a full <w:hdr>/<w:ftr> root, append it into the
    // existing part instead of replacing the whole part — replacing wipes
    // existing header content and a bare <w:p> part is corrupt (no w:hdr root).
    const zipPartPath = `word/${partName}`;
    assertNoXxe(partXmlStr, zipPartPath);
    if (partXmlStr) {
      let fragmentRootTag: string | null = null;
      try {
        const probeDom = this.domParser.parseFromString(partXmlStr, 'text/xml');
        fragmentRootTag = probeDom.documentElement ? probeDom.documentElement.tagName : null;
      } catch { fragmentRootTag = null; }
      const expectedRoot = targetType === 'header' ? 'w:hdr' : 'w:ftr';
      if (fragmentRootTag && fragmentRootTag !== expectedRoot) {
        const fragDom = this.domParser.parseFromString(partXmlStr, 'text/xml');
        const fragRoot = fragDom.documentElement;
        let baseDom: Document | null = null;
        if (this.parser && this.parser.getZipAdapter().hasFile(zipPartPath)) {
          const baseXml = this.parser.getZipAdapter().readText(zipPartPath);
          assertNoXxe(baseXml, zipPartPath);
          baseDom = this.domParser.parseFromString(baseXml, 'text/xml');
        } else {
          const shell =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<${expectedRoot} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
            ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
            ` xmlns:v="urn:schemas-microsoft-com:vml"` +
            ` xmlns:o="urn:schemas-microsoft-com:office:office"` +
            ` xmlns:w10="urn:schemas-microsoft-com:office:word"></${expectedRoot}>`;
          baseDom = this.domParser.parseFromString(shell, 'text/xml');
        }
        if (fragRoot && baseDom && baseDom.documentElement) {
          // De-duplicate repeated watermark applications: drop a previous
          // watermark paragraph (identified by its shape id) so re-running
          // never stacks identical behind-text shapes nor wipes user content.
          const prevShapes = Array.from(baseDom.getElementsByTagName('v:shape'));
          for (const shape of prevShapes) {
            if (shape.getAttribute('id') === 'PowerPlusWaterMarkObject') {
              let ancestor: Node | null = shape;
              while (ancestor && ancestor !== baseDom.documentElement) {
                if (ancestor.nodeType === 1 && (ancestor as Element).tagName === 'w:p') {
                  ancestor.parentNode?.removeChild(ancestor);
                  break;
                }
                ancestor = ancestor.parentNode;
              }
            }
          }
          let imported: Element = fragRoot as unknown as Element;
          try {
            if (typeof baseDom.importNode === 'function') {
              imported = baseDom.importNode(fragRoot, true) as Element;
            }
          } catch { imported = fragRoot as unknown as Element; }
          baseDom.documentElement.appendChild(imported);
          partXmlStr = new XMLSerializer().serializeToString(baseDom);
        }
      }
    }

    // Parse XML part DOM
    const partDom = this.domParser.parseFromString(partXmlStr, 'text/xml');

    // Update AST map
    const targetMap = targetType === 'header' ? this.ast.headers : this.ast.footers;
    targetMap[rId] = {
      type: pageType,
      xml: partDom.documentElement
    };

    // Update Zip Adapter files if parser is provided
    if (this.parser) {
      const zipAdapter = this.parser.getZipAdapter();
      zipAdapter.writeText(zipPartPath, partXmlStr);

      // Ensure relationship in word/_rels/document.xml.rels
      this.ensureDocumentRelationship(zipAdapter, rId, relType, partName);

      // Ensure content type override in [Content_Types].xml
      this.ensureContentTypeOverride(zipAdapter, `/word/${partName}`, contentType);
    }

    // Update section <w:sectPr> in AST section DOM nodes
    this.updateSectPrReference(sectionIdx, targetType, pageType, rId);

    return { rId, partPath: `word/${partName}` };
  }

  /**
   * Removes a header or footer reference from section properties (<w:sectPr>).
   */
  public removeHeaderFooter(options: RemoveHeaderFooterOptions): void {
    const targetType = options.targetType;
    const pageType = options.pageType || 'default';
    assertHeaderFooterTarget(targetType, pageType);
    const sectionIdx = options.sectionIndex || 0;

    const section = this.ast.sections[sectionIdx];
    if (!section) return;

    const tagName = targetType === 'header' ? 'w:headerReference' : 'w:footerReference';

    // HDRFTR-02: scope strictly to this section's sectPr — never the whole document.
    let targetSectPr: Element | null = null;
    try {
      const sectPrs = this.parser?.getSectPrElements() ?? [];
      if (sectPrs.length > sectionIdx) {
        targetSectPr = sectPrs[sectionIdx];
      }
    } catch { targetSectPr = null; }
    if (!targetSectPr) {
      for (const node of section.nodes) {
        const el = node.xml as Element;
        if (!el || typeof el.getElementsByTagName !== 'function') continue;
        if (el.tagName === 'w:sectPr') { targetSectPr = el; break; }
        const local = el.getElementsByTagName('w:sectPr')[0] as Element | undefined;
        if (local) { targetSectPr = local; break; }
      }
    }
    if (!targetSectPr) {
      // Single-section fallback (no-parser AST): body-level w:sectPr is a
      // sibling, not inside section.nodes. Reuse when exactly 1 exists.
      try {
        const anchorEl = section.nodes[section.nodes.length - 1]?.xml as Element | undefined;
        const doc = anchorEl?.ownerDocument as Document | undefined;
        if (doc && typeof doc.getElementsByTagName === 'function') {
          const all = doc.getElementsByTagName('w:sectPr');
          if (all.length === 1) targetSectPr = all[0] as Element;
        }
      } catch { /* leave null */ }
    }
    if (!targetSectPr) return this.cannotAttach(sectionIdx, targetType);

    const refs = Array.from(targetSectPr.getElementsByTagName(tagName));
    for (const ref of refs) {
      const typeAttr = (ref as Element).getAttribute('w:type') || 'default';
      if (typeAttr === pageType) {
        const rId = (ref as Element).getAttribute('r:id');
        if (ref.parentNode) {
          ref.parentNode.removeChild(ref);
        }

        if (rId) {
          delete (targetType === 'header' ? this.ast.headers : this.ast.footers)[rId];
        }
      }
    }
  }

  /**
   * Copies relationships present in word/_rels/document.xml.rels but
   * missing from ast.relationships into the AST (never overwrites an AST entry).
   */
  private adoptZipRelationships(zipAdapter: DocxZipAdapter): void {
    const relsPath = 'word/_rels/document.xml.rels';
    const relsXml = zipAdapter.readTextOrNull(relsPath);
    if (!relsXml) return;
    assertNoXxe(relsXml, relsPath);
    const doc = this.domParser.parseFromString(relsXml, 'text/xml');
    for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
      const id = rel.getAttribute('Id');
      if (!id || this.ast.relationships[id]) continue;
      const mode = rel.getAttribute('TargetMode');
      this.ast.relationships[id] = {
        type: rel.getAttribute('Type') || '',
        target: rel.getAttribute('Target') || '',
        ...(mode ? { targetMode: mode as 'External' | 'Internal' } : {}),
      };
    }
  }

  /**
   * Helper to ensure relationship exists in word/_rels/document.xml.rels
   */
  private ensureDocumentRelationship(zipAdapter: DocxZipAdapter, rId: string, type: string, target: string) {
    const relsPath = 'word/_rels/document.xml.rels';
    let relsXml = zipAdapter.readTextOrNull(relsPath);
    if (relsXml) assertNoXxe(relsXml, relsPath);

    if (!relsXml) {
      relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${rId}" Type="${type}" Target="${target}"/>
</Relationships>`;
    } else {
      const doc = this.domParser.parseFromString(relsXml, 'text/xml');
      const relsRoot = doc.documentElement;
      let found = false;

      const existingRels = Array.from(relsRoot.getElementsByTagName('Relationship'));
      for (const rel of existingRels) {
        if (rel.getAttribute('Id') === rId) {
          rel.setAttribute('Type', type);
          rel.setAttribute('Target', target);
          found = true;
          break;
        }
      }

      if (!found) {
        const newRel = doc.createElement('Relationship');
        newRel.setAttribute('Id', rId);
        newRel.setAttribute('Type', type);
        newRel.setAttribute('Target', target);
        relsRoot.appendChild(newRel);
      }
      relsXml = this.serializer.serializeToString(doc);
    }

    zipAdapter.writeText(relsPath, relsXml);
  }

  /**
   * Helper to ensure content type override exists in [Content_Types].xml
   */
  private ensureContentTypeOverride(zipAdapter: DocxZipAdapter, partName: string, contentType: string) {
    const typesPath = '[Content_Types].xml';
    let typesXml = zipAdapter.readTextOrNull(typesPath);
    if (typesXml) assertNoXxe(typesXml, typesPath);

    if (!typesXml) {
      typesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="${partName}" ContentType="${contentType}"/>
</Types>`;
    } else {
      const doc = this.domParser.parseFromString(typesXml, 'text/xml');
      const typesRoot = doc.documentElement;
      let found = false;

      const overrides = Array.from(typesRoot.getElementsByTagName('Override'));
      for (const ov of overrides) {
        if (ov.getAttribute('PartName') === partName) {
          ov.setAttribute('ContentType', contentType);
          found = true;
          break;
        }
      }

      if (!found) {
        const newOv = doc.createElement('Override');
        newOv.setAttribute('PartName', partName);
        newOv.setAttribute('ContentType', contentType);
        typesRoot.appendChild(newOv);
      }
      typesXml = this.serializer.serializeToString(doc);
    }

    zipAdapter.writeText(typesPath, typesXml);
  }

  /**
   * A header or footer that cannot be attached to (or removed from) its section must not pass
   * silently: its part is already in the package, so the batch has to fail
   * (and roll back) rather than leave a part no section refers to. Without a
   * parser (a headless AST) the section properties may legitimately be out of
   * reach, so the old no-op stays.
   */
  private cannotAttach(sectionIdx: number, targetType: 'header' | 'footer'): void {
    if (!this.parser) return;
    throw toError(
      'E_VALIDATION',
      `Cannot update the ${targetType} of section ${sectionIdx + 1}: the document has no section properties (w:sectPr) for it.`,
      'Re-read the document with inspect_document and check the section count; if it looks wrong, save and reopen the file, then retry.',
    );
  }

  /**
   * Helper to insert or update <w:headerReference> or <w:footerReference> inside section's <w:sectPr>
   */
  private updateSectPrReference(sectionIdx: number, targetType: 'header' | 'footer', pageType: 'default' | 'first' | 'even', rId: string) {
    const section = this.ast.sections[sectionIdx];
    if (!section) return this.cannotAttach(sectionIdx, targetType);

    // Find the sectPr belonging to THIS section only. Never fall back to the
    // document-wide first sectPr (that hijacks Section 0).
    let sectPrElement: Element | null = null;
    let parentDoc: Document | null = null;

    // Prefer the live parser sectPr for this section index (scoped, no hijack).
    try {
      const sectPrs = this.parser?.getSectPrElements() ?? [];
      if (sectPrs.length > sectionIdx) {
        sectPrElement = sectPrs[sectionIdx];
        parentDoc = (sectPrElement.ownerDocument as unknown as Document) || null;
      }
    } catch { /* fall through to local search */ }

    if (!sectPrElement) {
      for (const node of section.nodes) {
        const el = node.xml as Element;
        if (!el || typeof el.getElementsByTagName !== 'function') continue;
        if (el.tagName === 'w:sectPr') {
          sectPrElement = el;
          parentDoc = el.ownerDocument;
          break;
        }
        const found = el.getElementsByTagName('w:sectPr')[0] as Element | undefined;
        if (found) {
          sectPrElement = found;
          parentDoc = el.ownerDocument;
          break;
        }
      }
    }

    if (!sectPrElement) {
      // Single-section fallback (no-parser AST): w:sectPr is a body-level
      // sibling (w:body>w:sectPr), not inside section.nodes. Reuse it when
      // exactly 1 exists; refuse when >1 (true multi-section, ambiguous).
      try {
        const anchorEl = section.nodes[section.nodes.length - 1]?.xml as Element | undefined;
        const doc = anchorEl?.ownerDocument as Document | undefined;
        if (doc && typeof doc.getElementsByTagName === 'function') {
          const all = doc.getElementsByTagName('w:sectPr');
          if (all.length === 1) {
            sectPrElement = all[0] as Element;
            parentDoc = (sectPrElement.ownerDocument || doc) as Document | null;
          }
        }
      } catch { /* fall through to scoped creation */ }
    }

    if (!sectPrElement) {
      // No local sectPr for this section — create one scoped to this section's
      // final paragraph (a non-final section break lives in its pPr), or at
      // body level when this is the final section. Never touch Section 0.
      // Fail (return) if this section has no usable anchor.
      const lastNode = section.nodes[section.nodes.length - 1];
      const lastEl = lastNode?.xml as Element | undefined;
      parentDoc = lastEl?.ownerDocument ?? null;
      if (!lastEl || !parentDoc) return this.cannotAttach(sectionIdx, targetType);
      const isLastSection = sectionIdx === this.ast.sections.length - 1;
      if (isLastSection) {
        const body = parentDoc.getElementsByTagName('w:body')[0];
        if (!body) return this.cannotAttach(sectionIdx, targetType);
        sectPrElement = createWordElement(parentDoc, 'w:sectPr');
        body.appendChild(sectPrElement);
      } else {
        if (lastEl.tagName !== 'w:p') return this.cannotAttach(sectionIdx, targetType);
        let pPr = lastEl.getElementsByTagName('w:pPr')[0] as Element | undefined;
        if (!pPr) {
          pPr = createWordElement(parentDoc, 'w:pPr');
          lastEl.insertBefore(pPr, lastEl.firstChild);
        }
        sectPrElement = createWordElement(parentDoc, 'w:sectPr');
        pPr.appendChild(sectPrElement);
      }
    }

    if (!sectPrElement || !parentDoc) return this.cannotAttach(sectionIdx, targetType);

    // Documents built without xmlns:r (e.g. minimal blank templates) would
    // otherwise serialize r:id attributes with an undeclared prefix — strict
    // schema validation rejects the whole package. Declare it on the root.
    try {
      const root = parentDoc.documentElement;
      if (root && !root.getAttribute('xmlns:r')) {
        root.setAttribute(
          'xmlns:r',
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        );
      }
    } catch { /* namespace hardening is best-effort — the ref write below still applies */ }

    const refTagName = targetType === 'header' ? 'w:headerReference' : 'w:footerReference';
    const refs = Array.from(sectPrElement.getElementsByTagName(refTagName));
    
    let targetRef: Element | null = null;
    for (const ref of refs) {
      if ((ref.getAttribute('w:type') || 'default') === pageType) {
        targetRef = ref;
        break;
      }
    }

    if (!targetRef) {
      // HDRFTR-01: schema-correct ordered insertion via upsertSectPrChild.
      // upsertSectPrChild returns the first same-tag child regardless of
      // w:type, so only use it when no sibling of this kind exists; otherwise
      // create an additional ref (e.g. default + first) and insert in order.
      if (sectPrElement.getElementsByTagName(refTagName).length === 0) {
        targetRef = upsertSectPrChild(sectPrElement, refTagName);
      } else {
        const created = createWordElement(parentDoc, refTagName);
        const orderOf = (tag: string) => {
          const idx = SECT_PR_CHILD_ORDER.indexOf(tag);
          return idx === -1 ? SECT_PR_CHILD_ORDER.length : idx;
        };
        const targetIdx = orderOf(refTagName);
        let inserted = false;
        for (let i = 0; i < sectPrElement.childNodes.length; i++) {
          const child = sectPrElement.childNodes[i];
          if (child.nodeType !== 1) continue;
          if (orderOf((child as Element).tagName) > targetIdx) {
            sectPrElement.insertBefore(created, child);
            inserted = true;
            break;
          }
        }
        if (!inserted) sectPrElement.appendChild(created);
        targetRef = created;
      }
      targetRef.setAttribute('w:type', pageType);
    }

    targetRef.setAttribute('r:id', rId);

    // If first page header/footer, ensure <w:titlePg/> tag is in sectPr
    if (pageType === 'first') {
      upsertSectPrChild(sectPrElement, 'w:titlePg');
    }
  }
}
