/**
 * Documents with a known section structure, for tests that must not depend on
 * ASTra's own writer. Each shape is built from plain XML so the file is exactly
 * what the test says: where the section breaks are, what page each section has,
 * and which odd constructs sit around them.
 *
 * A section ends at the paragraph whose w:pPr carries a w:sectPr; the last one
 * ends at the w:sectPr that is a direct child of w:body.
 */
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser } from '@attentionform/astra-engine';
import { listFileSections } from '../../packages/engine/src/parser/sections';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const LETTER = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const A4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const LANDSCAPE = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>';

export const sectPr = (inner: string, attrs = ''): string => `<w:sectPr${attrs}>${inner}</w:sectPr>`;

/** A paragraph; `section` makes it the last paragraph of a section with those properties. */
export const para = (text: string, section?: string): string =>
  `<w:p>${section ? `<w:pPr>${section}</w:pPr>` : ''}${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;

const table = (cell: string): string =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr></w:tbl>`;

export interface HeaderPart {
  /** File name under word/, for example header1.xml. */
  name: string;
  /** Relationship id the section references. */
  id: string;
  text: string;
}

const headerXml = (text: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`;

/** A complete minimal package around `body`. */
export function docxFromBody(body: string, headers: HeaderPart[] = []): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        headers
          .map((h) => `<Override PartName="/word/${h.name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`)
          .join('') +
        `</Types>`,
      'utf8',
    ),
  );
  zip.addFile(
    '_rels/.rels',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
      'utf8',
    ),
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        headers.map((h) => `<Relationship Id="${h.id}" Type="${R_NS}/header" Target="${h.name}"/>`).join('') +
        `</Relationships>`,
      'utf8',
    ),
  );
  for (const h of headers) zip.addFile(`word/${h.name}`, Buffer.from(headerXml(h.text), 'utf8'));
  zip.addFile(
    'word/document.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`,
      'utf8',
    ),
  );
  // The zip format stores a time per entry, and adm-zip writes the machine's local clock reading, so a date built
  // in UTC gave different bytes in different time zones. Build the date from local parts: 2024-01-01 05:30:00 is
  // what the golden file was recorded with, and it is now the same bytes in every time zone.
  for (const entry of zip.getEntries()) entry.header.time = new Date(2024, 0, 1, 5, 30, 0);
  return zip.toBuffer();
}

const headerRef = (type: 'default' | 'first' | 'even', id: string): string => `<w:headerReference w:type="${type}" r:id="${id}"/>`;

export interface SectionShape {
  /** What the file's sections are, in order: how many paragraphs/blocks each holds. */
  sections: number;
  /** One line that says what is unusual about the shape. */
  about: string;
  build: () => Buffer;
}

/** The shapes the parser has to read correctly. Keys are stable; tests and goldens use them. */
export const SECTION_SHAPES: Record<string, SectionShape> = {
  single: {
    sections: 1,
    about: 'one section, body sectPr only',
    build: () => docxFromBody(para('Alpha') + para('Beta') + sectPr(LETTER)),
  },
  'two-sections': {
    sections: 2,
    about: 'a Letter section, then an A4 section',
    build: () =>
      docxFromBody(
        para('One A') + para('One B', sectPr(LETTER, '')) + para('Two A') + para('Two B') + sectPr(`<w:type w:val="nextPage"/>${A4}`),
      ),
  },
  'three-sections': {
    sections: 3,
    about: 'portrait, landscape, portrait',
    build: () =>
      docxFromBody(
        para('One A', sectPr(LETTER)) +
          para('Two A') +
          para('Two B', sectPr(`<w:type w:val="nextPage"/>${LANDSCAPE}`)) +
          para('Three A') +
          sectPr(`<w:type w:val="nextPage"/>${LETTER}`),
      ),
  },
  'trailing-empty-section': {
    sections: 2,
    about: 'the file ends with a section-break paragraph, so the last section holds no blocks',
    build: () => docxFromBody(para('One A') + para('One B', sectPr(LETTER)) + sectPr(`<w:type w:val="nextPage"/>${A4}`)),
  },
  'break-paragraph-first': {
    sections: 2,
    about: 'the very first paragraph is an empty section break, so section 1 is that one paragraph',
    build: () => docxFromBody(para('', sectPr(LETTER)) + para('Two A') + para('Two B') + sectPr(`<w:type w:val="nextPage"/>${A4}`)),
  },
  'table-before-break': {
    sections: 2,
    about: 'a table sits directly before the break paragraph',
    build: () =>
      docxFromBody(
        para('One A') + table(para('Cell')) + para('', sectPr(LETTER)) + para('Two A') + sectPr(`<w:type w:val="nextPage"/>${A4}`),
      ),
  },
  'break-in-content-control': {
    sections: 2,
    about: 'the break paragraph is inside a w:sdt content control',
    build: () =>
      docxFromBody(
        para('One A') +
          `<w:sdt><w:sdtPr><w:alias w:val="Wrapper"/></w:sdtPr><w:sdtContent>${para('One B', sectPr(LETTER))}</w:sdtContent></w:sdt>` +
          para('Two A') +
          sectPr(`<w:type w:val="nextPage"/>${A4}`),
      ),
  },
  'tracked-section-change': {
    sections: 2,
    about: 'a w:sectPrChange keeps the old properties inside a sectPr; it is not a section',
    build: () =>
      docxFromBody(
        para(
          'One A',
          sectPr(`${LETTER}<w:sectPrChange w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">${sectPr(A4)}</w:sectPrChange>`),
        ) +
          para('Two A') +
          sectPr(`<w:type w:val="nextPage"/>${A4}`),
      ),
  },
  'sectpr-in-table-cell': {
    sections: 1,
    about: 'a sectPr inside a paragraph in a table cell; Word ignores it, so the file still has one section',
    build: () => docxFromBody(para('One A') + table(para('Cell', sectPr(A4))) + para('One B') + sectPr(LETTER)),
  },
  'continuous-break': {
    sections: 2,
    about: 'the second section starts on the same page (continuous)',
    build: () => docxFromBody(para('One A', sectPr(LETTER)) + para('Two A') + sectPr(`<w:type w:val="continuous"/>${LETTER}`)),
  },
  'first-page-header': {
    sections: 2,
    about: 'section 2 has a different first-page header; section 1 only a default one',
    build: () =>
      docxFromBody(
        para('One A', sectPr(`${headerRef('default', 'rIdH1')}${LETTER}`)) +
          para('Two A') +
          sectPr(`${headerRef('first', 'rIdH2')}<w:type w:val="nextPage"/>${LETTER}<w:titlePg/>`),
        [
          { name: 'header1.xml', id: 'rIdH1', text: 'Default header' },
          { name: 'header2.xml', id: 'rIdH2', text: 'First page header' },
        ],
      ),
  },
  'landscape-section': {
    sections: 2,
    about: 'section 2 is landscape',
    build: () => docxFromBody(para('One A', sectPr(LETTER)) + para('Two A') + sectPr(`<w:type w:val="nextPage"/>${LANDSCAPE}`)),
  },
  'headers-per-section': {
    sections: 2,
    about: 'each section has its own default header',
    build: () =>
      docxFromBody(
        para('One A', sectPr(`${headerRef('default', 'rIdH1')}${LETTER}`)) +
          para('Two A') +
          sectPr(`${headerRef('default', 'rIdH2')}<w:type w:val="nextPage"/>${LETTER}`),
        [
          { name: 'header1.xml', id: 'rIdH1', text: 'Header of section one' },
          { name: 'header2.xml', id: 'rIdH2', text: 'Header of section two' },
        ],
      ),
  },
};

/**
 * Sections as the file defines them: a paragraph with a w:sectPr directly in its
 * w:pPr ends one, and the w:sectPr that is a child of w:body ends the last. A
 * w:sectPr inside a table cell or inside a w:sectPrChange does not count.
 */
export function countFileSections(docx: Buffer): number {
  const xml = new AdmZip(docx).readAsText('word/document.xml');
  const doc = new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;
  let count = 0;
  for (const sectPr of Array.from(doc.getElementsByTagName('w:sectPr'))) {
    const parent = sectPr.parentNode as Element;
    if (parent.tagName === 'w:body') {
      count++;
    } else if (parent.tagName === 'w:pPr') {
      let inCell = false;
      for (let el = parent.parentNode as Element | null; el && el.nodeType === 1; el = el.parentNode as Element | null) {
        if (el.tagName === 'w:tc') inCell = true;
      }
      if (!inCell) count++;
    }
  }
  return count;
}

/**
 * The text of every body-level block, grouped by the section the file puts it
 * in, read straight from the XML without the parser. A break paragraph is the
 * last block of the section it ends. Blocks in a w:sdt content control count;
 * blocks in table cells do not.
 */
export function fileBlocksBySection(docx: Buffer): string[][] {
  return fileSectionElements(docx).map((blocks) => blocks.map(blockSignature));
}

/** What identifies a block when it has no text: its kind and whether it carries a drawing, a chart or a field. */
export function blockSignature(el: Element): string {
  const marks = [
    el.getElementsByTagName('w:drawing').length > 0 ? 'drawing' : '',
    el.getElementsByTagName('c:chart').length > 0 ? 'chart' : '',
    el.getElementsByTagName('w:fldChar').length + el.getElementsByTagName('w:fldSimple').length > 0 ? 'field' : '',
  ].filter(Boolean);
  return `${el.tagName}:${marks.join('+')}:${el.textContent ?? ''}`;
}

/** The body-level block elements of each section, in the order the file has them. */
export function fileSectionElements(docx: Buffer): Element[][] {
  const xml = new AdmZip(docx).readAsText('word/document.xml');
  const doc = new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;
  const body = doc.getElementsByTagName('w:body')[0] as unknown as Element;
  const sections: Element[][] = [[]];
  const visit = (container: Element): void => {
    for (let node = container.firstChild; node; node = node.nextSibling) {
      if (node.nodeType !== 1) continue;
      const el = node as Element;
      if (el.tagName === 'w:p' || el.tagName === 'w:tbl') {
        sections[sections.length - 1].push(el);
        if (el.tagName === 'w:p') {
          const pPr = Array.from(el.childNodes as unknown as Element[]).find((c) => c.tagName === 'w:pPr');
          const breaks = pPr && Array.from(pPr.childNodes as unknown as Element[]).some((c) => c.tagName === 'w:sectPr');
          if (breaks) sections.push([]);
        }
      } else if (el.tagName === 'w:sdt') {
        for (const content of Array.from(el.childNodes as unknown as Element[]).filter((c) => c.tagName === 'w:sdtContent')) visit(content);
      }
    }
  };
  visit(body);
  return sections;
}

/** The section properties of each section of a saved file, in order. */
export const sectPrsOf = (docx: Buffer): Element[] => {
  const parser = new DocxParser(docx);
  const body = parser.getDocumentDom().getElementsByTagName('w:body')[0] as Element;
  return listFileSections(body).map((fileSection) => fileSection.sectPr as Element);
};
export const refsIn = (sectPr: Element, tag: 'w:headerReference' | 'w:footerReference'): string[] =>
  Array.from(sectPr.getElementsByTagName(tag)).map((ref) => (ref as Element).getAttribute('r:id') ?? '');

/** The text of the part a relationship id points to. */
export const partText = (docx: Buffer, rId: string): string => {
  const zip = new AdmZip(docx);
  const rels = zip.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
  const match = new RegExp(`<Relationship [^>]*Id="${rId}"[^>]*>`).exec(rels);
  const target = /Target="([^"]+)"/.exec(match?.[0] ?? '')?.[1];
  return target ? (zip.getEntry(`word/${target}`)?.getData().toString('utf8') ?? '') : '';
};

/** Header and footer parts in the package that no section refers to. */
export const orphanParts = (docx: Buffer): string[] => {
  const zip = new AdmZip(docx);
  const referenced = new Set<string>();
  const rels = zip.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
  const ids = sectPrsOf(docx).flatMap((sp) => [...refsIn(sp, 'w:headerReference'), ...refsIn(sp, 'w:footerReference')]);
  for (const id of ids) {
    const match = new RegExp(`<Relationship [^>]*Id="${id}"[^>]*>`).exec(rels);
    const target = /Target="([^"]+)"/.exec(match?.[0] ?? '')?.[1];
    if (target) referenced.add(`word/${target}`);
  }
  return zip
    .getEntries()
    .map((e) => e.entryName)
    .filter((n) => /^word\/(header|footer)\d*\.xml$/.test(n) && !referenced.has(n));
};
